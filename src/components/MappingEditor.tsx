import { Fragment, useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { FT, IMPORTABLE_TYPES, fieldTypeLabel, fieldTypeTone, matchFieldsByName } from '../lib/field-meta'
import type { FieldBrief, SourceColumn, SourceSheet, TableBrief } from '../lib/types'
import { IconImage, IconLink, IconTable } from './icons'
import { Popover, Tip } from './ui'

type Props = {
  sheets: SourceSheet[]
  tables: TableBrief[]
  onChange: (sheets: SourceSheet[]) => void
  onLoadFields: (tableId: string) => Promise<FieldBrief[]>
}

function cloneSheets(sheets: SourceSheet[]): SourceSheet[] {
  return sheets.map((s) => ({ ...s, columns: s.columns.map((c) => ({ ...c })) }))
}

/**
 * 字段名：单行显示，超出容器宽度时自动来回滚动。
 *
 * 原表里的列名常常长短悬殊（「工号」和「云南贝泰妮生物科技集团采购部名称全称」），
 * 直接铺开会把每行的类型列挤到不同位置，看着就错位。这里固定成单行 + 溢出滚动：
 * 既保证每列对齐，又不会因为截断而看不到全名。
 * 默认是只读按钮，点一下才切成 input 编辑。
 */
function FieldName({
  text,
  title,
  onClick,
}: {
  text: string
  title: string
  onClick: () => void
}) {
  const boxRef = useRef<HTMLButtonElement | null>(null)
  const innerRef = useRef<HTMLSpanElement | null>(null)
  const [overflow, setOverflow] = useState(false)

  useLayoutEffect(() => {
    const box = boxRef.current
    const inner = innerRef.current
    if (!box || !inner) return
    // clientWidth 为 0 说明还没布局（例如测试环境），此时不做判断
    const over = inner.scrollWidth - box.clientWidth
    if (box.clientWidth > 0 && over > 4) {
      setOverflow(true)
      inner.style.setProperty('--fld-shift', `-${over + 2}px`)
    } else {
      setOverflow(false)
      inner.style.removeProperty('--fld-shift')
    }
  }, [text])

  return (
    <button
      type="button"
      ref={boxRef}
      className={overflow ? 'fld-name overflow' : 'fld-name'}
      title={title}
      onClick={onClick}
    >
      <span className="fld-name-inner" ref={innerRef}>
        {text}
      </span>
    </button>
  )
}

export default function MappingEditor({ sheets, tables, onChange, onLoadFields }: Props) {
  const [fieldCache, setFieldCache] = useState<Record<string, FieldBrief[]>>({})
  const inflight = useRef<Set<string>>(new Set())
  const [collapsed, setCollapsed] = useState<Record<number, boolean>>({})
  /** 正在编辑字段名的列 key（null = 全部处于只读展示态） */
  const [editingKey, setEditingKey] = useState<string | null>(null)

  // 追加模式下按需加载目标表的字段清单
  useEffect(() => {
    for (const s of sheets) {
      if (s.importMode !== 'append' || !s.importTableId) continue
      if (fieldCache[s.importTableId] || inflight.current.has(s.importTableId)) continue
      inflight.current.add(s.importTableId)
      onLoadFields(s.importTableId)
        .then((list) => setFieldCache((p) => ({ ...p, [s.importTableId]: list })))
        .catch(() => setFieldCache((p) => ({ ...p, [s.importTableId]: [] })))
        .finally(() => inflight.current.delete(s.importTableId))
    }
  }, [sheets, fieldCache, onLoadFields])

  const mutate = useCallback(
    (fn: (draft: SourceSheet[]) => void) => {
      const draft = cloneSheets(sheets)
      fn(draft)
      onChange(draft)
    },
    [sheets, onChange],
  )

  /*
   * 合并到已有表时，按**字段名自动匹配**。
   *
   * 规则（与「写入新表」共用同一套字段行，但语义不同）：
   *   A 合并进 B，A 的「甲/乙/丙」自动写进 B 的同名字段；
   *   B 里没有的（如「丁」）保持未绑定 —— 用户勾选即新增该列，取消勾选即不导入。
   *
   * 用户手动调过的列（mappingTouched）不参与，否则刚解绑就又被自动绑回去。
   * 只在「确实有列能匹配上」时才改数据，避免 onChange 触发无限循环。
   */
  useEffect(() => {
    for (const s of sheets) {
      if (s.importMode !== 'append' || !s.importTableId) continue
      const list = fieldCache[s.importTableId]
      if (!list) continue
      const hits = matchFieldsByName(s.columns, list)
      // 只在「确实有列能匹配上且还没绑」时才改数据，避免 onChange 触发无限循环
      const need = hits.some((h, i) => !!h && !s.columns[i].targetFieldId)
      if (!need) continue
      mutate((draft) => {
        const sheet = draft.find((x) => x.name === s.name)
        if (!sheet) return
        const mapped = matchFieldsByName(sheet.columns, list)
        sheet.columns.forEach((col, i) => {
          if (col.targetFieldId) return
          const hit = mapped[i]
          if (!hit) return
          col.targetFieldId = hit.id
          col.targetFieldName = hit.name
          col.targetFieldType = hit.type
        })
      })
    }
  }, [sheets, fieldCache, mutate])

  const updateColumn = (sheetName: string, colKey: string, patch: Partial<SourceColumn>) => {
    mutate((draft) => {
      const sheet = draft.find((s) => s.name === sheetName)
      if (!sheet) return
      const col = sheet.columns.find((c) => c.key === colKey)
      if (!col) return
      Object.assign(col, patch)
    })
  }

  if (!sheets.length) return <div className="empty">没有解析到任何工作表</div>

  return (
    <div className="map-list">
      {sheets.map((sheet, si) => {
        const fields = sheet.importTableId ? fieldCache[sheet.importTableId] ?? [] : []
        const appendMode = sheet.importMode === 'append' && !!sheet.importTableId
        const enabledCount = sheet.columns.filter((c) => c.enabled).length
        const isCollapsed = !!collapsed[si]
        /** 本表是否全选 / 部分选中（表头的三态勾选框要用） */
        const allOn = sheet.columns.length > 0 && enabledCount === sheet.columns.length
        const someOn = enabledCount > 0 && !allOn
        /*
         * 合并到已有表时，先把「自动匹配」的结果汇总出来 ——
         * 用户最关心的就是「哪些字段接上了、哪些会新建」，不必逐行去看。
         */
        const matchedCount = appendMode
          ? sheet.columns.filter((c) => c.enabled && c.targetFieldId).length
          : 0
        const newCount = appendMode
          ? sheet.columns.filter((c) => c.enabled && !c.targetFieldId).length
          : 0

        return (
          <div className="sheet-block" key={`sheet-${si}`}>
            {/*
              头部：工作表名 + 统计 + 目标表。
              目标表做成「下拉 + 名称」的组合控件（.dest-combo）——
              选「新建数据表」时右侧直接就是可编辑的名称，省掉一整行和两个标签。
            */}
            <div className="sheet-head">
              <button
                type="button"
                className="sheet-toggle"
                aria-expanded={!isCollapsed}
                onClick={() => setCollapsed((p) => ({ ...p, [si]: !p[si] }))}
              >
                <span className={isCollapsed ? 'caret' : 'caret open'} />
                <span className="sheet-name">{sheet.name}</span>
              </button>
              <span className="sheet-meta">
                {sheet.totalDataRows} 行 · {enabledCount}/{sheet.columns.length} 列
              </span>

              {/* 合并模式下把「匹配 / 新增」的结论摆出来，不用逐行看 */}
              {appendMode && (matchedCount > 0 || newCount > 0) && (
                <span className="sheet-merge-stat" title="按字段名自动匹配的结果">
                  {matchedCount > 0 && <span className="ms-item ms-ok">✓ {matchedCount} 个接上</span>}
                  {newCount > 0 && <span className="ms-item ms-new">＋ {newCount} 个新增</span>}
                </span>
              )}

              <div className="dest-combo">
                <select
                  className="dest-select"
                  aria-label="导入到"
                  title="选择导入到新建数据表，还是追加到已有数据表"
                  value={sheet.importMode === 'append' ? sheet.importTableId : '__new__'}
                  onChange={(e) => {
                    const v = e.target.value
                    mutate((draft) => {
                      const s = draft[si]
                      if (v === '__new__') {
                        s.importMode = 'create'
                        s.importTableId = ''
                        // 回到「新建表」模式：清掉绑定，所有字段都变成"将新建"
                        s.columns.forEach((c) => {
                          c.targetFieldId = ''
                          c.mappingTouched = false
                        })
                      } else {
                        s.importMode = 'append'
                        s.importTableId = v
                        // 换了目标表：清空绑定并允许重新自动匹配
                        s.columns.forEach((c) => {
                          c.targetFieldId = ''
                          c.mappingTouched = false
                        })
                      }
                    })
                  }}
                >
                  <option value="__new__">新建数据表</option>
                  {tables.map((t) => (
                    <option key={t.id} value={t.id}>
                      追加到：{t.name}
                    </option>
                  ))}
                </select>
                {sheet.importMode === 'create' && (
                  <input
                    className="dest-name"
                    value={sheet.importTableName}
                    placeholder="表名称"
                    title="新数据表的名称"
                    onChange={(e) => {
                      const v = e.target.value
                      mutate((draft) => {
                        draft[si].importTableName = v
                      })
                    }}
                  />
                )}
              </div>
            </div>

            {/* ---------- 字段映射 ---------- */}
            {!isCollapsed && (
              <>
                {/* 表头：与 .map-row 共用同一套 grid 模板，保证各列对齐 */}
                <div className="map-head">
                  {/* 第一格是「本表全选」——放在字段名前面，作用范围限于当前工作表 */}
                  <label className="mh-check" title={allOn ? '取消本表全部字段' : '选中本表全部字段'}>
                    <input
                      type="checkbox"
                      checked={allOn}
                      ref={(el) => {
                        // 部分选中显示为「不确定」态
                        if (el) el.indeterminate = !allOn && someOn
                      }}
                      onChange={() =>
                        mutate((draft) => {
                          const s = draft[si]
                          s.columns.forEach((c) => (c.enabled = !allOn))
                        })
                      }
                    />
                  </label>
                  <span>字段名</span>
                  <span className="mh-count">非空</span>
                  <span />
                  {/* 两套逻辑的语义不同：新建表时选类型；合并到已有表时看"写进哪个字段" */}
                  <span>{appendMode ? '写入到' : '字段类型'}</span>
                </div>
                <div className="map-rows">
                {sheet.columns.length === 0 && <div className="empty sm">该工作表没有可导入的列</div>}
                {sheet.columns.map((col) => (
                  <div className={col.enabled ? 'map-row' : 'map-row off'} key={col.key}>
                    <label className="map-check" title={col.enabled ? '取消该列的导入' : '启用该列的导入'}>
                      <input
                        type="checkbox"
                        checked={col.enabled}
                        onChange={(e) => updateColumn(sheet.name, col.key, { enabled: e.target.checked })}
                      />
                    </label>

                    {/* 字段名称：只读展示（超长自动滚动），点击才进入编辑 */}
                    <div className="map-name">
                      {editingKey === col.key ? (
                        <input
                          className="input src-input"
                          autoFocus
                          value={col.header}
                          onChange={(e) => {
                            const v = e.target.value
                            updateColumn(sheet.name, col.key, {
                              header: v,
                              targetFieldName: col.targetFieldId ? col.targetFieldName : v,
                            })
                          }}
                          onBlur={() => setEditingKey(null)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter' || e.key === 'Escape') setEditingKey(null)
                          }}
                        />
                      ) : (
                        <FieldName
                          text={col.header}
                          title={`来源列 ${col.letter}：${col.header}（点击可改名）`}
                          onClick={() => setEditingKey(col.key)}
                        />
                      )}
                    </div>

                    {/*
                      非空值 / 图片数：单独一列，上方表头写着「非空」——
                      之前它跟在字段名后面、没有任何标注，用户根本猜不出这个数字是什么。
                    */}
                    <div className="map-count">
                      {col.mediaCount > 0 ? (
                        <span className="badge acc" title={`该列识别到 ${col.mediaCount} 个图片/附件`}>
                          <IconImage size={11} />
                          {col.mediaCount}
                        </span>
                      ) : col.valueCount > 0 ? (
                        <span className="badge subtle" title={`该列有 ${col.valueCount} 个非空单元格`}>
                          {col.valueCount}
                        </span>
                      ) : null}
                    </div>

                    <span className="map-arrow" aria-hidden>
                      →
                    </span>

                    {/*
                      右侧有两种语义 —— 这是「写入新表」与「合并到已有表」的关键差异：
                        · 新建数据表：这里选**字段类型**（所有字段都是新建的）
                        · 合并到已有表：这里显示**写进哪个字段**
                            已按名匹配上 → 显示目标字段名（绿）
                            目标表里没有 → 显示「+ 新增…」（会为新字段建列）
                      两种情况下点开都是同一个弹层，可以随时改。
                    */}
                    <div className="map-type">
                      <Popover
                        ariaLabel={appendMode ? '选择写入的目标字段' : '选择字段类型'}
                        trigger={({ open, toggle }) => {
                          if (!appendMode) {
                            return (
                              <button
                                type="button"
                                className={`tk tk-${fieldTypeTone(col.targetFieldType)} tk-btn${open ? ' open' : ''}`}
                                title={`导入后的字段类型：${fieldTypeLabel(col.targetFieldType)}，点击更换`}
                                aria-haspopup="dialog"
                                aria-expanded={open}
                                onClick={toggle}
                              >
                                <i className="tk-dot" />
                                <span className="tk-text">{fieldTypeLabel(col.targetFieldType)}</span>
                                <span className="tk-caret" aria-hidden />
                              </button>
                            )
                          }
                          // 合并模式：把匹配结果显示出来
                          return col.targetFieldId ? (
                            <button
                              type="button"
                              className={`tk tk-matched tk-btn${open ? ' open' : ''}`}
                              title={`已按字段名匹配到「${col.targetFieldName}」（${fieldTypeLabel(col.targetFieldType)}），点击更改`}
                              aria-haspopup="dialog"
                              aria-expanded={open}
                              onClick={toggle}
                            >
                              <IconLink size={11} />
                              <span className="tk-text">{col.targetFieldName}</span>
                              <span className="tk-caret" aria-hidden />
                            </button>
                          ) : (
                            <button
                              type="button"
                              className={`tk tk-newfield tk-btn${open ? ' open' : ''}`}
                              title={`目标表里没有「${col.header}」，导入时会新增该字段（${fieldTypeLabel(col.targetFieldType)}）；不想新增就取消左侧勾选`}
                              aria-haspopup="dialog"
                              aria-expanded={open}
                              onClick={toggle}
                            >
                              <span className="tk-text">＋ 新增「{col.header}」</span>
                              <span className="tk-caret" aria-hidden />
                            </button>
                          )
                        }}
                      >
                        {({ close }) => (
                          <span className="tk-menu" role="listbox">
                            {appendMode && fields.length > 0 && (
                              <>
                                <span className="tk-sec">
                                  <IconTable size={11} />
                                  写入到
                                </span>
                                <span className="tk-target">
                                  <select
                                    className="input"
                                    value={col.targetFieldId}
                                    onChange={(e) => {
                                      const v = e.target.value
                                      // 用户手动指定过 → 之后不再自动匹配
                                      if (!v) {
                                        updateColumn(sheet.name, col.key, {
                                          targetFieldId: '',
                                          targetFieldName: col.header,
                                          mappingTouched: true,
                                        })
                                        return
                                      }
                                      const meta = fields.find((f) => f.id === v)
                                      updateColumn(sheet.name, col.key, {
                                        targetFieldId: v,
                                        targetFieldName: meta?.name ?? col.targetFieldName,
                                        targetFieldType: meta ? meta.type : col.targetFieldType,
                                        mappingTouched: true,
                                      })
                                    }}
                                  >
                                    <option value="">＋ 新建字段「{col.header}」</option>
                                    {fields.map((f) => (
                                      <option key={f.id} value={f.id}>
                                        {f.name}（{fieldTypeLabel(f.type)}）
                                      </option>
                                    ))}
                                  </select>
                                </span>

                                {/*
                                 * 只有「新建字段」才显示名称和类型。
                                 *
                                 * 写入已有字段时，类型和目标字段名都由那个字段决定 ——
                                 * 在这里给出类型列表会让人以为能改，实际改的是
                                 * `targetFieldType`（只影响"新建时用什么类型"），
                                 * 写入时仍然按目标字段的类型走。这是误导。
                                 */}
                                {!col.targetFieldId && (
                                  <>
                                    <span className="tk-sec">新字段名称</span>
                                    <span className="tk-target">
                                      <input
                                        className="input"
                                        value={col.targetFieldName}
                                        placeholder={col.header}
                                        aria-label="新字段名称"
                                        onChange={(e) =>
                                          updateColumn(sheet.name, col.key, {
                                            targetFieldName: e.target.value,
                                            mappingTouched: true,
                                          })
                                        }
                                      />
                                    </span>
                                    <span className="tk-sec">新字段类型</span>
                                  </>
                                )}
                              </>
                            )}
                            {/* 新建表模式：没有目标表字段可绑，名称和类型都可以改 */}
                            {!appendMode && (
                              <>
                                <span className="tk-sec">字段名称</span>
                                <span className="tk-target">
                                  <input
                                    className="input"
                                    value={col.targetFieldName}
                                    placeholder={col.header}
                                    aria-label="字段名称"
                                    onChange={(e) =>
                                      updateColumn(sheet.name, col.key, {
                                        targetFieldName: e.target.value,
                                      })
                                    }
                                  />
                                </span>
                                <span className="tk-sec">字段类型</span>
                              </>
                            )}
                            {( !appendMode || !col.targetFieldId ) &&
                              IMPORTABLE_TYPES.map((t) => (
                                <button
                                  key={t.value}
                                  type="button"
                                  role="option"
                                  aria-selected={t.value === col.targetFieldType}
                                  className={`tk-item${t.value === col.targetFieldType ? ' on' : ''}`}
                                  onClick={() => {
                                    updateColumn(sheet.name, col.key, {
                                      targetFieldType: t.value,
                                      inferredType: t.value,
                                      typeTouched: true,
                                    })
                                    close()
                                  }}
                                >
                                  <span className={`tk tk-${fieldTypeTone(t.value)}`}>
                                    <i className="tk-dot" />
                                    {t.label}
                                  </span>
                                  {t.value === col.targetFieldType && <span className="tk-check">✓</span>}
                                </button>
                              ))}
                          </span>
                        )}
                      </Popover>
                      {col.targetFieldType === FT.Attachment && !col.typeTouched && col.mediaCount === 0 && (
                        <Tip text="该列被识别为附件字段，但没有检测到图片，导入后会得到空附件。">
                          <span className="dot warn" />
                        </Tip>
                      )}
                    </div>
                  </div>
                ))}
                </div>
              </>
            )}
          </div>
        )
      })}
    </div>
  )
}
