import { FT, extOf, findOptionId, hasOption, mimeOfName } from './field-meta'
import { formatDate, parseBool, parseDateMs, parseNumber, rawToText, splitMulti } from './infer'
import type { CellRaw, MediaRef } from './types'

export type ConvertResult = { ok: true; value: unknown } | { ok: false; reason: string }

export type ConvertContext = {
  /** 该单元格上传成功后的附件 */
  media: MediaRef[]
  /**
   * 单选/多选的「选项名 -> 选项 id」映射。
   *
   * 多维表格写单选/多选单元格时，值必须是 { id, text }：
   * 只给纯文本不会报错，但单元格会保持空白（选项建好了却「没有值」）。
   */
  options?: Map<string, string>
}

/**
 * 把 Excel 单元格的原始值转成多维表格能写入的形态。
 * 附件值形如 [{ name, size, type, token, timeStamp }]，与官方 setCellValue 示例一致。
 */
export function toBitableValue(type: number, raw: CellRaw, ctx: ConvertContext): ConvertResult {
  switch (type) {
    case FT.Attachment: {
      if (!ctx.media.length) return { ok: false, reason: '单元格内没有可用的附件' }
      return {
        ok: true,
        value: ctx.media.map((m) => ({
          name: m.name,
          size: m.size,
          type: m.type,
          token: m.token,
          timeStamp: Date.now(),
        })),
      }
    }

    case FT.Text:
      return { ok: true, value: rawToText(raw) }

    case FT.Number: {
      const n = parseNumber(raw)
      if (n === null) return { ok: false, reason: '不是合法数字' }
      return { ok: true, value: n }
    }

    case FT.Checkbox: {
      // 空单元格 / 无法识别的文本 → ok:false（不写），绝不当成 false 写进去
      const b = parseBool(raw)
      if (b === null) return { ok: false, reason: '空值' }
      return { ok: true, value: b }
    }

    case FT.DateTime: {
      const ms = parseDateMs(raw)
      if (ms === null) return { ok: false, reason: '不是合法日期' }
      return { ok: true, value: ms }
    }

    case FT.SingleSelect: {
      const t = rawToText(raw).trim()
      if (!t) return { ok: false, reason: '空值' }
      if (!ctx.options) {
        return { ok: false, reason: '读不到该单选字段的选项，写入会变成空白' }
      }
      const val = selectValue(t, ctx.options)
      if (!val) {
        const brief = t.length > 30 ? `${t.slice(0, 30)}…` : t
        return { ok: false, reason: `单选选项里没有「${brief}」（选项未建成，写入会变空白）` }
      }
      return { ok: true, value: val }
    }

    case FT.MultiSelect: {
      const arr = splitMulti(raw)
      if (!arr.length) return { ok: false, reason: '空值' }
      if (!ctx.options) {
        return { ok: false, reason: '读不到该多选字段的选项，写入会丢失' }
      }
      /*
       * 判断"选项是否存在"必须走 hasOption（三级比对），
       * 和单选分支保持一致。
       *
       * 早前这里直接 `ctx.options.has(t)` —— 只做精确匹配，
       * 于是「， 挤不出」这种只差一个空格的值会被判成"没有这个选项"，
       * 明明能写进去却报错跳过。单选那边修过，多选漏了。
       */
      const missing = arr.filter((t) => !hasOption(ctx.options as Map<string, string>, t))
      if (missing.length) {
        const first = missing[0]
        const brief = first.length > 30 ? `${first.slice(0, 30)}…` : first
        return {
          ok: false,
          reason: `多选选项里没有「${brief}」${missing.length > 1 ? ` 等 ${missing.length} 个值` : ''}（写入会丢失）`,
        }
      }
      const picked = arr
        .map((t) => selectValue(t, ctx.options))
        .filter((x): x is { id: string; text: string } => !!x)
      return { ok: true, value: picked }
    }

    case FT.Phone:
      return { ok: true, value: rawToText(raw).trim() }

    case FT.Email:
      return { ok: true, value: rawToText(raw).trim() }

    case FT.Url: {
      const t = rawToText(raw).trim()
      if (!t) return { ok: false, reason: '空值' }
      // UrlTransformVal = string | IOpenUrlSegment | IOpenUrlSegment[]
      // 传纯字符串最稳（IOpenUrlSegment 必须带 type:'url'，少字段会被拒）
      const link = /^[a-z][a-z0-9+.-]*:\/\//i.test(t) ? t : `https://${t}`
      return { ok: true, value: link }
    }

    case FT.Currency:
    case FT.Progress:
    case FT.Rating: {
      const n = parseNumber(raw)
      if (n === null) return { ok: false, reason: '不是合法数字' }
      return { ok: true, value: n }
    }

    default:
      return { ok: true, value: rawToText(raw) }
  }
}

/**
 * 单选/多选单元格的值。
 *
 * 返回 `null` 表示**无法给出合法值**（选项表没读到，或表里没有这个值）。
 *
 * ⚠️ 早前拿不到 id 时会「退回纯文本」，而飞书对「纯文本写单选」不报错、
 * 只是把单元格留空 —— 于是数据被静默吞掉，用户以为导入成功了。
 * 现在改为显式失败，让上层记进错误清单，看得见。
 */
function selectValue(
  text: string,
  options?: Map<string, string>,
): { id: string; text: string } | null {
  if (!options) return null
  /*
   * 走统一的 findOptionId（三级比对）。**不要在这里另写一套查找逻辑** ——
   * 写入用的规则和"判断有没有建成"用的规则必须完全一致，
   * 否则会出现"写成功了却报没建成"这种自相矛盾的提示。
   */
  const id = findOptionId(options, text)
  return id ? { id, text } : null
}

/** 反向：多维表格单元格 -> 写入 Excel 的文本/数字 */
export function bitableValueToExcel(type: number, v: unknown): string | number | boolean | null {
  if (v === null || v === undefined) return null

  if (type === FT.DateTime || type === FT.CreatedTime || type === FT.ModifiedTime) {
    const ms = typeof v === 'number' ? v : Number(v)
    if (!Number.isFinite(ms) || ms <= 0) return null
    /*
     * 返回 **Excel 序列号**（数字），而不是格式化好的字符串。
     *
     * 字符串在 Excel 里是「文本日期」：看着像日期，但左对齐、不能排序、
     * 不能筛选、也不能参与日期运算。转成序列号后由写入侧给这一列设日期格式，
     * Excel 才会把它当真正的日期。
     */
    return ms / 86_400_000 + 25_569
  }
  if (type === FT.Checkbox) return v === true || v === 'true' || v === 1

  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v
  if (v instanceof Date) return formatDate(v)

  if (Array.isArray(v)) {
    return v
      .map((item) => {
        if (item === null || item === undefined) return ''
        if (typeof item === 'string' || typeof item === 'number') return String(item)
        if (typeof item === 'object') {
          const o = item as Record<string, unknown>
          return String(o.text ?? o.name ?? o.en_name ?? o.link ?? '')
        }
        return String(item)
      })
      .filter(Boolean)
      .join(', ')
  }

  if (typeof v === 'object') {
    const o = v as Record<string, unknown>
    if ('text' in o && 'link' in o) return String(o.text ?? o.link ?? '')
    if ('link' in o) return String(o.link ?? '')
    if ('text' in o) return String(o.text ?? '')
    if ('name' in o) return String(o.name ?? '')
    if ('value' in o) return String(o.value ?? '')
    if ('full_address' in o || 'address' in o) return String(o.full_address ?? o.address ?? '')
  }
  return null
}

/** 从多维表格附件单元格里取出 token / 名称 / 类型 */
export function extractAttachments(v: unknown): { token: string; name: string; size: number; type: string }[] {
  if (!Array.isArray(v)) return []
  return v
    .map((item) => {
      if (!item || typeof item !== 'object') return null
      const o = item as Record<string, unknown>
      const token = String(o.token ?? o.file_token ?? '')
      if (!token) return null
      const name = String(o.name ?? `${token}.png`)
      return {
        token,
        name,
        size: Number(o.size ?? 0),
        type: String(o.type ?? mimeOfName(name, 'image/png')),
      }
    })
    .filter((x): x is { token: string; name: string; size: number; type: string } => !!x)
}

export function guessExtFromMime(mime: string, fallbackName?: string): string {
  const m = (mime || '').toLowerCase()
  if (m.includes('png')) return 'png'
  if (m.includes('jpeg') || m.includes('jpg')) return 'jpg'
  if (m.includes('gif')) return 'gif'
  if (m.includes('webp')) return 'webp'
  if (m.includes('bmp')) return 'bmp'
  if (m.includes('tiff')) return 'tiff'
  if (m.includes('svg')) return 'svg'
  if (fallbackName) {
    const e = extOf(fallbackName)
    if (e) return e
  }
  return 'png'
}
