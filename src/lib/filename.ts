/**
 * 文件名相关的纯函数 —— 供「仅导出附件」使用。
 *
 * 单独成文件的原因：
 * 1. 这些逻辑与飞书 SDK 无关，可以独立单测；
 * 2. `exporter.ts` 依赖 `@lark-base-open/js-sdk`，测试里必须打 stub 才能跑，
 *    而文件名规则恰恰是最需要断言的部分（一个 `:` 就能让 Windows 解压失败）。
 */

import { guessExtFromMime } from './value-convert'

/**
 * 把飞书的单元格值转成纯文本（用于「按字段值命名」取标题）。
 *
 * 不能直接用 `rawToText` —— 那个只接受 `string | number | boolean | Date | null`，
 * 而飞书的单选是 `{ id, text }`、人员是 `[{ id, name }]`、多选是数组，
 * 直接 String() 会得到 `[object Object]`，变成一堆同名文件。
 */
export function cellToPlainText(v: unknown): string {
  if (v === null || v === undefined) return ''
  if (typeof v === 'string') return v
  if (typeof v === 'number' || typeof v === 'boolean') return String(v)
  if (v instanceof Date) {
    // 不要 toISOString（会带时区和 T），和导出 Excel 的格式保持一致
    const p = (n: number) => String(n).padStart(2, '0')
    return `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`
  }
  if (Array.isArray(v)) {
    return v
      .map((x) => cellToPlainText(x))
      .filter((s) => s.trim())
      .join(' ')
  }
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>
    // 单选/多选 { id, text }、人员 { id, name }、链接 { link, text }
    const pick = o.text ?? o.name ?? o.en_name ?? o.link ?? ''
    return typeof pick === 'string' ? pick : ''
  }
  return ''
}

/**
 * 清洗成安全的文件/文件夹名。
 *
 * Windows 禁止 `\ / : * ? " < > |`；控制字符和首尾空白也要去掉，
 * 否则解压时会被系统改名甚至失败。长度截到 80，给「-序号.扩展名」留余量。
 */
export function sanitizeName(s: string): string {
  const cleaned = s
    // eslint-disable-next-line no-control-regex
    .replace(/[\\/:*?"<>|\x00-\x1f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    // 前导点会让文件变成隐藏文件，也可能被当成上级目录
    .replace(/^\.+/, '')
    .replace(/\.+$/, '')
  return cleaned.slice(0, 80) || '未命名'
}

/**
 * 取扩展名（含点，小写）。
 *
 * 优先用附件自己的文件名；文件名没有扩展名时才用 MIME 推断。
 *
 * ⚠️ **不能直接用 `guessExtFromMime` 的返回值兜底** —— 它在完全不认识时会返回
 * `'png'`（对图片是个合理默认，但对 PDF/Word/Excel 附件就是错的，
 * 用户拿到一堆后缀为 .png 的合同文件）。这里对非图片类型单独映射。
 */
export function pickExtension(fileName: string, mime: string): string {
  // 1. 文件自带的扩展名最可信
  const m = /\.([A-Za-z0-9]{1,8})$/.exec((fileName || '').trim())
  if (m) return `.${m[1].toLowerCase()}`

  // 2. 常见非图片类型（guessExtFromMime 不覆盖这些）
  const type = (mime || '').toLowerCase()
  const KNOWN: [string, string][] = [
    ['pdf', 'pdf'],
    ['msword', 'doc'],
    ['wordprocessingml', 'docx'],
    ['ms-excel', 'xls'],
    ['spreadsheetml', 'xlsx'],
    ['ms-powerpoint', 'ppt'],
    ['presentationml', 'pptx'],
    ['zip', 'zip'],
    ['rar', 'rar'],
    ['7z', '7z'],
    ['csv', 'csv'],
    ['json', 'json'],
    ['plain', 'txt'],
    ['mp4', 'mp4'],
    ['quicktime', 'mov'],
    ['webm', 'webm'],
    ['mpeg', 'mp3'],
    ['wav', 'wav'],
  ]
  for (const [needle, ext] of KNOWN) {
    if (type.includes(needle)) return `.${ext}`
  }

  // 3. 图片/音视频里 guessExtFromMime 认识的类型
  if (type.startsWith('image/') || /png|jpeg|jpg|gif|webp|bmp|tiff|svg/.test(type)) {
    const guessed = guessExtFromMime(type)
    return guessed ? `.${guessed.replace(/^\./, '').toLowerCase()}` : ''
  }

  // 4. 实在认不出来：宁可不要扩展名，也不要给一个错的
  return ''
}

/**
 * 在已用集合里避免重名：`a.jpg` → `a(2).jpg` → `a(3).jpg`。
 *
 * 为什么不用 `a-1.jpg`：`a-1` 可能和「按字段值命名」产生的 `张三-1` 撞车，
 * 括号是 Windows 资源管理器自己的重名规则，用户一眼就懂。
 */
export function dedupePath(path: string, used: Set<string>): string {
  if (!used.has(path)) {
    used.add(path)
    return path
  }
  const slash = path.lastIndexOf('/')
  const dir = slash >= 0 ? path.slice(0, slash + 1) : ''
  const base = slash >= 0 ? path.slice(slash + 1) : path
  const dot = base.lastIndexOf('.')
  const stem = dot > 0 ? base.slice(0, dot) : base
  const ext = dot > 0 ? base.slice(dot) : ''
  for (let n = 2; n < 10_000; n++) {
    const candidate = `${dir}${stem}(${n})${ext}`
    if (!used.has(candidate)) {
      used.add(candidate)
      return candidate
    }
  }
  const fallback = `${dir}${stem}-${Date.now()}${ext}`
  used.add(fallback)
  return fallback
}

/**
 * 生成附件文件名的主干（不含扩展名）。
 *
 * - `row`   按行号：`0001`；同行同列多张时为 `0001-1`、`0001-2`
 * - `field` 按标题：`张三-1`、`张三-2`（序号按标题全局递增，跨行不会撞名）
 *
 * @param opts.rowNo   四位行号，如 `0001`
 * @param opts.title   标题字段的文本（已清洗）；为空时回退到行号
 * @param opts.index   该单元格内的第几张（0 起）
 * @param opts.total   该单元格共几张 —— 只有多张时才加 `-序号`
 * @param opts.seqByTitle 按标题的累加器，field 模式下保证 `张三-1/2/3` 不重复
 */
export function buildAttachmentStem(opts: {
  naming: 'row' | 'field'
  rowNo: string
  title: string
  index: number
  total: number
  seqByTitle: Map<string, number>
}): string {
  const { naming, rowNo, title, index, total, seqByTitle } = opts

  if (naming === 'field' && title && title !== '未命名') {
    const next = (seqByTitle.get(title) ?? 0) + 1
    seqByTitle.set(title, next)
    return `${title}-${next}`
  }

  // row 模式（或标题为空）→ 行号；多张时补 -1/-2
  return total > 1 ? `${rowNo}-${index + 1}` : rowNo
}
