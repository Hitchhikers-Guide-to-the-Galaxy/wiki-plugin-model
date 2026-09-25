// xlsx.js — read the values a spreadsheet cached when it was last saved.
// No dependencies: an xlsx file is a zip of XML parts. The caller supplies
// inflateRaw (DecompressionStream in the browser, zlib in node), so the same
// code runs in both. Formulas are never evaluated; <v> is what Excel,
// LibreOffice or openpyxl-with-a-recalc last wrote.

const u16 = (b, o) => b[o] | (b[o + 1] << 8)
const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0
const utf8 = bytes => new TextDecoder('utf-8').decode(bytes)

/** Index a zip by its central directory -> Map(name -> entry). */
export const unzip = bytes => {
  let eocd = -1
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (u32(bytes, i) === 0x06054b50) { eocd = i; break }
  }
  if (eocd < 0) throw new Error('not a zip file (no end of central directory)')
  const count = u16(bytes, eocd + 10)
  let p = u32(bytes, eocd + 16)
  const entries = new Map()
  for (let n = 0; n < count; n++) {
    if (u32(bytes, p) !== 0x02014b50) throw new Error('bad central directory')
    const method = u16(bytes, p + 10)
    const size = u32(bytes, p + 20)
    const nameLen = u16(bytes, p + 28)
    const extraLen = u16(bytes, p + 30)
    const commentLen = u16(bytes, p + 32)
    const local = u32(bytes, p + 42)
    const name = utf8(bytes.subarray(p + 46, p + 46 + nameLen))
    const start = local + 30 + u16(bytes, local + 26) + u16(bytes, local + 28)
    entries.set(name, { method, data: bytes.subarray(start, start + size) })
    p += 46 + nameLen + extraLen + commentLen
  }
  return entries
}

// Element prefixes (<x:sheet>, written by the .NET OpenXML SDK) are dropped so
// one set of patterns reads Excel, LibreOffice and SDK-written files alike.
const unprefix = xml => xml.replace(/<(\/?)[A-Za-z_][\w.-]*:(?=[A-Za-z_])/g, '<$1')

const readEntry = async (entries, name, inflateRaw) => {
  const e = entries.get(name)
  if (!e) return null
  if (e.method === 0) return unprefix(utf8(e.data))
  if (e.method === 8) return unprefix(utf8(await inflateRaw(e.data)))
  throw new Error(`unsupported zip method ${e.method} for ${name}`)
}

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }
export const decode = s =>
  s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, k) =>
    k[0] === '#' ? String.fromCodePoint(k[1] === 'x' || k[1] === 'X' ? parseInt(k.slice(2), 16) : +k.slice(1)) : ENTITIES[k] ?? m)

const attr = (tag, name) => {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(tag)
  return m ? decode(m[1]) : undefined
}

// Excel's built-in number formats that workbooks refer to by id alone
const BUILTIN = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
  37: '#,##0;(#,##0)', 38: '#,##0;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 49: '@',
}

/** A1 address -> { col (0-based), row (1-based) } */
export const cellRef = ref => {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref.toUpperCase())
  if (!m) throw new Error(`bad cell reference ${ref}`)
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { col: col - 1, row: +m[2] }
}

/** "C9:E17" or "Summary!$C$9:$E$17" -> { sheet?, from, to } */
export const parseRange = text => {
  let t = String(text).trim()
  let sheet
  const bang = t.lastIndexOf('!')
  if (bang >= 0) {
    sheet = t.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'")
    t = t.slice(bang + 1)
  }
  const [a, b = a] = t.split(':')
  return { sheet, from: cellRef(a), to: cellRef(b) }
}

/**
 * Open a workbook. Returns { sheets: [names], names: {definedName: ref}, read(sheet) }.
 * read(sheet) -> Map("C9" -> { value, text?, format }) with raw cached values.
 */
export const openWorkbook = async (bytes, inflateRaw) => {
  const entries = unzip(bytes)
  const wb = await readEntry(entries, 'xl/workbook.xml', inflateRaw)
  if (!wb) throw new Error('not a spreadsheet (no xl/workbook.xml)')
  const rels = (await readEntry(entries, 'xl/_rels/workbook.xml.rels', inflateRaw)) || ''
  const targets = {}
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    let target = attr(m[0], 'Target')
    target = target.startsWith('/') ? target.slice(1) : `xl/${target}`
    targets[attr(m[0], 'Id')] = target
  }
  const sheetPaths = {}
  const sheets = []
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = attr(m[0], 'name')
    sheets.push(name)
    sheetPaths[name] = targets[attr(m[0], 'r:id')]
  }
  const names = {}
  for (const m of wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    if (attr(m[1], 'localSheetId') !== undefined) continue // sheet-scoped names are not outputs
    names[attr(m[1], 'name')] = decode(m[2])
  }

  const ss = (await readEntry(entries, 'xl/sharedStrings.xml', inflateRaw)) || ''
  const strings = [...ss.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m =>
    [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(t => decode(t[1])).join(''))

  const styles = (await readEntry(entries, 'xl/styles.xml', inflateRaw)) || ''
  const fmts = { ...BUILTIN }
  for (const m of styles.matchAll(/<numFmt\b[^>]*>/g)) fmts[attr(m[0], 'numFmtId')] = attr(m[0], 'formatCode')
  const xfBlock = (/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles) || [, ''])[1]
  const xfFormats = [...xfBlock.matchAll(/<xf\b[^>]*>/g)].map(m => fmts[attr(m[0], 'numFmtId') || '0'] || 'General')

  const cache = {}
  const read = async sheet => {
    if (cache[sheet]) return cache[sheet]
    const path = sheetPaths[sheet]
    if (!path) throw new Error(`no sheet named ${sheet} (sheets: ${sheets.join(', ')})`)
    const xml = await readEntry(entries, path, inflateRaw)
    const cells = new Map()
    for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = attr(m[1], 'r')
      const type = attr(m[1], 't') || 'n'
      const format = xfFormats[+(attr(m[1], 's') || 0)] || 'General'
      const inner = m[2] || ''
      const v = /<v>([\s\S]*?)<\/v>/.exec(inner)
      let value = null
      if (type === 's' && v) value = strings[+v[1]]
      else if (type === 'inlineStr') value = [...inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(t => decode(t[1])).join('')
      else if (type === 'str' || type === 'e') value = v ? decode(v[1]) : null
      else if (type === 'b') value = v ? v[1] === '1' : null
      else if (v) value = Number(v[1])
      if (value === null || value === '') continue
      cells.set(ref, { value, format })
    }
    return (cache[sheet] = cells)
  }
  return { sheets, names, read }
}

const colName = n => {
  let s = ''
  for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}

/** Cells of a rectangle, row by row: [[cell|null, ...], ...] with empty rows dropped. */
export const readRange = async (book, { sheet, from, to }) => {
  const cells = await book.read(sheet)
  const rows = []
  for (let r = from.row; r <= to.row; r++) {
    const row = []
    for (let c = from.col; c <= to.col; c++) row.push(cells.get(`${colName(c)}${r}`) || null)
    if (row.some(Boolean)) rows.push(row)
  }
  return rows
}

/** sha256 of the file bytes, hex — the revision a frozen table records. */
export const revisionOf = async bytes => {
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes)
  return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, '0')).join('')
}
