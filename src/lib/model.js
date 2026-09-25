// model.js — the model item's text, and the table it freezes.
//
// Item text:
//   MODEL <url>            the workbook (site-relative /assets/... or absolute)
//   SHEET <name>           with RANGE C9:E17, or
//   NAME <definedName>     a named range in the workbook
//   COLUMNS a, b, c        headings; without it the range's first row heads the table
//   REVISION <sha256> <iso date>   which version of the workbook the table was frozen from
//   ...any Table Plugin directive (CAPTION, LAYOUT, SORT, KEY, FOLD, FIT, INDEX)
//   | frozen | GFM | table |
//
// Model directives are peeled off; everything else is Table Plugin text.

import { parseRange, readRange, openWorkbook } from './xlsx.js'
import { formatValue } from './format.js'
import { sheetRef, readSheet } from './gsheet.js'

export const MODEL_DIRECTIVES = ['MODEL', 'SHEET', 'RANGE', 'NAME', 'COLUMNS', 'REVISION']

export const parseItem = text => {
  const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n')
  const model = {}
  const head = []
  let i = 0
  for (; i < lines.length; i++) {
    const line = lines[i].trim()
    if (!line) continue
    const m = /^([A-Z]+)(?::)?(?:\s+(.*))?$/.exec(line)
    if (!m) break
    const [, key, value = ''] = m
    if (key === 'MODEL') model.source = value.trim()
    else if (key === 'SHEET') model.sheet = value.trim()
    else if (key === 'RANGE') model.range = value.trim()
    else if (key === 'NAME') model.name = value.trim()
    else if (key === 'COLUMNS') model.columns = value.split(',').map(s => s.trim()).filter(Boolean)
    else if (key === 'REVISION') {
      const [hash = '', at = ''] = value.trim().split(/\s+/)
      model.revision = { hash: hash.replace(/^sha256:/, ''), at }
    } else head.push(line)
  }
  const body = lines.slice(i).join('\n').trim()
  return { model, head, body, tableText: [...head, '', body].join('\n').trim() }
}

export const serializeItem = ({ model, head = [], body = '' }) => {
  const out = []
  if (model.source) out.push(`MODEL ${model.source}`)
  if (model.name) out.push(`NAME ${model.name}`)
  if (model.sheet) out.push(`SHEET ${model.sheet}`)
  if (model.range) out.push(`RANGE ${model.range}`)
  if (model.columns && model.columns.length) out.push(`COLUMNS ${model.columns.join(', ')}`)
  if (model.revision && model.revision.hash) out.push(`REVISION sha256:${model.revision.hash} ${model.revision.at || ''}`.trim())
  return [...out, ...head, '', body].join('\n').trim()
}

/** Where the item points in the workbook -> { sheet, from, to }. */
export const selection = (model, book) => {
  if (model.name) {
    const ref = book.names[model.name]
    if (!ref) throw new Error(`no named range ${model.name} in the workbook`)
    const r = parseRange(ref)
    if (!r.sheet) throw new Error(`named range ${model.name} has no sheet`)
    return r
  }
  if (!model.range) throw new Error('say which cells: RANGE C9:E17 (with SHEET) or NAME')
  const r = parseRange(model.range)
  const sheet = r.sheet || model.sheet || book.sheets[0]
  return { ...r, sheet }
}

/** Rows of display strings -> { columns, rows }: COLUMNS, else the first row heads it. */
export const shape = (rows, model) => {
  rows = rows.map(r => r.slice())
  let columns = model.columns && model.columns.length ? model.columns.slice() : null
  if (!columns) columns = (rows.shift() || []).map((c, n) => c || `Column ${n + 1}`)
  const width = columns.length
  rows = rows.map(r => (r.length >= width ? r.slice(0, width) : [...r, ...Array(width - r.length).fill('')]))
  return { columns, rows }
}

/** Read the selected cells of an opened workbook as a display table. */
export const liveTable = async (model, book) => {
  const sel = selection(model, book)
  const cells = await readRange(book, sel)
  const rows = cells.map(row => row.map(c => (c ? formatValue(c.value, c.format) : '')))
  return { ...shape(rows, model), sheet: sel.sheet }
}

const sha256 = async text => {
  const digest = await globalThis.crypto.subtle.digest('SHA-256', new TextEncoder().encode(text))
  return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, '0')).join('')
}

/** The revision of what the item shows: a sha256 of the displayed values. It
 *  ignores the file itself, so a re-export, or an edit elsewhere in the
 *  workbook, leaves it unchanged. */
export const tableRevision = table => sha256(JSON.stringify([table.columns, table.rows]))

/**
 * Read the live table from whatever MODEL names -> { table, revision, kind }.
 * io: { fetchBytes(url) -> Uint8Array, inflateRaw, fetcher }
 */
export const readLive = async (model, url, io) => {
  const ref = sheetRef(model.source)
  let table
  if (ref) {
    table = shape(await readSheet(ref, model, io.fetcher), model)
  } else {
    const book = await openWorkbook(await io.fetchBytes(url), io.inflateRaw)
    table = await liveTable(model, book)
  }
  return { table, revision: await tableRevision(table), kind: ref ? 'sheet' : 'workbook' }
}

const cell = s => String(s).replace(/\|/g, '\\|').replace(/\n/g, ' ')

export const toGfm = ({ columns, rows }) =>
  [
    `| ${columns.map(cell).join(' | ')} |`,
    `|${columns.map(() => '---').join('|')}|`,
    ...rows.map(r => `| ${r.map(cell).join(' | ')} |`),
  ].join('\n')

/** Parse a frozen GFM body back to { columns, rows } for comparison. */
export const fromGfm = body => {
  const lines = String(body || '').split('\n').map(l => l.trim()).filter(l => l.startsWith('|'))
  const split = l => l.replace(/^\|/, '').replace(/\|$/, '').split(/(?<!\\)\|/).map(c => c.trim().replace(/\\\|/g, '|'))
  const rows = lines.filter(l => !/^\|[\s:|-]+\|$/.test(l)).map(split)
  return { columns: rows.shift() || [], rows }
}

/** Cells that differ between frozen and live, matched by row label (first column). */
export const changes = (frozen, live) => {
  const out = []
  const byLabel = new Map(frozen.rows.map(r => [r[0], r]))
  for (const r of live.rows) {
    const old = byLabel.get(r[0])
    if (!old) { out.push({ label: r[0], column: '', from: '', to: 'new row' }); continue }
    r.forEach((v, n) => {
      if (n && (old[n] ?? '') !== v) out.push({ label: r[0], column: live.columns[n] || '', from: old[n] ?? '', to: v })
    })
  }
  const liveLabels = new Set(live.rows.map(r => r[0]))
  for (const r of frozen.rows) if (!liveLabels.has(r[0])) out.push({ label: r[0], column: '', from: 'row', to: 'removed' })
  return out
}
