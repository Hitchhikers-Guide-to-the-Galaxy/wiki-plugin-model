// gsheet.js — read a link-shared Google Sheet through gviz, the Google
// Visualization data source endpoint every sheet exposes at /gviz/tq.
// It answers CORS for any origin and returns display-formatted, computed
// values, so no number formatting happens here. It does not take named
// ranges ("Cannot load spreadsheet"): use SHEET (or the gid in the URL) and
// RANGE. A private sheet answers 401, which a browser sees as a failed fetch.

/** Is this MODEL a Google Sheet? -> { id, gid } or null */
export const sheetRef = source => {
  const m = /docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]+)/.exec(String(source || ''))
  if (!m) return null
  const gid = /[#?&]gid=(\d+)/.exec(source)
  return { id: m[1], gid: gid ? gid[1] : null }
}

// The plain CSV export: untyped, honours range, CORS for any origin, but it
// needs the tab's gid. gviz types each column by majority and blanks the cells
// that do not fit (a text heading over numbers, "Not reached" among figures),
// so it is only the fallback for a SHEET name without a gid.
export const exportUrl = (ref, { range } = {}) => {
  const q = new URLSearchParams({ format: 'csv', gid: ref.gid })
  if (range) q.set('range', range)
  return `https://docs.google.com/spreadsheets/d/${ref.id}/export?${q}`
}

export const gvizUrl = (ref, { sheet, range } = {}) => {
  const q = new URLSearchParams({ tqx: 'out:csv' })
  if (sheet) q.set('sheet', sheet)
  else if (ref.gid) q.set('gid', ref.gid)
  if (range) q.set('range', range)
  return `https://docs.google.com/spreadsheets/d/${ref.id}/gviz/tq?${q}`
}

/** RFC 4180 CSV -> rows of strings */
export const parseCsv = text => {
  const rows = []
  let row = []
  let cell = ''
  let quoted = false
  const s = String(text).replace(/\r\n?/g, '\n')
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (quoted) {
      if (ch === '"' && s[i + 1] === '"') { cell += '"'; i++ }
      else if (ch === '"') quoted = false
      else cell += ch
    } else if (ch === '"') quoted = true
    else if (ch === ',') { row.push(cell); cell = '' }
    else if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = '' }
    else cell += ch
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row) }
  return rows
}

export class PrivateSheet extends Error {}

/** Fetch the selected cells as rows of display strings. */
export const readSheet = async (ref, model, fetcher = fetch) => {
  if (model.name) throw new Error('gviz cannot read named ranges: use SHEET and RANGE')
  const range = model.range ? model.range.replace(/^.*!/, '') : undefined
  let res
  try {
    const url = ref.gid && !model.sheet ? exportUrl(ref, { range }) : gvizUrl(ref, { sheet: model.sheet, range })
    res = await fetcher(url, { cache: 'no-store' })
  } catch (e) {
    throw new PrivateSheet('the sheet did not answer: share it as "anyone with the link can view"')
  }
  const type = res.headers.get('content-type') || ''
  if (res.status === 401 || res.status === 403 || /html/.test(type)) {
    throw new PrivateSheet('the sheet is private: share it as "anyone with the link can view"')
  }
  if (!res.ok) throw new Error(`${res.status} from Google Sheets`)
  const text = await res.text()
  if (/^\)\]\}'/.test(text) || /"status":"error"/.test(text)) {
    const m = /"detailed_message":"([^"]*)"/.exec(text)
    throw new Error(`Google Sheets: ${m ? m[1] : 'error'}`)
  }
  return parseCsv(text).filter(r => r.some(c => c.trim() !== ''))
}
