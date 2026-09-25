import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { inflateRawSync } from 'node:zlib'
import { openWorkbook, parseRange, cellRef, revisionOf } from '../src/lib/xlsx.js'
import { formatValue } from '../src/lib/format.js'
import { parseItem, serializeItem, liveTable, toGfm, fromGfm, changes } from '../src/lib/model.js'

const inflate = async b => inflateRawSync(b)
const bytes = new Uint8Array(readFileSync(new URL('./fixtures/itt.xlsx', import.meta.url)))

test('cell and range references', () => {
  assert.deepEqual(cellRef('E9'), { col: 4, row: 9 })
  assert.deepEqual(cellRef('$AA$10'), { col: 26, row: 10 })
  assert.deepEqual(parseRange("'Partner Service'!$C$1:$I$4"), { sheet: 'Partner Service', from: { col: 2, row: 1 }, to: { col: 8, row: 4 } })
})

test('number formats used by financial models', () => {
  const acc = '#,##0;(#,##0);"–"'
  assert.equal(formatValue(110000, acc), '110,000')
  assert.equal(formatValue(-52000, acc), '(52,000)')
  assert.equal(formatValue(0, acc), '–')
  assert.equal(formatValue(33803.63805095887, acc), '33,804')
  assert.equal(formatValue(-354909.9, '#,##0;[Red](#,##0);"–"'), '(354,910)')
  assert.equal(formatValue(1.0309846324311447, '0.000"x"'), '1.031x')
  assert.equal(formatValue(0.011461192179219051, '0.00%;(0.00%);"–"'), '1.15%')
  assert.equal(formatValue(0.24465810783744, '0.0%'), '24.5%')
  assert.equal(formatValue(520000, '"£"#,##0;("£"#,##0);"–"'), '£520,000')
  assert.equal(formatValue(-12.5, '0.00'), '-12.50')
  assert.equal(formatValue(2.75, 'General'), '2.75')
  assert.equal(formatValue('Not reached', acc), 'Not reached')
})

test('reads the ITT workbook cached values', async () => {
  const book = await openWorkbook(bytes, inflate)
  assert.deepEqual(book.sheets, ['Summary', 'Assumptions', 'Compute', 'Institute', 'Founding Budget', 'Checks', 'Partner Service'])
  const cells = await book.read('Summary')
  assert.equal(cells.get('C9').value, 'Year one customer billings')
  assert.equal(cells.get('E9').value, 110000)
  assert.equal(cells.get('E38').format, '0.000"x"')
})

test('a summary table from SHEET and RANGE', async () => {
  const book = await openWorkbook(bytes, inflate)
  const { model } = parseItem('MODEL /assets/x.xlsx\nSHEET Summary\nRANGE C33:E47\nCOLUMNS Measure, Unit, Base\nCAPTION Ownership')
  const t = await liveTable(model, book)
  assert.deepEqual(t.columns, ['Measure', 'Unit', 'Base'])
  const row = label => t.rows.find(r => r[0] === label)
  assert.deepEqual(row('Five year distributed cash multiple'), ['Five year distributed cash multiple', 'x', '1.031x'])
  assert.deepEqual(row('Five year cash flow IRR'), ['Five year cash flow IRR', '%', '1.15%'])
  assert.deepEqual(row('Renewal funded'), ['Renewal funded', '%', '24.5%'])
  assert.deepEqual(row('Unfunded renewal amount'), ['Unfunded renewal amount', 'GBP', '392,778'])
  assert.deepEqual(row('Optional managed GPU ownership investment'), ['Optional managed GPU ownership investment', '', ''])
})

test('item text round-trips and keeps table directives', () => {
  const text = 'MODEL /assets/x.xlsx\nSHEET Summary\nRANGE C9:E17\nREVISION sha256:abc 2026-09-25\nCAPTION Feasibility\nLAYOUT table\n\n| A | B |\n|---|---|\n| 1 | 2 |'
  const p = parseItem(text)
  assert.equal(p.model.revision.hash, 'abc')
  assert.deepEqual(p.head, ['CAPTION Feasibility', 'LAYOUT table'])
  assert.ok(p.tableText.startsWith('CAPTION Feasibility\nLAYOUT table'))
  assert.equal(serializeItem(p), 'MODEL /assets/x.xlsx\nSHEET Summary\nRANGE C9:E17\nREVISION sha256:abc 2026-09-25\nCAPTION Feasibility\nLAYOUT table\n\n| A | B |\n|---|---|\n| 1 | 2 |')
})

test('changes between a frozen and a live table', () => {
  const frozen = fromGfm(toGfm({ columns: ['M', 'V'], rows: [['a', '1'], ['b', '2'], ['c', '3']] }))
  const live = { columns: ['M', 'V'], rows: [['a', '1'], ['b', '5'], ['d', '4']] }
  assert.deepEqual(changes(frozen, live), [
    { label: 'b', column: 'V', from: '2', to: '5' },
    { label: 'd', column: '', from: '', to: 'new row' },
    { label: 'c', column: '', from: 'row', to: 'removed' },
  ])
})

test('revision is the sha256 of the bytes', async () => {
  const r = await revisionOf(bytes)
  assert.match(r, /^[0-9a-f]{64}$/)
})
