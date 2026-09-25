#!/usr/bin/env node
// wiki-model-freeze — print a model item's text, frozen from a workbook on disk.
// The same code the browser runs; lets an agent create or refresh an item.
//
//   wiki-model-freeze FILE.xlsx --source /assets/page/FILE.xlsx --sheet Summary --range C9:E17 \
//     [--name NAME] [--columns "Measure, Unit, Base"] [--caption TEXT] [--layout table|list|auto]
//   wiki-model-freeze FILE.xlsx --sheets        list sheets and named ranges

import { readFileSync } from 'node:fs'
import { inflateRawSync } from 'node:zlib'
import { basename } from 'node:path'
import { openWorkbook, revisionOf } from '../src/lib/xlsx.js'
import { liveTable, serializeItem, toGfm } from '../src/lib/model.js'

const args = process.argv.slice(2)
const file = args.find(a => !a.startsWith('--') && !args[args.indexOf(a) - 1]?.startsWith('--'))
const opt = k => { const i = args.indexOf(`--${k}`); return i >= 0 ? args[i + 1] : undefined }
if (!file) { console.error('usage: wiki-model-freeze FILE.xlsx --source URL --sheet S --range R [--columns ...]'); process.exit(2) }

const bytes = new Uint8Array(readFileSync(file))
const book = await openWorkbook(bytes, async b => inflateRawSync(b))
if (args.includes('--sheets')) {
  console.log(`sheets: ${book.sheets.join(', ')}`)
  console.log(`named ranges: ${Object.entries(book.names).map(([k, v]) => `${k} = ${v}`).join('; ') || 'none'}`)
  process.exit(0)
}
const model = {
  source: opt('source') || `/assets/${basename(file)}`,
  name: opt('name'),
  sheet: opt('sheet'),
  range: opt('range'),
  columns: opt('columns') ? opt('columns').split(',').map(s => s.trim()) : undefined,
}
const table = await liveTable(model, book)
model.revision = { hash: await revisionOf(bytes), at: new Date().toISOString().slice(0, 16) + 'Z' }
const head = []
if (opt('caption')) head.push(`CAPTION ${opt('caption')}`)
if (opt('layout')) head.push(`LAYOUT ${opt('layout')}`)
process.stdout.write(serializeItem({ model, head, body: toGfm(table) }) + '\n')
