#!/usr/bin/env node
// wiki-model-freeze — print a model item's text, frozen from a workbook on disk
// or a link-shared Google Sheet. The same code the browser runs; lets an agent
// create or refresh an item.
//
//   wiki-model-freeze FILE.xlsx --source /assets/page/FILE.xlsx --sheet Summary --range C9:E17 \
//     [--name NAME] [--columns "Measure, Unit, Base"] [--caption TEXT] [--layout table|list|auto]
//   wiki-model-freeze https://docs.google.com/spreadsheets/d/ID/edit --sheet Outputs --range A1:C11
//   wiki-model-freeze FILE.xlsx --sheets        list sheets and named ranges

import { readFileSync } from 'node:fs'
import { inflateRawSync } from 'node:zlib'
import { basename } from 'node:path'
import { openWorkbook } from '../src/lib/xlsx.js'
import { readLive, serializeItem, toGfm } from '../src/lib/model.js'
import { sheetRef } from '../src/lib/gsheet.js'

const args = process.argv.slice(2)
const flags = new Set(['--sheets'])
const file = args.find((a, i) => !a.startsWith('--') && !(args[i - 1] && args[i - 1].startsWith('--') && !flags.has(args[i - 1])))
const opt = k => { const i = args.indexOf(`--${k}`); return i >= 0 ? args[i + 1] : undefined }
if (!file) { console.error('usage: wiki-model-freeze FILE.xlsx|SHEET-URL --sheet S --range R [--columns ...]'); process.exit(2) }

const inflateRaw = async b => inflateRawSync(b)
const isSheet = !!sheetRef(file)

if (args.includes('--sheets')) {
  if (isSheet) { console.error('--sheets reads workbooks on disk only'); process.exit(2) }
  const book = await openWorkbook(new Uint8Array(readFileSync(file)), inflateRaw)
  console.log(`sheets: ${book.sheets.join(', ')}`)
  console.log(`named ranges: ${Object.entries(book.names).map(([k, v]) => `${k} = ${v}`).join('; ') || 'none'}`)
  process.exit(0)
}

const model = {
  source: opt('source') || (isSheet ? file : `/assets/${basename(file)}`),
  name: opt('name'),
  sheet: opt('sheet'),
  range: opt('range'),
  columns: opt('columns') ? opt('columns').split(',').map(s => s.trim()) : undefined,
}
const io = { fetchBytes: async () => new Uint8Array(readFileSync(file)), inflateRaw, fetcher: (u, o) => fetch(u, o) }
const readModel = isSheet ? model : { ...model, source: file } // read the local file, record the published source
const { table, revision } = await readLive(readModel, file, io)
model.revision = { hash: revision, at: new Date().toISOString().slice(0, 16) + 'Z' }
const head = []
if (opt('caption')) head.push(`CAPTION ${opt('caption')}`)
if (opt('layout')) head.push(`LAYOUT ${opt('layout')}`)
process.stdout.write(serializeItem({ model, head, body: toGfm(table) }) + '\n')
