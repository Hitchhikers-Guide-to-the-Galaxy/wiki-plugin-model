// wiki-plugin-model — client. A frozen summary of a spreadsheet model, drawn
// through the Table Plugin; a badge compares the frozen REVISION with the
// workbook's current bytes; clicking it thaws (live figures, changes marked);
// an owner can Freeze the live figures back into the item's own text.

import { parseItem, serializeItem, readLive, toGfm, fromGfm, changes } from '../lib/model.js'
import { sheetRef, PrivateSheet } from '../lib/gsheet.js'

const escape = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

const CSS = `
.item.model.data { float: none; width: auto; margin: 0 0 1em 0; padding: 0; background: transparent; overflow: visible; }
.item.model.data p { text-align: left; color: inherit; }
.item.model .item.table { margin: 0; padding: 0; }
.model-foot { font-size: 12px; color: #777; margin-top: 2px; }
.model-badge { font-size: 11px; padding: 0 6px; border-radius: 8px; cursor: pointer; white-space: nowrap; user-select: none; }
.model-badge.ok { background: #e3f5e3; color: #1e6b1e; }
.model-badge.stale { background: #fdf0d8; color: #8a5a00; }
.model-badge.missing { background: #fbe3e3; color: #9b1c1c; }
.model-badge.live { background: #e3ecfb; color: #1f4b99; }
.model-badge.pending { background: #eee; color: #666; }
.model-changes { font-size: 12px; margin: 4px 0; }
.model-changes li { margin: 0; }
.model-button { font-size: 11px; margin-left: 6px; cursor: pointer; }
`
const cssOnce = () => {
  if (document.getElementById('model-plugin-css')) return
  const style = document.createElement('style')
  style.id = 'model-plugin-css'
  style.textContent = CSS
  document.head.appendChild(style)
}

const inflateRaw = async bytes =>
  new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())

// A site-relative MODEL resolves against the site the page came from.
const sourceUrl = ($item, source) => {
  if (/^https?:\/\//.test(source)) return source
  const site = $item.parents('.page:first').data('site')
  return site && site !== location.host ? `//${site}${source}` : source
}

const fetchBytes = async url => {
  const res = await fetch(url, { cache: 'no-store' })
  if (!res.ok) throw new Error(`${res.status} fetching the workbook`)
  return new Uint8Array(await res.arrayBuffer())
}

const io = { fetchBytes, inflateRaw, fetcher: (u, o) => fetch(u, o) }

// ---- drawing through the Table Plugin ---------------------------------------

const drawTable = async ($item, item, $target, text, resource = null) => {
  $target.empty()
  let table = null
  try { table = await wiki.getPlugin('table') } catch (e) { table = null }
  if (!table || !table.emit) {
    $target.html(`<pre style="white-space:pre-wrap;font-size:12px">${escape(resource || text)}</pre><p class="model-foot">Install the Table Plugin to draw this as a table.</p>`)
    return
  }
  const $child = $('<div class="item table"></div>').appendTo($target)
  const pseudo = { type: 'table', id: `${item.id}-table`, text, resource }
  table.emit($child, pseudo)
  table.bind && table.bind($child, pseudo)
  // the child is drawing, not a story item: its editor must open the model item,
  // and wiki.getData must find the model item, not the child
  $child.removeClass('data')
  $child.children().on('dblclick', e => {
    e.stopPropagation()
    if (!$(e.target).closest('a, button').length) wiki.textEditor($item, item)
  })
  item.data = pseudo.data
  item.columns = pseudo.columns
  $item.addClass('data')
}

// ---- badge ------------------------------------------------------------------

const setBadge = ($item, cls, label, title = '') =>
  $item.find('.model-badge').attr('class', `model-badge ${cls}`).attr('title', title).text(label)

const check = async ($item, item, state) => {
  const { model } = state.parsed
  if (!model.source) return setBadge($item, 'missing', 'No MODEL', 'Say which workbook or sheet: MODEL /assets/<page>/<file>.xlsx or a Google Sheets link')
  try {
    state.live = await readLive(model, sourceUrl($item, model.source), io)
  } catch (e) {
    state.live = null
    if (e instanceof PrivateSheet) return setBadge($item, 'missing', 'Private sheet', e.message)
    return setBadge($item, 'missing', sheetRef(model.source) ? 'Cannot read sheet' : 'Source missing', e.message)
  }
  const frozen = model.revision && model.revision.hash
  if (!frozen) return setBadge($item, 'stale', 'Not frozen', 'Click to see the live figures and freeze them')
  if (state.live.revision === frozen) {
    return setBadge($item, 'ok', 'Up to date', `Frozen ${model.revision.at || ''}; the figures shown are the ${state.live.kind}'s current figures`)
  }
  setBadge($item, 'stale', 'Update available', `The ${state.live.kind}'s figures have changed since ${model.revision.at || 'the freeze'}. Click to see what changed`)
}

// ---- thaw and freeze ---------------------------------------------------------

const canWrite = () => typeof isOwner !== 'undefined' && !!isOwner

const thaw = async ($item, item, state) => {
  setBadge($item, 'pending', 'Reading…')
  if (!state.live) await check($item, item, state)
  if (!state.live) return
  const { head, body } = state.parsed
  const live = state.live.table
  state.thawed = live
  const diff = changes(fromGfm(body), live)
  const changed = new Set(diff.map(d => `${d.label}\u0000${d.column}`))
  const marked = { columns: live.columns, rows: live.rows.map(r => r.map((v, n) => (n && changed.has(`${r[0]}\u0000${live.columns[n]}`) ? `**${v}**` : v))) }
  await drawTable($item, item, $item.find('.model-table'), head.join('\n'), toGfm(marked))
  setBadge($item, 'live', 'Live', `Live from the ${state.live.kind} — click to return to the frozen figures`)
  const list = diff.length
    ? `<ul>${diff.map(d => `<li>${escape(d.label)}${d.column ? ` · ${escape(d.column)}` : ''}: ${escape(d.from || '—')} → <b>${escape(d.to || '—')}</b></li>`).join('')}</ul>`
    : '<p>No figures differ from the frozen table.</p>'
  const freeze = canWrite()
    ? `<button class="model-button model-freeze">Freeze these figures</button>`
    : '<span> · log in as the owner to freeze</span>'
  $item.find('.model-changes').html(`${list}<p>${diff.length} change${diff.length === 1 ? '' : 's'}${freeze}<button class="model-button model-cancel">Back to frozen</button><span class="model-out"></span></p>`)
}

const unthaw = async ($item, item, state) => {
  state.thawed = null
  $item.find('.model-changes').empty()
  await drawTable($item, item, $item.find('.model-table'), state.parsed.tableText)
  await check($item, item, state)
}

const freeze = async ($item, item, state) => {
  const { model, head } = state.parsed
  const next = { ...model, revision: { hash: state.live.revision, at: new Date().toISOString().slice(0, 16) + 'Z' } }
  item.text = serializeItem({ model: next, head, body: toGfm(state.thawed) })
  const $page = $item.parents('.page:first')
  wiki.pageHandler.put($page, { type: 'edit', id: item.id, item })
  // pageHandler.put has no callback and falls back silently to local storage,
  // so read the page back before claiming the freeze was saved
  const slug = $page.attr('id')
  let saved = false
  for (const wait of [600, 1200, 2400]) {
    await new Promise(r => setTimeout(r, wait))
    try {
      const page = await (await fetch(`/${slug}.json`, { cache: 'no-store' })).json()
      saved = page.story.some(i => i.id === item.id && i.text === item.text)
    } catch (e) { saved = false }
    if (saved) break
  }
  state.parsed = parseItem(item.text)
  await unthaw($item, item, state)
  $item.find('.model-changes').html(saved
    ? '<p>Frozen and saved to the page journal.</p>'
    : '<p style="color:#9b1c1c">Not saved to the server: this browser kept a local copy. Log in as the site owner and freeze again.</p>')
}

// ---- plugin interface --------------------------------------------------------

const emit = ($item, item) => {
  cssOnce()
  const state = { parsed: parseItem(item.text) }
  $item.data('modelState', state)
  const src = state.parsed.model.source
  $item.html(`
    <div class="model-item">
      <div class="model-table"></div>
      <div class="model-foot"><span class="model-badge pending">Checking…</span>
        ${src ? `<a class="model-source" href="${escape(src)}" target="_blank" title="open the source">${escape(sheetRef(src) ? 'Google Sheet' : src.split('/').pop())}</a>` : ''}</div>
      <div class="model-changes"></div>
    </div>`)
  if (!state.parsed.body && src) {
    // nothing frozen yet: show the live figures straight away
    thaw($item, item, state)
  } else {
    drawTable($item, item, $item.find('.model-table'), state.parsed.tableText).then(() => check($item, item, state))
  }
}

const bind = ($item, item) => {
  $item.on('dblclick', e => {
    if ($(e.target).closest('a, button, .model-badge').length) return
    wiki.textEditor($item, item)
  })
  $item.on('click', '.model-badge', e => {
    e.stopPropagation()
    const state = $item.data('modelState')
    if (state.thawed) unthaw($item, item, state)
    else thaw($item, item, state)
  })
  $item.on('click', '.model-cancel', e => { e.stopPropagation(); unthaw($item, item, $item.data('modelState')) })
  $item.on('click', '.model-freeze', e => {
    e.stopPropagation()
    $(e.target).prop('disabled', true).text('Freezing…')
    freeze($item, item, $item.data('modelState'))
  })
}

if (typeof window !== 'undefined') {
  window.plugins = window.plugins || {}
  window.plugins.model = { emit, bind }
}

export { emit, bind }
