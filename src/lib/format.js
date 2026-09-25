// format.js — show a cached number the way the spreadsheet shows it, from the
// cell's own number format code. Covers what financial models use: sections
// (positive;negative;zero), thousands separators, fixed decimals, percent,
// quoted literals such as "£" or "x", brackets for negatives and colour tags.
// Dates and scientific notation fall back to General.

const splitSections = code => {
  const out = []
  let cur = ''
  let quoted = false
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') quoted = !quoted
    if (ch === '\\' && !quoted) { cur += ch + (code[++i] ?? ''); continue }
    if (ch === ';' && !quoted) { out.push(cur); cur = ''; continue }
    cur += ch
  }
  out.push(cur)
  return out
}

const general = v => {
  if (Number.isInteger(v)) return String(v)
  return String(Number(v.toPrecision(10)))
}

const withThousands = s => s.replace(/\B(?=(\d{3})+(?!\d))/g, ',')

// one section, applied to a non-negative number
const applySection = (section, n) => {
  const s = section.replace(/\[[^\]]*\]/g, '') // [Red], [$£-809], conditions
  if (/[ymdhs]/i.test(s.replace(/"[^"]*"/g, '')) || /E[+-]/i.test(s)) return null
  let out = ''
  let numeric = ''
  let state = 'before' // before -> in -> after the digit run
  let percent = false
  const literal = t => {
    if (state === 'in') state = 'after'
    out += t
  }
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (ch === '"') {
      const end = s.indexOf('"', i + 1) < 0 ? s.length : s.indexOf('"', i + 1)
      literal(s.slice(i + 1, end))
      i = end
    } else if (ch === '\\') literal(s[++i] ?? '')
    else if (ch === '_') { literal(' '); i++ } // the width of the next char
    else if (ch === '*') i++ // fill character
    else if ('0#?,.'.includes(ch) && state !== 'after') {
      if (state === 'before') { out += '\u0000'; state = 'in' }
      numeric += ch
    } else {
      if (ch === '%') percent = true
      literal(ch)
    }
  }
  if (!numeric) return out // a literal section, e.g. "–" for zero
  if (percent) n *= 100
  const [intPart, fracPart = ''] = numeric.split('.')
  const decimals = (fracPart.match(/[0#?]/g) || []).length
  const thousands = /[0#?],[0#?]/.test(intPart)
  let [whole, frac = ''] = n.toFixed(decimals).split('.')
  if (whole === '0' && !intPart.includes('0')) whole = ''
  if (thousands) whole = withThousands(whole)
  return out.replace('\u0000', decimals ? `${whole}.${frac}` : whole || '0')
}

/** Format a raw cell value with its number format code. Strings pass through. */
export const formatValue = (value, code = 'General') => {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value !== 'number' || !Number.isFinite(value)) return value == null ? '' : String(value)
  if (!code || code === 'General' || code === '@') return general(value)
  const sections = splitSections(code)
  let section = sections[0]
  let n = value
  let sign = ''
  if (value < 0) {
    if (sections.length > 1 && sections[1] !== '') { section = sections[1]; n = -value }
    else { n = -value; sign = '-' }
  } else if (value === 0 && sections.length > 2) section = sections[2]
  const out = applySection(section, n)
  return out === null ? general(value) : sign + out
}
