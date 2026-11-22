import type { CellStyle } from '../types'

export type NumberFormatFamily = NonNullable<CellStyle['numberFormat']>

/**
 * Number format codes, shared by the Univer adapter and Excel import/export.
 * PureSheets names a format by family (number, currency, percent, date, text)
 * plus decimals; an Excel code that says more than that is kept verbatim as
 * `numberPattern`, so it renders and exports exactly as it came in.
 */

/** The code PureSheets writes for a family and decimal count. */
export function patternForFormat(
  format: NumberFormatFamily,
  decimals?: number,
  currencySymbol: '$' | '€' | '£' | '¥' = '$',
): string {
  const zeros = decimals !== undefined && decimals > 0 ? `.${'0'.repeat(Math.min(decimals, 8))}` : ''
  if (format === 'currency') {
    const sym = currencySymbol === '$' ? '$' : `[$${currencySymbol}]`
    return decimals === undefined ? `${sym}#,##0.00` : `${sym}#,##0${zeros}`
  }
  if (format === 'accounting') {
    const sym = currencySymbol === '$' ? '$' : `[$${currencySymbol}]`
    const body = decimals === undefined ? '#,##0.00' : `#,##0${zeros}`
    return `_("${sym}"* ${body}_);_("${sym}"* (${body});_("${sym}"* "-"??_);_(@_)`
  }
  if (format === 'percent') return decimals === undefined ? '0.00%' : `0${zeros}%`
  if (format === 'date') return 'yyyy-mm-dd'
  if (format === 'time') return 'hh:mm:ss'
  if (format === 'scientific') return decimals === undefined ? '0.00E+00' : `0${zeros}E+00`
  if (format === 'number') return decimals === undefined ? '#,##0.########' : `#,##0${zeros}`
  return '@'
}

/**
 * The family a format code belongs to, or null for General. Quoted text,
 * escapes and bracketed parts ([Red], [$€-407], [h]) are read past first, so
 * "[Red]" does not look like a date for its "d".
 */
export function familyOfPattern(pattern: string | undefined): NumberFormatFamily | null {
  if (!pattern || /^general$/i.test(pattern.trim())) return null
  const first = firstSection(pattern)
  const bare = stripLiterals(first)
  if (bare.trim() === '@') return 'text'
  if (bare.includes('%')) return 'percent'
  if (/e[+-]0/i.test(bare)) return 'scientific'
  // A currency sign counts quoted too: Excel's accounting format writes "$".
  if (/[$€£¥]/.test(first.replace(/\[\$-[^\]]*\]/g, ''))) {
    if (first.includes('* ') && pattern.includes(';')) return 'accounting'
    return 'currency'
  }
  if (/^h{1,2}:m{1,2}(:s{1,2})?/i.test(bare.trim()) && !/\[h\]/i.test(first)) return 'time'
  if (/[ymdhs]/i.test(bare) && !/[#?]/.test(bare)) return 'date'
  if (/[#0?]/.test(bare)) return 'number'
  if (bare.includes('@')) return 'text'
  return null
}

/** Fixed fraction digits in a code's first section; undefined when they are optional (`.##`) or absent. */
export function decimalsOfPattern(pattern: string, format: NumberFormatFamily): number | undefined {
  if (format === 'date' || format === 'time' || format === 'text') return undefined
  const bare = stripLiterals(firstSection(pattern))
  if (/\.#/.test(bare)) return undefined
  const zeros = /\.(0+)/.exec(bare)?.[1]?.length ?? 0
  if (format === 'number') {
    if (zeros > 0) return zeros
    return /^[#,0]+$/.test(bare) ? 0 : undefined
  }
  // currency, accounting, percent, scientific default to two places
  return zeros === 2 ? undefined : zeros
}

/**
 * Read a code into the style fields: family, decimals, and the code itself
 * when the family and decimals would not write it back the same.
 */
export function styleFromPattern(
  pattern: string | undefined,
): Pick<CellStyle, 'numberFormat' | 'decimals' | 'numberPattern' | 'currencySymbol'> | null {
  const format = familyOfPattern(pattern)
  if (!format || !pattern) return null
  const decimals = decimalsOfPattern(pattern, format)
  const currencyMatch = /([€£¥$])/.exec(pattern)
  const currencySymbol =
    (format === 'currency' || format === 'accounting') && currencyMatch
      ? (currencyMatch[1] as '$' | '€' | '£' | '¥')
      : undefined
  return {
    numberFormat: format,
    ...(currencySymbol && currencySymbol !== '$' ? { currencySymbol } : {}),
    ...(decimals !== undefined ? { decimals } : {}),
    ...(patternForFormat(format, decimals, currencySymbol ?? '$') !== pattern ? { numberPattern: pattern } : {}),
  }
}

/** The code a styled cell shows with, or undefined for General. */
export function patternOfStyle(style: CellStyle | undefined): string | undefined {
  if (!style) return undefined
  if (style.numberPattern) return style.numberPattern
  return style.numberFormat
    ? patternForFormat(style.numberFormat, style.decimals, style.currencySymbol)
    : undefined
}

function firstSection(pattern: string): string {
  let quoted = false
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]
    if (char === '\\') index += 1
    else if (char === '"') quoted = !quoted
    else if (char === ';' && !quoted) return pattern.slice(0, index)
  }
  return pattern
}

function stripLiterals(section: string): string {
  return section
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/_./g, '')
    .replace(/\*./g, '')
}
