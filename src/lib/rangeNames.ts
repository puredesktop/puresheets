import type { NamedRange, WorkbookSheet } from '../types'

const CELL = /^(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})$/i

export function validRangeName(name: string): boolean {
  return /^[A-Za-z_][A-Za-z_.0-9]{0,254}$/.test(name) &&
    !CELL.test(name) && !/^(?:R|C|R\d+C\d+|TRUE|FALSE)$/i.test(name)
}

export function qualifiedRange(sheet: string, range: string): string {
  return `'${sheet.replace(/'/g, "''")}'!${range.replace(/^(.+):\1$/i, '$1').replace(/([A-Z]+)(\d+)/gi, '$$$1$$$2')}`
}

/** The one contiguous A1 range PureSheets can edit under a name. */
export function parseQualifiedRange(value: string): { sheet: string; range: string } | null {
  const match = /^(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?)$/i.exec(value.trim().replace(/^=/, ''))
  if (!match) return null
  return { sheet: match[1]?.replace(/''/g, "'") ?? match[2], range: match[3].replace(/\$/g, '').toUpperCase() }
}

export function normalizeNamedRanges(value: unknown, sheets: readonly WorkbookSheet[]): NamedRange[] | undefined {
  if (!Array.isArray(value)) return undefined
  const seen = new Set<string>()
  const names: NamedRange[] = []
  for (const entry of value) {
    if (!entry || typeof entry !== 'object' || typeof entry.name !== 'string' || typeof entry.range !== 'string' || !validRangeName(entry.name)) continue
    const key = entry.name.toUpperCase()
    const sheet = sheets.find(sheet => sheet.id === entry.sheetId)
    if (seen.has(key) || !sheet || !parseQualifiedRange(qualifiedRange(sheet.name, entry.range))) continue
    seen.add(key)
    names.push({ name: entry.name, sheetId: entry.sheetId, range: entry.range })
  }
  return names.length ? names : undefined
}

// Match quoted strings and existing references first. Names are workbook scoped.
export function expandRangeNames(formula: string, names: readonly NamedRange[], sheets: readonly WorkbookSheet[]): string {
  if (!names.length) return formula
  const refs = new Map(names.map(name => [name.name.toUpperCase(), name]))
  return formula.replace(/"(?:[^"]|"")*"|(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|[A-Za-z_][\w.]*/gi, (token, offset: number) => {
    if (/^\s*\(/.test(formula.slice(offset + token.length))) return token
    const name = refs.get(token.toUpperCase())
    const sheet = name && sheets.find(sheet => sheet.id === name.sheetId)
    return name && sheet ? qualifiedRange(sheet.name, name.range) : token
  })
}
