import {
  cellKey,
  computedCellValue,
  createBlankSheet,
  normalizeRangeToken,
  parseCellKey,
  upsertCell,
} from './workbookModel'
import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT } from '../constants'
import type { WorkbookSheet } from '../types'

/**
 * CSV import/export (Phase S4). RFC 4180-style parsing: quoted fields,
 * escaped quotes (""), embedded delimiters and newlines inside quotes, CRLF
 * or LF row breaks. The delimiter is auto-detected (comma, semicolon, tab)
 * from the first rows unless given explicitly.
 */

export function detectCsvDelimiter(text: string): ',' | ';' | '\t' {
  const sample = text.slice(0, 4096)
  const counts: Array<{ delimiter: ',' | ';' | '\t'; count: number }> = [
    { delimiter: ',', count: countOutsideQuotes(sample, ',') },
    { delimiter: ';', count: countOutsideQuotes(sample, ';') },
    { delimiter: '\t', count: countOutsideQuotes(sample, '\t') },
  ]
  counts.sort((left, right) => right.count - left.count)
  return counts[0].count > 0 ? counts[0].delimiter : ','
}

function countOutsideQuotes(text: string, char: string): number {
  let count = 0
  let quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const current = text[index]
    if (current === '"') quoted = !quoted
    else if (!quoted && current === char) count += 1
  }
  return count
}

export function parseCsv(
  text: string,
  options: { delimiter?: ',' | ';' | '\t' } = {},
): string[][] {
  const delimiter = options.delimiter ?? detectCsvDelimiter(text)
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  let index = 0
  const pushField = (): void => {
    row.push(field)
    field = ''
  }
  const pushRow = (): void => {
    pushField()
    rows.push(row)
    row = []
  }
  while (index < text.length) {
    const current = text[index]
    if (quoted) {
      if (current === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 2
          continue
        }
        quoted = false
        index += 1
        continue
      }
      field += current
      index += 1
      continue
    }
    if (current === '"' && field === '') {
      quoted = true
      index += 1
      continue
    }
    if (current === delimiter) {
      pushField()
      index += 1
      continue
    }
    if (current === '\r') {
      if (text[index + 1] === '\n') index += 1
      pushRow()
      index += 1
      continue
    }
    if (current === '\n') {
      pushRow()
      index += 1
      continue
    }
    field += current
    index += 1
  }
  // Final field/row — unless the file ended exactly on a row break, in which
  // case the last record was already terminated (a trailing newline does not
  // create an empty extra row).
  if (field !== '' || row.length > 0) pushRow()
  return rows
}

/** Serialize rows to CSV with minimal quoting (RFC 4180). */
export function toCsv(rows: string[][], delimiter: ',' | ';' | '\t' = ','): string {
  return rows
    .map(row =>
      row
        .map(value => {
          if (
            value.includes('"') ||
            value.includes(delimiter) ||
            value.includes('\n') ||
            value.includes('\r')
          ) {
            return `"${value.replace(/"/g, '""')}"`
          }
          return value
        })
        .join(delimiter),
    )
    .join('\n')
}

/** Build a new sheet from CSV text, through the normal document model. */
export function sheetFromCsv(
  text: string,
  id: string,
  name: string,
): WorkbookSheet {
  const rows = parseCsv(text)
  let sheet = createBlankSheet(id, name)
  sheet = {
    ...sheet,
    rowCount: Math.max(DEFAULT_ROW_COUNT, rows.length),
    columnCount: Math.max(
      DEFAULT_COLUMN_COUNT,
      rows.reduce((widest, row) => Math.max(widest, row.length), 0),
    ),
  }
  rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      if (value === '') return
      sheet = upsertCell(sheet, cellKey(rowIndex, columnIndex), value)
    })
  })
  return sheet
}

/**
 * Export a sheet (or a range of it) as CSV. Formula cells export their
 * computed values — CSV carries data, not formulas.
 */
export function csvFromSheet(
  sheet: WorkbookSheet,
  rangeToken?: string,
): string {
  let startRow = 0
  let startColumn = 0
  let endRow = -1
  let endColumn = -1
  const normalized = rangeToken ? normalizeRangeToken(rangeToken) : null
  if (normalized) {
    const [startKey, endKey] = normalized.split(':')
    const start = parseCellKey(startKey ?? '')
    const end = parseCellKey(endKey ?? startKey ?? '')
    if (start && end) {
      startRow = start.row
      startColumn = start.column
      endRow = end.row
      endColumn = end.column
    }
  }
  if (endRow < 0 || endColumn < 0) {
    // Used range: bounding box of populated cells.
    for (const key of Object.keys(sheet.cells)) {
      const position = parseCellKey(key)
      if (!position) continue
      endRow = Math.max(endRow, position.row)
      endColumn = Math.max(endColumn, position.column)
    }
    if (endRow < 0) return ''
  }
  const rows: string[][] = []
  for (let row = startRow; row <= endRow; row += 1) {
    const values: string[] = []
    for (let column = startColumn; column <= endColumn; column += 1) {
      values.push(computedCellValue(sheet, cellKey(row, column)))
    }
    rows.push(values)
  }
  return `${toCsv(rows)}\n`
}
