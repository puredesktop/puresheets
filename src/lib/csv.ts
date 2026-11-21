import {
  cellKey,
  computedCellValue,
  createBlankSheet,
  createBlankWorkbook,
  normalizeRangeToken,
  parseCellKey,
  type FormulaWorkbook,
} from './workbookModel'
import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT } from '../constants'
import { MAX_PERSISTED_ROWS, MAX_SHEET_COLUMNS } from './sheetMetadataNormalize'
import { withFreshUniverSnapshot } from './univerAdapter'
import type { PureSheetsDocument, SheetCell, WorkbookSheet } from '../types'

const MAX_CSV_IMPORT_CELLS = 250_000
const CSV_FORMULA_INJECTION_RE = /^(?:[\x00-\x1f]|[\s\uFEFF\x00-\x1f]*[=+\-@])/
const PLAIN_SIGNED_NUMBER_RE = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/

function isPlainFiniteNumber(value: string): boolean {
  return PLAIN_SIGNED_NUMBER_RE.test(value) && Number.isFinite(Number(value))
}

function stripBom(text: string): string {
  return text.startsWith('\uFEFF') ? text.slice(1) : text
}

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
  options: {
    delimiter?: ',' | ';' | '\t'
    maxRows?: number
    maxColumns?: number
    maxPopulatedCells?: number
  } = {},
): string[][] {
  const input = stripBom(text)
  const delimiter = options.delimiter ?? detectCsvDelimiter(input)
  const maxRows = options.maxRows ?? MAX_PERSISTED_ROWS
  const maxColumns = options.maxColumns ?? MAX_SHEET_COLUMNS
  const maxPopulatedCells = options.maxPopulatedCells ?? MAX_CSV_IMPORT_CELLS
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  let index = 0
  let populatedCells = 0
  const pushField = (): void => {
    if (row.length >= maxColumns) {
      throw new Error(
        `This CSV exceeds the maximum of ${maxColumns.toLocaleString()} columns.`,
      )
    }
    if (field.trim() !== '') {
      populatedCells += 1
      if (populatedCells > maxPopulatedCells) {
        throw new Error(
          'This CSV has more than 250,000 filled cells, more than PureSheets imports at once. Split it and import the parts.',
        )
      }
    }
    row.push(field)
    field = ''
  }
  const pushRow = (): void => {
    if (rows.length >= maxRows) {
      throw new Error(
        `This CSV exceeds the maximum of ${maxRows.toLocaleString()} rows.`,
      )
    }
    pushField()
    rows.push(row)
    row = []
  }
  while (index < input.length) {
    const current = input[index]
    if (quoted) {
      if (current === '"') {
        if (input[index + 1] === '"') {
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
      if (input[index + 1] === '\n') index += 1
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

/**
 * Serialize rows to CSV with minimal quoting (RFC 4180) and formula-injection
 * neutralization: non-numeric text values starting with spreadsheet formula
 * triggers (=, +, -, @, tab, CR, LF) are prefixed with a single quote and
 * quoted so Excel/Calc never execute untrusted CSV content as formulas.
 */
export function toCsv(rows: string[][], delimiter: ',' | ';' | '\t' = ','): string {
  return rows
    .map(row =>
      row
        .map(value => {
          if (!isPlainFiniteNumber(value) && CSV_FORMULA_INJECTION_RE.test(value)) {
            return `"'${value.replace(/"/g, '""')}"`
          }
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

/**
 * Build a new sheet from CSV text in O(N) time. Leading UTF-8 BOM is stripped,
 * populated cell count is capped at 250,000, and CSV fields beginning with `=`
 * are stored strictly as literal text (never evaluated as formulas).
 */
export function sheetFromCsv(
  text: string,
  id: string,
  name: string,
  options: { delimiter?: ',' | ';' | '\t' } = {},
): WorkbookSheet {
  const rows = parseCsv(stripBom(text), options)
  const cells: Record<string, SheetCell> = {}
  let populated = 0
  let maxColumn = 0
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    if (rowIndex >= MAX_PERSISTED_ROWS) {
      throw new Error(
        `This CSV exceeds the maximum of ${MAX_PERSISTED_ROWS.toLocaleString()} rows.`,
      )
    }
    const row = rows[rowIndex]
    if (row.length > maxColumn) maxColumn = row.length
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (columnIndex >= MAX_SHEET_COLUMNS) {
        throw new Error(
          `This CSV exceeds the maximum of ${MAX_SHEET_COLUMNS.toLocaleString()} columns.`,
        )
      }
      const value = row[columnIndex]
      if (value === '') continue
      const trimmed = value.trim()
      if (!trimmed) continue
      populated += 1
      if (populated > MAX_CSV_IMPORT_CELLS) {
        throw new Error(
          'This CSV has more than 250,000 filled cells, more than PureSheets imports at once. Split it and import the parts.',
        )
      }
      const kind: SheetCell['kind'] =
        !trimmed.startsWith('=') && isPlainFiniteNumber(trimmed)
          ? 'number'
          : 'text'
      cells[cellKey(rowIndex, columnIndex)] = {
        value,
        kind,
        style: undefined,
      }
    }
  }
  const sheet = createBlankSheet(id, name)
  return {
    ...sheet,
    rowCount: Math.max(DEFAULT_ROW_COUNT, rows.length),
    columnCount: Math.max(DEFAULT_COLUMN_COUNT, maxColumn),
    cells,
  }
}

function sheetRowsForExport(
  sheet: WorkbookSheet,
  rangeToken?: string,
  workbook?: FormulaWorkbook,
): string[][] {
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
    if (endRow < 0) return []
  }
  const rows: string[][] = []
  for (let row = startRow; row <= endRow; row += 1) {
    const values: string[] = []
    for (let column = startColumn; column <= endColumn; column += 1) {
      values.push(computedCellValue(sheet, cellKey(row, column), workbook))
    }
    rows.push(values)
  }
  return rows
}

/**
 * Export a sheet (or a range of it) as CSV (or TSV when delimiter is '\t').
 * Formula cells export their computed values — CSV carries data, not formulas.
 */
export function csvFromSheet(
  sheet: WorkbookSheet,
  rangeToken?: string,
  workbook?: FormulaWorkbook,
  delimiter: ',' | ';' | '\t' = ',',
): string {
  const rows = sheetRowsForExport(sheet, rangeToken, workbook)
  if (!rows.length) return ''
  return `${toCsv(rows, delimiter)}\n`
}

/**
 * Export a sheet (or a range of it) as formatted JSON. When the range has at
 * least 2 rows and the first row contains non-empty header labels, emits an
 * array of keyed objects; otherwise emits a 2D array of values.
 */
export function jsonFromSheet(
  sheet: WorkbookSheet,
  rangeToken?: string,
  workbook?: FormulaWorkbook,
): string {
  const rows = sheetRowsForExport(sheet, rangeToken, workbook)
  if (!rows.length) return '[]\n'
  const toTyped = (value: string): string | number | null => {
    if (value === '') return null
    if (isPlainFiniteNumber(value)) return Number(value)
    return value
  }
  if (rows.length >= 2 && rows[0].some(cell => cell.trim() !== '')) {
    const headers = rows[0].map((cell, index) => cell.trim() || `column_${index + 1}`)
    const records = rows.slice(1).map(row => {
      const record: Record<string, string | number | null> = {}
      headers.forEach((header, index) => {
        Object.defineProperty(record, header, {
          value: toTyped(row[index] ?? ''),
          enumerable: true,
          writable: true,
          configurable: true,
        })
      })
      return record
    })
    return `${JSON.stringify(records, null, 2)}\n`
  }
  return `${JSON.stringify(rows.map(row => row.map(toTyped)), null, 2)}\n`
}

/** A `.csv` or `.tsv` path: delimited text, opened as a workbook of one sheet, never read as workbook JSON. */
export function isDelimitedTextPath(path: string): boolean {
  return /\.(csv|tsv)$/i.test(path.trim())
}

/**
 * A whole workbook from one delimited-text file, as opening a CSV as a document makes it: one sheet named after
 * the file (its name without the suffix), and the workbook titled the same.
 */
export function workbookFromCsv(text: string, fileName: string): PureSheetsDocument {
  const base = fileName.replace(/\.[^.]+$/, '').trim() || 'Imported'
  const sheet = sheetFromCsv(text, 'sheet-1', base, {
    ...(fileName.toLowerCase().endsWith('.tsv') ? { delimiter: '\t' as const } : {}),
  })
  const workbook = createBlankWorkbook(base)
  return withFreshUniverSnapshot({
    ...workbook,
    workbook: { activeSheetId: sheet.id, sheets: [sheet] },
  })
}
