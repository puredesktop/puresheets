import type { PivotDefinition, PureSheetsDocument, SheetCell, WorkbookSheet } from '../types'
import { expandRangeNames } from './rangeNames'
import {
  cellKey, cellsInRange, columnLabel, computedCellValue, isCellLocked, isFormulaError,
  normalizeRangeToken, parseCellKey, rangesIntersect,
} from './workbookModel'

type Workbook = PureSheetsDocument['workbook']
export const MAX_ANALYSIS_CELLS = 25_000

export function readSelection(sheet: WorkbookSheet, range: string, workbook: Workbook) {
  const keys = cellsInRange(range, MAX_ANALYSIS_CELLS)
  const first = parseCellKey(keys[0] ?? '')
  const last = parseCellKey(keys.at(-1) ?? '')
  if (!first || !last || last.column >= sheet.columnCount || last.row >= sheet.rowCount)
    throw new Error('Select a range within the sheet with up to 25,000 cells.')
  const width = last.column - first.column + 1
  const rows: Array<Array<{ key: string; cell?: SheetCell; value: string }>> = []
  for (let index = 0; index < keys.length; index += width) {
    rows.push(keys.slice(index, index + width).map(key => ({
      key,
      cell: sheet.cells[key],
      value: computedCellValue(sheet, key, workbook),
    })))
  }
  return { rows, width, first }
}

export function cleanedRows(
  selection: ReturnType<typeof readSelection>,
  options: { mode: 'dedupe' | 'split'; columns: number[]; separator: string; header: boolean },
): SheetCell[][] {
  const { rows, width } = selection
  const original = (row: typeof rows[number]) => row.map(({ cell, value }): SheetCell => ({
    ...cell,
    value,
    kind: cell?.kind === 'formula' ? (value !== '' && Number.isFinite(Number(value)) ? 'number' : 'text') : cell?.kind ?? 'blank',
  }))
  if (options.mode === 'dedupe') {
    if (!options.columns.length) throw new Error('Select at least one column to compare.')
    const seen = new Set<string>()
    return rows.flatMap((row, index) => {
      if (index === 0 && options.header) return [original(row)]
      const key = JSON.stringify(options.columns.map(column => row[column]?.value ?? ''))
      if (seen.has(key)) return []
      seen.add(key)
      return [original(row)]
    })
  }
  if (width !== 1) throw new Error('Select one column to split.')
  if (!options.separator.length) throw new Error('Enter a delimiter.')
  const result = rows.map((row, index) => index === 0 && options.header
    ? original(row)
    : splitText(row[0].value, options.separator).map(value => ({ value, kind: 'text' as const })))
  if (result.some(row => row.length > 256) || result.reduce((size, row) => size + row.length, 0) > MAX_ANALYSIS_CELLS)
    throw new Error('The split makes too many cells. Select a smaller range.')
  if (options.header) {
    const label = result[0][0].value || 'Part'
    const count = Math.max(...result.map(row => row.length))
    if (count > 1) result[0] = Array.from({ length: count }, (_, index) => ({ value: `${label} ${index + 1}`, kind: 'text' }))
  }
  return result
}

function splitText(value: string, delimiter: string): string[] {
  // Quoted CSV/TSV fields can contain the delimiter. Other delimiters are literal.
  if (delimiter !== ',' && delimiter !== '\t') return value.split(delimiter)
  const cells: string[] = []
  let field = ''
  let quoted = false
  for (let i = 0; i < value.length; i++) {
    const ch = value[i]
    if (ch === '"' && quoted && value[i + 1] === '"') { field += ch; i++; continue }
    if (ch === '"' && (quoted || field === '')) { quoted = !quoted; continue }
    if (ch === delimiter && !quoted) { cells.push(field); field = ''; continue }
    field += ch
  }
  if (quoted) throw new Error('A field has an unclosed quote. Fix it before splitting.')
  cells.push(field)
  return cells
}

export function applyRows(sheet: WorkbookSheet, rows: SheetCell[][], at = 'A1', previous?: { rows: number; columns: number }): WorkbookSheet {
  const origin = parseCellKey(at)
  if (!origin) throw new Error('Enter a valid output cell.')
  const width = Math.max(0, ...rows.map(row => row.length))
  if (!width || origin.column + width > 16_384 || origin.row + rows.length > 1_048_576)
    throw new Error('The result does not fit on this sheet.')
  const cells = { ...sheet.cells }
  const owned = (r: number, c: number) => !!previous && r < previous.rows && c < previous.columns
  for (let r = 0; r < Math.max(rows.length, previous?.rows ?? 0); r++) {
    for (let c = 0; c < Math.max(width, previous?.columns ?? 0); c++) {
      if (!(r < rows.length && c < width) && !owned(r, c)) continue
      const key = cellKey(origin.row + r, origin.column + c)
      if (isCellLocked(sheet, key)) throw new Error(`${key} is protected.`)
      if (!owned(r, c) && cells[key]?.value) throw new Error(`${key} already contains data. Make room for the pivot result first.`)
      const cell = rows[r]?.[c]
      if (cell?.value || cell?.style) cells[key] = cell
      else delete cells[key]
    }
  }
  return { ...sheet, cells, rowCount: Math.max(sheet.rowCount, origin.row + rows.length), columnCount: Math.max(sheet.columnCount, origin.column + width) }
}

export function pivotRows(selection: ReturnType<typeof readSelection>, options: Pick<PivotDefinition, 'row' | 'column' | 'value' | 'summary'>): SheetCell[][] {
  const [header, ...body] = selection.rows
  const { row, column, value, summary } = options
  if (!header[row] || !header[value] || (column !== null && (!header[column] || column === row)) || !body.length)
    throw new Error('Choose valid fields and include a header and at least one data row.')
  const groups = new Map<string, Map<string, { sum: number; numeric: number; count: number }>>()
  const headings = new Set<string>(column === null ? [`${summary === 'sum' ? 'Sum' : summary === 'count' ? 'Count' : 'Average'} of ${header[value].value || 'values'}`] : [])
  for (const entry of body) {
    if (entry.every(cell => cell.value === '')) continue
    const label = entry[row].value
    const heading = column === null ? [...headings][0] : entry[column].value
    headings.add(heading)
    if (!groups.has(label)) groups.set(label, new Map())
    const group = groups.get(label)!
    const stat = group.get(heading) ?? { sum: 0, numeric: 0, count: 0 }
    const datum = entry[value]
    if (datum.value !== '') stat.count++
    if (datum.value !== '' && datum.cell?.kind !== 'text' && Number.isFinite(Number(datum.value))) {
      stat.sum += Number(datum.value)
      stat.numeric++
    }
    group.set(heading, stat)
    if ((groups.size + 1) * (headings.size + 1) > MAX_ANALYSIS_CELLS)
      throw new Error('This pivot has too many groups. Choose a smaller source or fewer distinct column values.')
  }
  const text = (value: string): SheetCell => ({ kind: 'text', value })
  return [
    [text(header[row].value || 'Group'), ...[...headings].map(label => text(label || '(blank)'))],
    ...[...groups].map(([label, group]) => [
      text(label || '(blank)'),
      ...[...headings].map((heading): SheetCell => {
        const stat = group.get(heading)
        if (!stat || (summary === 'average' && !stat.numeric)) return { kind: 'blank', value: '' }
        return { kind: 'number', value: String(summary === 'count' ? stat.count : summary === 'average' ? stat.sum / stat.numeric : stat.sum) }
      }),
    ]),
  ]
}

export function normalizePivot(input: unknown): PivotDefinition | undefined {
  if (!input || typeof input !== 'object') return
  const item = input as Partial<PivotDefinition>
  const index = (value: unknown): value is number => Number.isInteger(value) && Number(value) >= 0
  if (typeof item.sourceSheetId !== 'string' || typeof item.range !== 'string' || !normalizeRangeToken(item.range) ||
    typeof item.output !== 'string' || !parseCellKey(item.output) || !index(item.row) || !index(item.value) ||
    (item.column !== null && !index(item.column)) || !index(item.rows) || !index(item.columns) ||
    item.rows < 1 || item.columns < 1 || item.rows * item.columns > MAX_ANALYSIS_CELLS || !['sum', 'count', 'average'].includes(item.summary ?? '')) return
  return item as PivotDefinition
}

type CellReference = { sheetId: string; range: string }
const referenceTokens = /"(?:[^"]|"")*"|(?<![\w.])(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(?:\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?[1-9]\d{0,6}:\$?[1-9]\d{0,6})(?![\w.]|\s*\()/gi

function formulaReferences(workbook: Workbook, sheet: WorkbookSheet, formula: string): CellReference[] {
  const expanded = expandRangeNames(formula, workbook.namedRanges ?? [], workbook.sheets)
  return [...expanded.matchAll(referenceTokens)].flatMap(([token]) => {
    if (token.startsWith('"')) return []
    const split = token.lastIndexOf('!')
    const name = token.slice(0, split).replace(/^'|'$/g, '').replace(/''/g, "'")
    const target = split < 0 ? sheet : workbook.sheets.find(sheet => sheet.name.toLowerCase() === name.toLowerCase())
    if (!target) return []
    const address = token.slice(split + 1).replace(/\$/g, '')
    const columns = /^([A-Z]+):([A-Z]+)$/i.exec(address)
    const rows = /^(\d+):(\d+)$/.exec(address)
    const range = normalizeRangeToken(columns ? `${columns[1]}1:${columns[2]}${target.rowCount}` :
      rows ? `A${rows[1]}:${columnLabel(target.columnCount - 1)}${rows[2]}` : address)
    return range ? [{ sheetId: target.id, range: range.replace(/^(.+):\1$/, '$1') }] : []
  })
}

export function traceFormulaErrors(workbook: Workbook, references: CellReference[]): CellReference[] {
  const found: CellReference[] = []
  const seen = new Set<string>()
  for (const reference of references) {
    const sheet = workbook.sheets.find(sheet => sheet.id === reference.sheetId)
    if (!sheet) continue
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (cell.kind !== 'formula' || !rangesIntersect(reference.range, key)) continue
      const id = sheet.id + '!' + key
      if (seen.has(id) || !isFormulaError(computedCellValue(sheet, key, workbook))) continue
      seen.add(id)
      found.push({ sheetId: sheet.id, range: key })
      if (found.length >= 50) return found
    }
  }
  return found
}

export function auditFormula(workbook: Workbook, sheet: WorkbookSheet, key: string) {
  const cell = sheet.cells[key]
  const precedents = cell?.kind === 'formula' ? formulaReferences(workbook, sheet, cell.value) : []
  const position = parseCellKey(key)
  const dependents: CellReference[] = []
  if (position) {
    for (const target of workbook.sheets) {
      for (const [candidate, cell] of Object.entries(target.cells)) {
        if (cell.kind !== 'formula' || !formulaReferences(workbook, target, cell.value).some(ref => {
          if (ref.sheetId !== sheet.id) return false
          const [start, end = start] = ref.range.split(':').map(parseCellKey)
          return start && end && position.row >= start.row && position.row <= end.row &&
            position.column >= start.column && position.column <= end.column
        })) continue
        dependents.push({ sheetId: target.id, range: candidate })
        if (dependents.length >= 200) return { precedents, dependents, limited: true }
      }
    }
  }
  return { precedents, dependents, limited: false }
}

export function goalSeek(workbook: Workbook, sheet: WorkbookSheet, formula: string, variable: string, target: number) {
  if (sheet.cells[formula]?.kind !== 'formula') throw new Error('Choose a formula cell.')
  const original = sheet.cells[variable]
  if (original?.kind !== 'number' || !Number.isFinite(Number(original.value)) || isCellLocked(sheet, variable))
    throw new Error('Choose an unlocked cell that contains a number.')
  if (!Number.isFinite(target)) throw new Error('Enter a finite target value.')
  const calculate = (value: number): number => {
    const changed = { ...sheet, cells: { ...sheet.cells, [variable]: { ...original, value: String(value) } } }
    const next = { ...workbook, sheets: workbook.sheets.map(sheet => sheet.id === changed.id ? changed : sheet) }
    const result = computedCellValue(changed, formula, next)
    return result === '' || isFormulaError(result) ? NaN : Number(result)
  }
  const initial = Number(original.value)
  const initialResult = calculate(initial)
  if (!Number.isFinite(initialResult)) throw new Error('The formula must evaluate to a number.')
  const tolerance = 1e-8 * Math.max(1, Math.abs(target))
  if (Math.abs(initialResult - target) <= tolerance) return { initial, value: initial, result: initialResult }
  for (const scale of [0.01, -0.01, 1, -1, 10, -10]) {
    let before = initial
    let oldResult = initialResult
    let value = initial + Math.max(Math.abs(initial), 1) * scale
    for (let step = 0; step < 32 && Number.isFinite(value) && Math.abs(value) <= 1e15; step++) {
      const result = calculate(value)
      if (!Number.isFinite(result)) break
      if (Math.abs(result - target) <= tolerance) return { initial, value, result }
      if (result === oldResult) break
      const next = value - (result - target) * ((value - before) / (result - oldResult))
      before = value
      oldResult = result
      value = next
    }
  }
  throw new Error('No solution was found from the current input. Try a different starting value or a reachable target.')
}

export function selectionFields(selection: ReturnType<typeof readSelection>) {
  return selection.rows[0].map(({ value }, index) => value || `Column ${columnLabel(selection.first.column + index)}`)
}
