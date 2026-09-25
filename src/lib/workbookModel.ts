import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT } from '../constants'
import type {
  AgentCellChange,
  AgentLogCellChange,
  AgentLogEntry,
  CellPosition,
  CellValueKind,
  PureSheetsDocument,
  SheetCell,
  WorkbookChart,
  WorkbookSheet,
} from '../types'
import { withFreshUniverSnapshot } from './univerAdapter'

const columnLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
let sheetIdCounter = 1

export function createBlankWorkbook(title = 'Untitled'): PureSheetsDocument {
  const now = new Date().toISOString()
  const sheet = createBlankSheet('sheet-1', 'Sheet 1')
  return {
    app: 'PureSheets',
    version: 1,
    engine: {
      name: 'univer',
    },
    metadata: {
      title,
      createdAt: now,
      updatedAt: now,
    },
    workbook: {
      activeSheetId: sheet.id,
      sheets: [sheet],
    },
  }
}

export function createBlankSheet(id: string, name: string): WorkbookSheet {
  return {
    id,
    name,
    rowCount: DEFAULT_ROW_COUNT,
    columnCount: DEFAULT_COLUMN_COUNT,
    cells: {},
  }
}

export function addSheet(
  document: PureSheetsDocument,
  name = `Sheet ${document.workbook.sheets.length + 1}`,
): PureSheetsDocument {
  const sheet = createBlankSheet(nextSheetId(), name)
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      activeSheetId: sheet.id,
      sheets: [...document.workbook.sheets, sheet],
    },
  })
}

export function renameSheet(
  document: PureSheetsDocument,
  sheetId: string,
  name: string,
): PureSheetsDocument {
  const trimmed = name.trim()
  if (!trimmed) return document
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      sheets: document.workbook.sheets.map(sheet =>
        sheet.id === sheetId ? { ...sheet, name: trimmed } : sheet,
      ),
    },
  })
}

export function duplicateSheet(
  document: PureSheetsDocument,
  sheetId: string,
): PureSheetsDocument {
  const source = document.workbook.sheets.find(sheet => sheet.id === sheetId)
  if (!source) return document
  const copy: WorkbookSheet = {
    ...source,
    id: nextSheetId(),
    name: `${source.name} copy`,
    // A duplicate is an ordinary sheet — it must not claim to be an agent
    // draft of the original's source (QA operations only track the original).
    agentDraft: undefined,
    cells: Object.fromEntries(
      Object.entries(source.cells).map(([key, cell]) => [
        key,
        { ...cell, style: cell.style ? { ...cell.style } : undefined },
      ]),
    ),
    charts: source.charts?.map(chart => ({
      ...chart,
      id: nextChartId(),
      labels: [...chart.labels],
      values: [...chart.values],
    })),
  }
  return touchDocument({
    ...document,
    workbook: {
      activeSheetId: copy.id,
      sheets: [...document.workbook.sheets, copy],
    },
  })
}

export function deleteSheet(
  document: PureSheetsDocument,
  sheetId: string,
): PureSheetsDocument {
  if (document.workbook.sheets.length <= 1) return document
  const sheets = document.workbook.sheets.filter(sheet => sheet.id !== sheetId)
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      activeSheetId:
        document.workbook.activeSheetId === sheetId
          ? sheets[0].id
          : document.workbook.activeSheetId,
      sheets,
    },
  })
}

export function cellKey(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`
}

export function parseCellKey(key: string): CellPosition | null {
  const match = /^([A-Z]+)([1-9][0-9]*)$/i.exec(key.trim())
  if (!match) return null
  return {
    row: Number(match[2]) - 1,
    column: columnIndex(match[1].toUpperCase()),
  }
}

export function columnLabel(index: number): string {
  let value = index + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = columnLetters[remainder] + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

export function columnIndex(label: string): number {
  return (
    label.split('').reduce((total, letter) => {
      return total * 26 + columnLetters.indexOf(letter) + 1
    }, 0) - 1
  )
}

export function normalizeRangeToken(rangeToken: string): string | null {
  const trimmed = rangeToken.trim()
  const singleCell = parseCellKey(trimmed)
  if (singleCell) {
    const key = cellKey(singleCell.row, singleCell.column)
    return `${key}:${key}`
  }
  const range = parseRange(trimmed)
  return range
    ? `${cellKey(range.start.row, range.start.column)}:${cellKey(
        range.end.row,
        range.end.column,
      )}`
    : null
}

export function cellsInRange(rangeToken: string): string[] {
  const bounds = parseRange(rangeToken)
  if (!bounds) return []
  const keys: string[] = []
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    for (
      let column = bounds.start.column;
      column <= bounds.end.column;
      column += 1
    ) {
      keys.push(cellKey(row, column))
    }
  }
  return keys
}

export function upsertCell(
  sheet: WorkbookSheet,
  key: string,
  value: string,
): WorkbookSheet {
  // Protection (Phase S3): locked cells are never mutated at the model layer.
  if (isCellLocked(sheet, key)) return sheet
  const trimmed = value.trim()
  const nextCells = { ...sheet.cells }
  if (!trimmed) {
    delete nextCells[key]
  } else {
    nextCells[key] = {
      value,
      kind: trimmed.startsWith('=')
        ? 'formula'
        : numberLike(trimmed)
        ? 'number'
        : 'text',
      style: sheet.cells[key]?.style,
    }
  }
  return {
    ...sheet,
    cells: nextCells,
  }
}

export function clearRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const keys = cellsInRange(rangeToken)
  if (!keys.length) return sheet
  const nextCells = { ...sheet.cells }
  for (const key of keys) {
    if (isCellLocked(sheet, key)) continue
    delete nextCells[key]
  }
  return {
    ...sheet,
    cells: nextCells,
  }
}

// ---------------------------------------------------------------------------
// Sheet protection (Phase S3).
//
// Locked ranges are enforced HERE, in the mutation paths (upsertCell,
// clearRange, pasteCellBlock, fillRange, sortRange), not in the UI — this is
// the substrate the S5 agent-edit invariant leans on: locking a source sheet
// makes every model-level write a no-op regardless of who asks.
// ---------------------------------------------------------------------------

export function lockRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const normalized = normalizeRangeToken(rangeToken)
  if (!normalized) return sheet
  const existing = sheet.protection?.lockedRanges ?? []
  if (existing.includes(normalized)) return sheet
  return {
    ...sheet,
    protection: { lockedRanges: [...existing, normalized] },
  }
}

/** Remove every locked range that intersects the given range. */
export function unlockRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const normalized = normalizeRangeToken(rangeToken)
  const bounds = normalized ? parseRange(normalized) : null
  const existing = sheet.protection?.lockedRanges ?? []
  if (!bounds || !existing.length) return sheet
  const remaining = existing.filter(locked => {
    const lockedBounds = parseRange(locked)
    if (!lockedBounds) return false
    return !boundsIntersect(bounds, lockedBounds)
  })
  if (remaining.length === existing.length) return sheet
  return {
    ...sheet,
    protection: remaining.length ? { lockedRanges: remaining } : undefined,
  }
}

export function isCellLocked(sheet: WorkbookSheet, key: string): boolean {
  const lockedRanges = sheet.protection?.lockedRanges
  if (!lockedRanges?.length) return false
  const position = parseCellKey(key)
  if (!position) return false
  return lockedRanges.some(locked => {
    const bounds = parseRange(locked)
    if (!bounds) return false
    return (
      position.row >= bounds.start.row &&
      position.row <= bounds.end.row &&
      position.column >= bounds.start.column &&
      position.column <= bounds.end.column
    )
  })
}

/** True when any cell in the range is locked. */
export function isRangeLocked(
  sheet: WorkbookSheet,
  rangeToken: string,
): boolean {
  return cellsInRange(rangeToken).some(key => isCellLocked(sheet, key))
}

function boundsIntersect(
  a: { start: CellPosition; end: CellPosition },
  b: { start: CellPosition; end: CellPosition },
): boolean {
  return (
    a.start.row <= b.end.row &&
    a.end.row >= b.start.row &&
    a.start.column <= b.end.column &&
    a.end.column >= b.start.column
  )
}

// ---------------------------------------------------------------------------
// Cell comments/notes (Phase S3). Keyed by cell; keys shift with structural
// edits like the cells themselves.
// ---------------------------------------------------------------------------

export function setCellComment(
  sheet: WorkbookSheet,
  key: string,
  text: string,
): WorkbookSheet {
  const trimmed = text.trim()
  const next = { ...sheet.comments }
  if (!trimmed) delete next[key]
  else next[key] = trimmed
  return {
    ...sheet,
    comments: Object.keys(next).length ? next : undefined,
  }
}

// ---------------------------------------------------------------------------
// Filter query (Phase S3): optional contains-filter on the filter column.
// ---------------------------------------------------------------------------

export function setFilterQuery(
  sheet: WorkbookSheet,
  query: string,
): WorkbookSheet {
  const trimmed = query.trim()
  if ((sheet.filterQuery ?? '') === trimmed) return sheet
  return { ...sheet, filterQuery: trimmed || undefined }
}

// ---------------------------------------------------------------------------
// Find (Phase S3): ordered matches across the sheet, optional case matching.
// ---------------------------------------------------------------------------

export function findMatches(
  sheet: WorkbookSheet,
  query: string,
  options: { matchCase?: boolean } = {},
): string[] {
  const matchCase = options.matchCase ?? false
  const needle = matchCase ? query.trim() : query.trim().toLowerCase()
  if (!needle) return []
  return Object.keys(sheet.cells)
    .sort((left, right) => {
      const a = parseCellKey(left)
      const b = parseCellKey(right)
      if (!a || !b) return left.localeCompare(right)
      return a.row === b.row ? a.column - b.column : a.row - b.row
    })
    .filter(key => {
      const raw = sheet.cells[key]?.value ?? ''
      const display = displayCellValue(sheet, key)
      const haystackRaw = matchCase ? raw : raw.toLowerCase()
      const haystackDisplay = matchCase ? display : display.toLowerCase()
      return haystackRaw.includes(needle) || haystackDisplay.includes(needle)
    })
}

/**
 * Replace only the FIRST occurrence of `findText` within a single cell — so a
 * "Replace" click steps through matches one at a time instead of clobbering
 * every occurrence in the cell at once (#277). Returns `replaced: false` when
 * the cell is missing, locked, or has no match.
 */
export function replaceFirstOccurrence(
  sheet: WorkbookSheet,
  cellKey: string,
  findText: string,
  replaceText: string,
  options: { matchCase?: boolean } = {},
): { sheet: WorkbookSheet; replaced: boolean } {
  if (!findText || isCellLocked(sheet, cellKey)) return { sheet, replaced: false }
  const cell = sheet.cells[cellKey]
  if (!cell) return { sheet, replaced: false }
  // Non-global regex ⇒ String.replace swaps only the first match.
  const pattern = new RegExp(escapeRegExp(findText), options.matchCase ? '' : 'i')
  if (!pattern.test(cell.value)) return { sheet, replaced: false }
  const nextValue = cell.value.replace(pattern, replaceText)
  if (nextValue === cell.value) return { sheet, replaced: false }
  return { sheet: upsertCell(sheet, cellKey, nextValue), replaced: true }
}

/** Replace across the whole used sheet (locked cells are skipped). */
export function replaceAllInSheet(
  sheet: WorkbookSheet,
  findText: string,
  replaceText: string,
  options: { matchCase?: boolean } = {},
): { sheet: WorkbookSheet; count: number } {
  if (!findText) return { sheet, count: 0 }
  const matchCase = options.matchCase ?? false
  const pattern = new RegExp(escapeRegExp(findText), matchCase ? 'g' : 'gi')
  let count = 0
  let nextSheet = sheet
  for (const key of Object.keys(sheet.cells)) {
    if (isCellLocked(sheet, key)) continue
    const cell = nextSheet.cells[key]
    if (!cell) continue
    const occurrences = cell.value.match(pattern)?.length ?? 0
    if (!occurrences) continue
    count += occurrences
    nextSheet = upsertCell(
      nextSheet,
      key,
      cell.value.replace(pattern, () => replaceText),
    )
  }
  return { sheet: nextSheet, count }
}

/** Public computed value (formula results, unformatted) for one cell. */
export function computedCellValue(sheet: WorkbookSheet, key: string): string {
  return rawCellValue(sheet, key)
}

function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

export function setCellStyle(
  sheet: WorkbookSheet,
  key: string,
  patch: Partial<NonNullable<SheetCell['style']>>,
): WorkbookSheet {
  const current = sheet.cells[key] ?? { value: '', kind: 'blank' as const }
  return {
    ...sheet,
    cells: {
      ...sheet.cells,
      [key]: {
        ...current,
        style: {
          ...current.style,
          ...patch,
        },
      },
    },
  }
}

export function setRangeStyle(
  sheet: WorkbookSheet,
  rangeToken: string,
  patch: Partial<NonNullable<SheetCell['style']>>,
): WorkbookSheet {
  const keys = cellsInRange(rangeToken)
  if (!keys.length) return sheet
  return keys.reduce(
    (nextSheet, key) => setCellStyle(nextSheet, key, patch),
    sheet,
  )
}

// Table-style presets (Phase S2). Quiet neutrals from the enterprise palette;
// everything is expressed through ordinary cell styles so the presets
// round-trip through the snapshot bridge like any other formatting.
const TABLE_HEADER_FILL = '#eef1eb'
const TABLE_BAND_FILL = '#f5f7f2'

export type TableStylePreset =
  | 'header-row'
  | 'banded-rows'
  | 'totals-row'
  | 'grid-borders'

export function applyTableStylePreset(
  sheet: WorkbookSheet,
  rangeToken: string,
  preset: TableStylePreset,
): WorkbookSheet {
  const normalized = normalizeRangeToken(rangeToken)
  const bounds = normalized ? parseRange(normalized) : null
  if (!bounds) return sheet
  const rowRange = (row: number): string =>
    `${cellKey(row, bounds.start.column)}:${cellKey(row, bounds.end.column)}`
  if (preset === 'grid-borders') {
    return setRangeStyle(sheet, normalized as string, { border: true })
  }
  if (preset === 'header-row') {
    return setRangeStyle(sheet, rowRange(bounds.start.row), {
      bold: true,
      fillColor: TABLE_HEADER_FILL,
      border: true,
    })
  }
  if (preset === 'totals-row') {
    return setRangeStyle(sheet, rowRange(bounds.end.row), {
      bold: true,
      fillColor: TABLE_HEADER_FILL,
      border: true,
    })
  }
  // banded-rows: the first row is assumed to be a header and left alone;
  // every other data row gets the band fill.
  let next = sheet
  for (let row = bounds.start.row + 1; row <= bounds.end.row; row += 1) {
    if ((row - bounds.start.row) % 2 === 1) {
      next = setRangeStyle(next, rowRange(row), { fillColor: TABLE_BAND_FILL })
    }
  }
  return next
}

export function rangeToDelimitedText(
  sheet: WorkbookSheet,
  rangeToken: string,
): string {
  const bounds = parseRange(rangeToken)
  if (!bounds) return ''
  const rows: string[] = []
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    const values: string[] = []
    for (
      let column = bounds.start.column;
      column <= bounds.end.column;
      column += 1
    ) {
      values.push(displayCellValue(sheet, cellKey(row, column)))
    }
    rows.push(values.join('\t'))
  }
  return rows.join('\n')
}

export function replaceInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
  findText: string,
  replaceText: string,
  options: { matchCase?: boolean } = {},
): { sheet: WorkbookSheet; count: number } {
  if (!findText) return { sheet, count: 0 }
  const keys = cellsInRange(rangeToken)
  if (!keys.length) return { sheet, count: 0 }
  const matchCase = options.matchCase ?? true
  const pattern = new RegExp(escapeRegExp(findText), matchCase ? 'g' : 'gi')
  let count = 0
  let nextSheet = sheet
  for (const key of keys) {
    if (isCellLocked(sheet, key)) continue
    const cell = nextSheet.cells[key]
    if (!cell) continue
    const occurrences = cell.value.match(pattern)?.length ?? 0
    if (!occurrences) continue
    count += occurrences
    nextSheet = upsertCell(
      nextSheet,
      key,
      cell.value.replace(pattern, () => replaceText),
    )
  }
  return { sheet: nextSheet, count }
}

export function setColumnWidth(
  sheet: WorkbookSheet,
  column: number,
  width: number,
): WorkbookSheet {
  const nextWidths = { ...sheet.columnWidths }
  nextWidths[String(column)] = clampDimension(width, 56, 360)
  return {
    ...sheet,
    columnWidths: nextWidths,
  }
}

export function setRowHeight(
  sheet: WorkbookSheet,
  row: number,
  height: number,
): WorkbookSheet {
  const nextHeights = { ...sheet.rowHeights }
  nextHeights[String(row)] = clampDimension(height, 24, 160)
  return {
    ...sheet,
    rowHeights: nextHeights,
  }
}

/**
 * Rewrite every cell reference in a formula (skipping quoted strings).
 * The callback returns the replacement reference text, or null to mark the
 * reference broken (#REF!).
 */
function mapFormulaRefs(
  formula: string,
  map: (ref: {
    row: number
    column: number
    absRow: boolean
    absColumn: boolean
  }) => { row: number; column: number } | null,
): string {
  return formula
    .split('"')
    .map((segment, index) => {
      if (index % 2 === 1) return segment // inside a quoted string
      return segment.replace(
        /(\$?)([A-Z]+)(\$?)(\d+)/gi,
        (
          token,
          absColumn: string,
          columnText: string,
          absRow: string,
          rowText: string,
        ) => {
          const position = parseCellKey(`${columnText}${rowText}`)
          if (!position) return token
          const mapped = map({
            row: position.row,
            column: position.column,
            absRow: absRow === '$',
            absColumn: absColumn === '$',
          })
          if (!mapped) return '#REF!'
          if (mapped.row < 0 || mapped.column < 0) return '#REF!'
          return `${absColumn}${columnLabel(mapped.column)}${absRow}${
            mapped.row + 1
          }`
        },
      )
    })
    .join('"')
}

/** Apply a formula rewrite to every formula cell in a cell map. */
function adjustFormulaCells(
  cells: Record<string, SheetCell>,
  adjust: (formula: string) => string,
): Record<string, SheetCell> {
  const next: Record<string, SheetCell> = {}
  for (const [key, cell] of Object.entries(cells)) {
    if (cell.kind !== 'formula') {
      next[key] = cell
      continue
    }
    const value = adjust(cell.value)
    next[key] = value === cell.value ? cell : { ...cell, value }
  }
  return next
}

function shiftFormulaRefsForAxis(
  formula: string,
  axis: 'row' | 'column',
  target: number,
  delta: 1 | -1,
): string {
  return mapFormulaRefs(formula, ref => {
    const index = axis === 'row' ? ref.row : ref.column
    if (delta === -1 && index === target) return null // deleted target
    const shifted = index >= target ? index + delta : index
    return axis === 'row'
      ? { row: shifted, column: ref.column }
      : { row: ref.row, column: shifted }
  })
}

/**
 * Shift positional metadata (comments keyed by cell; validation, protection
 * and conditional-format ranges) with a structural edit, mirroring how the
 * cells themselves move. Rules that lose their entire range are dropped.
 */
function shiftPositionalMetadata(
  sheet: WorkbookSheet,
  axis: 'row' | 'column',
  target: number,
  delta: 1 | -1,
): Pick<
  WorkbookSheet,
  'comments' | 'validations' | 'protection' | 'conditionalFormats'
> {
  const comments = shiftCellKeyRecord(sheet.comments, axis, target, delta)
  const validations = (sheet.validations ?? []).flatMap(rule => {
    const range = shiftRangeTokenForAxis(rule.range, axis, target, delta)
    return range ? [{ ...rule, range }] : []
  })
  const lockedRanges = (sheet.protection?.lockedRanges ?? []).flatMap(
    locked => {
      const range = shiftRangeTokenForAxis(locked, axis, target, delta)
      return range ? [range] : []
    },
  )
  const conditionalFormats = (sheet.conditionalFormats ?? []).flatMap(rule => {
    const range = shiftRangeTokenForAxis(rule.range, axis, target, delta)
    return range ? [{ ...rule, range }] : []
  })
  return {
    comments,
    validations: validations.length ? validations : undefined,
    protection: lockedRanges.length ? { lockedRanges } : undefined,
    conditionalFormats: conditionalFormats.length
      ? conditionalFormats
      : undefined,
  }
}

function shiftCellKeyRecord(
  record: Record<string, string> | undefined,
  axis: 'row' | 'column',
  target: number,
  delta: 1 | -1,
): Record<string, string> | undefined {
  if (!record) return undefined
  const next: Record<string, string> = {}
  for (const [key, text] of Object.entries(record)) {
    const position = parseCellKey(key)
    if (!position) continue
    const index = axis === 'row' ? position.row : position.column
    if (delta === -1 && index === target) continue // deleted line
    const shifted = index >= target ? index + delta : index
    const nextKey =
      axis === 'row'
        ? cellKey(shifted, position.column)
        : cellKey(position.row, shifted)
    next[nextKey] = text
  }
  return Object.keys(next).length ? next : undefined
}

function shiftRangeTokenForAxis(
  rangeToken: string,
  axis: 'row' | 'column',
  target: number,
  delta: 1 | -1,
): string | null {
  const normalized = normalizeRangeToken(rangeToken)
  const bounds = normalized ? parseRange(normalized) : null
  if (!bounds) return rangeToken
  let start = axis === 'row' ? bounds.start.row : bounds.start.column
  let end = axis === 'row' ? bounds.end.row : bounds.end.column
  if (delta === 1) {
    if (start >= target) start += 1
    if (end >= target) end += 1
  } else {
    if (start === target && end === target) return null
    if (start > target) start -= 1
    if (end >= target) end -= 1
    if (end < start) return null
  }
  return axis === 'row'
    ? `${cellKey(start, bounds.start.column)}:${cellKey(
        end,
        bounds.end.column,
      )}`
    : `${cellKey(bounds.start.row, start)}:${cellKey(bounds.end.row, end)}`
}

export function insertRow(
  sheet: WorkbookSheet,
  rowIndex: number,
): WorkbookSheet {
  const targetRow = clampIndex(rowIndex, sheet.rowCount)
  const nextCells: Record<string, SheetCell> = {}
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position) continue
    const nextRow = position.row >= targetRow ? position.row + 1 : position.row
    nextCells[cellKey(nextRow, position.column)] = cell
  }
  return {
    ...sheet,
    rowCount: sheet.rowCount + 1,
    cells: adjustFormulaCells(nextCells, formula =>
      shiftFormulaRefsForAxis(formula, 'row', targetRow, 1),
    ),
    rowHeights: shiftDimensions(sheet.rowHeights, targetRow, 1),
    ...shiftPositionalMetadata(sheet, 'row', targetRow, 1),
  }
}

export function deleteRow(
  sheet: WorkbookSheet,
  rowIndex: number,
): WorkbookSheet {
  if (sheet.rowCount <= 1) return sheet
  const targetRow = clampIndex(rowIndex, sheet.rowCount - 1)
  const nextCells: Record<string, SheetCell> = {}
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position || position.row === targetRow) continue
    const nextRow = position.row > targetRow ? position.row - 1 : position.row
    nextCells[cellKey(nextRow, position.column)] = cell
  }
  return {
    ...sheet,
    rowCount: sheet.rowCount - 1,
    cells: adjustFormulaCells(nextCells, formula =>
      shiftFormulaRefsForAxis(formula, 'row', targetRow, -1),
    ),
    rowHeights: shiftDimensions(sheet.rowHeights, targetRow, -1),
    ...shiftPositionalMetadata(sheet, 'row', targetRow, -1),
  }
}

export function insertColumn(
  sheet: WorkbookSheet,
  columnIndexValue: number,
): WorkbookSheet {
  const targetColumn = clampIndex(columnIndexValue, sheet.columnCount)
  const nextCells: Record<string, SheetCell> = {}
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position) continue
    const nextColumn =
      position.column >= targetColumn ? position.column + 1 : position.column
    nextCells[cellKey(position.row, nextColumn)] = cell
  }
  return {
    ...sheet,
    columnCount: sheet.columnCount + 1,
    cells: adjustFormulaCells(nextCells, formula =>
      shiftFormulaRefsForAxis(formula, 'column', targetColumn, 1),
    ),
    columnWidths: shiftDimensions(sheet.columnWidths, targetColumn, 1),
    ...shiftPositionalMetadata(sheet, 'column', targetColumn, 1),
  }
}

export function deleteColumn(
  sheet: WorkbookSheet,
  columnIndexValue: number,
): WorkbookSheet {
  if (sheet.columnCount <= 1) return sheet
  const targetColumn = clampIndex(columnIndexValue, sheet.columnCount - 1)
  const nextCells: Record<string, SheetCell> = {}
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position || position.column === targetColumn) continue
    const nextColumn =
      position.column > targetColumn ? position.column - 1 : position.column
    nextCells[cellKey(position.row, nextColumn)] = cell
  }
  return {
    ...sheet,
    columnCount: sheet.columnCount - 1,
    cells: adjustFormulaCells(nextCells, formula =>
      shiftFormulaRefsForAxis(formula, 'column', targetColumn, -1),
    ),
    columnWidths: shiftDimensions(sheet.columnWidths, targetColumn, -1),
    ...shiftPositionalMetadata(sheet, 'column', targetColumn, -1),
  }
}

export function pasteTabularData(
  sheet: WorkbookSheet,
  startCell: string,
  text: string,
): WorkbookSheet {
  const start = parseCellKey(startCell)
  if (!start) return sheet
  return text
    .replace(/\r\n/g, '\n')
    .split('\n')
    .filter((row, index, rows) => row.length > 0 || index < rows.length - 1)
    .reduce((nextSheet, rowText, rowOffset) => {
      return rowText.split('\t').reduce((innerSheet, value, columnOffset) => {
        return upsertCell(
          innerSheet,
          cellKey(start.row + rowOffset, start.column + columnOffset),
          value,
        )
      }, nextSheet)
    }, sheet)
}

// ---------------------------------------------------------------------------
// Clipboard blocks and fill (Phase S1 — core editing parity).
//
// The shell keeps an in-app clipboard block alongside the system-clipboard
// TSV so in-app paste preserves formulas and formatting (TSV only carries
// display values). These are model-layer semantics shared by both surfaces:
// the Univer editor persists the result through the snapshot bridge, and the
// fallback grid applies them directly.
// ---------------------------------------------------------------------------

export type PasteMode = 'all' | 'values' | 'formats'

function cloneCell(cell: SheetCell | undefined): SheetCell | undefined {
  return cell
    ? {
        ...cell,
        style: cell.style ? { ...cell.style } : undefined,
      }
    : undefined
}

function formatTransformedNumber(value: number): string {
  if (!Number.isFinite(value)) return ''
  return Number.isInteger(value)
    ? String(value)
    : String(Number(value.toFixed(10)))
}

export interface CopiedCellBlock {
  rows: number
  columns: number
  /** Origin of the copied range — used for relative formula shifting. */
  sourceStart: CellPosition
  /** Raw cells as copied: formulas and styles preserved. */
  cells: Array<Array<SheetCell | undefined>>
  /** Computed raw values at copy time (formula results, unformatted). */
  computedValues: string[][]
}

export function copyRangeCells(
  sheet: WorkbookSheet,
  rangeToken: string,
): CopiedCellBlock | null {
  const normalized = normalizeRangeToken(rangeToken)
  const bounds = normalized ? parseRange(normalized) : null
  if (!bounds) return null
  const cells: Array<Array<SheetCell | undefined>> = []
  const computedValues: string[][] = []
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    const cellRow: Array<SheetCell | undefined> = []
    const valueRow: string[] = []
    for (
      let column = bounds.start.column;
      column <= bounds.end.column;
      column += 1
    ) {
      const key = cellKey(row, column)
      cellRow.push(cloneCell(sheet.cells[key]))
      valueRow.push(rawCellValue(sheet, key))
    }
    cells.push(cellRow)
    computedValues.push(valueRow)
  }
  return {
    rows: bounds.end.row - bounds.start.row + 1,
    columns: bounds.end.column - bounds.start.column + 1,
    sourceStart: bounds.start,
    cells,
    computedValues,
  }
}

/**
 * Paste a copied block at an anchor cell.
 * - `all`: full cells — relative formula references shift with the paste
 *   offset (absolute `$` references stay put); blank source cells clear the
 *   target, like a spreadsheet paste.
 * - `values`: computed values only (formulas become their results); the
 *   target keeps its own formatting.
 * - `formats`: styles only; target values are untouched. A style-less source
 *   cell clears the target's formatting.
 */
export function pasteCellBlock(
  sheet: WorkbookSheet,
  anchorCell: string,
  block: CopiedCellBlock,
  mode: PasteMode = 'all',
): WorkbookSheet {
  const start = parseCellKey(anchorCell)
  if (!start || !block.rows || !block.columns) return sheet
  const nextCells = { ...sheet.cells }
  const rowDelta = start.row - block.sourceStart.row
  const columnDelta = start.column - block.sourceStart.column
  for (let row = 0; row < block.rows; row += 1) {
    for (let column = 0; column < block.columns; column += 1) {
      const key = cellKey(start.row + row, start.column + column)
      if (isCellLocked(sheet, key)) continue
      const source = block.cells[row]?.[column]
      const existing = nextCells[key]
      if (mode === 'formats') {
        const style = source?.style ? { ...source.style } : undefined
        if (existing) {
          nextCells[key] = style
            ? { ...existing, style }
            : { value: existing.value, kind: existing.kind }
        } else if (style) {
          nextCells[key] = { value: '', kind: 'blank', style }
        }
        continue
      }
      if (mode === 'values') {
        const value = block.computedValues[row]?.[column] ?? ''
        if (!value.trim()) {
          if (existing?.style) {
            nextCells[key] = {
              value: '',
              kind: 'blank',
              style: { ...existing.style },
            }
          } else {
            delete nextCells[key]
          }
          continue
        }
        nextCells[key] = {
          value,
          kind: deriveCellKind(value),
          style: existing?.style ? { ...existing.style } : undefined,
        }
        continue
      }
      // mode === 'all'
      if (!source || (!source.value && !source.style)) {
        delete nextCells[key]
        continue
      }
      const pasted = cloneCell(source) as SheetCell
      if (pasted.kind === 'formula') {
        pasted.value = mapFormulaRefs(pasted.value, ref => ({
          row: ref.absRow ? ref.row : ref.row + rowDelta,
          column: ref.absColumn ? ref.column : ref.column + columnDelta,
        }))
      }
      nextCells[key] = pasted
    }
  }
  return {
    ...sheet,
    rowCount: Math.max(sheet.rowCount, start.row + block.rows),
    columnCount: Math.max(sheet.columnCount, start.column + block.columns),
    cells: nextCells,
  }
}

/**
 * Fill-handle semantics at the model layer: extend a source range down or
 * across into a target range that shares its perpendicular bounds.
 * - All-numeric source lanes continue as a linear series (endpoints fit);
 *   a single number copies.
 * - Formulas tile with relative references shifted by the fill distance.
 * - Text and styles tile unchanged; blanks clear.
 */
export function fillRange(
  sheet: WorkbookSheet,
  sourceToken: string,
  targetToken: string,
): WorkbookSheet {
  const source = parseRange(sourceToken)
  const target = parseRange(targetToken)
  if (!source || !target) return sheet
  const sameColumns =
    source.start.column === target.start.column &&
    source.end.column === target.end.column
  const sameRows =
    source.start.row === target.start.row && source.end.row === target.end.row
  if (!sameColumns && !sameRows) return sheet
  const axis: 'row' | 'column' = sameColumns ? 'row' : 'column'
  const sourceStart = axis === 'row' ? source.start.row : source.start.column
  const sourceEnd = axis === 'row' ? source.end.row : source.end.column
  const targetStart = axis === 'row' ? target.start.row : target.start.column
  const targetEnd = axis === 'row' ? target.end.row : target.end.column
  const laneStart = axis === 'row' ? source.start.column : source.start.row
  const laneEnd = axis === 'row' ? source.end.column : source.end.row
  const sourceLength = sourceEnd - sourceStart + 1
  const keyAt = (lane: number, along: number): string =>
    axis === 'row' ? cellKey(along, lane) : cellKey(lane, along)

  const nextCells = { ...sheet.cells }
  for (let lane = laneStart; lane <= laneEnd; lane += 1) {
    const numbers: number[] = []
    let allNumbers = true
    for (let index = 0; index < sourceLength; index += 1) {
      const cell = sheet.cells[keyAt(lane, sourceStart + index)]
      const numeric = cell?.kind === 'number' ? Number(cell.value) : NaN
      if (Number.isFinite(numeric)) numbers.push(numeric)
      else {
        allNumbers = false
        break
      }
    }
    const step =
      allNumbers && sourceLength >= 2
        ? (numbers[sourceLength - 1] - numbers[0]) / (sourceLength - 1)
        : 0
    for (let along = targetStart; along <= targetEnd; along += 1) {
      if (along >= sourceStart && along <= sourceEnd) continue
      const key = keyAt(lane, along)
      if (isCellLocked(sheet, key)) continue
      const sequenceIndex = along - sourceStart
      const sourceIndex =
        ((sequenceIndex % sourceLength) + sourceLength) % sourceLength
      const sourceCell = sheet.cells[keyAt(lane, sourceStart + sourceIndex)]
      if (allNumbers && numbers.length) {
        const value = formatTransformedNumber(numbers[0] + step * sequenceIndex)
        nextCells[key] = {
          value,
          kind: 'number',
          style: sourceCell?.style ? { ...sourceCell.style } : undefined,
        }
        continue
      }
      if (!sourceCell) {
        delete nextCells[key]
        continue
      }
      const filled = cloneCell(sourceCell) as SheetCell
      if (filled.kind === 'formula') {
        const shift = along - (sourceStart + sourceIndex)
        filled.value = mapFormulaRefs(filled.value, ref =>
          axis === 'row'
            ? {
                row: ref.absRow ? ref.row : ref.row + shift,
                column: ref.column,
              }
            : {
                row: ref.row,
                column: ref.absColumn ? ref.column : ref.column + shift,
              },
        )
      }
      nextCells[key] = filled
    }
  }
  return {
    ...sheet,
    rowCount: Math.max(sheet.rowCount, target.end.row + 1),
    columnCount: Math.max(sheet.columnCount, target.end.column + 1),
    cells: nextCells,
  }
}

export function sortRange(
  sheet: WorkbookSheet,
  rangeToken: string,
  direction: 'asc' | 'desc' = 'asc',
): WorkbookSheet {
  const bounds = parseRange(rangeToken)
  if (!bounds) return sheet
  // Sorting rearranges values — refuse when the range touches locked cells.
  if (isRangeLocked(sheet, rangeToken)) return sheet
  const sortColumn = bounds.start.column
  const rows = []
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    const cells = []
    for (
      let column = bounds.start.column;
      column <= bounds.end.column;
      column += 1
    ) {
      const key = cellKey(row, column)
      cells.push({ key, cell: sheet.cells[key] })
    }
    const blank =
      displayCellValue(sheet, cellKey(row, sortColumn)).trim() === ''
    rows.push({ row, cells, blank })
  }
  // Empty cells are NOT accounted for: rows blank in the sort column keep their
  // position, and only the rows that actually have a value are reordered — they
  // refill the non-blank slots in sorted order (#304).
  const sortedNonBlank = rows
    .filter(entry => !entry.blank)
    .sort((left, right) => {
      const leftValue = displayCellValue(sheet, cellKey(left.row, sortColumn))
      const rightValue = displayCellValue(sheet, cellKey(right.row, sortColumn))
      const leftNumber = Number(leftValue)
      const rightNumber = Number(rightValue)
      const comparison =
        Number.isFinite(leftNumber) && Number.isFinite(rightNumber)
          ? leftNumber - rightNumber
          : leftValue.localeCompare(rightValue)
      return direction === 'asc' ? comparison : -comparison
    })
  let nextNonBlank = 0
  const targetOrder = rows.map(entry =>
    entry.blank ? entry : sortedNonBlank[nextNonBlank++],
  )
  const nextCells = { ...sheet.cells }
  targetOrder.forEach((sourceRow, targetOffset) => {
    const targetRow = bounds.start.row + targetOffset
    const rowDelta = targetRow - sourceRow.row
    sourceRow.cells.forEach(({ cell }, columnOffset) => {
      const targetKey = cellKey(targetRow, bounds.start.column + columnOffset)
      if (cell) {
        // A moved formula behaves like cut/paste: relative row references
        // travel with the row so per-row formulas stay correct after sorting.
        nextCells[targetKey] =
          cell.kind === 'formula' && rowDelta !== 0
            ? {
                ...cell,
                value: mapFormulaRefs(cell.value, ref =>
                  ref.absRow
                    ? { row: ref.row, column: ref.column }
                    : { row: ref.row + rowDelta, column: ref.column },
                ),
              }
            : cell
      } else delete nextCells[targetKey]
    })
  })
  return {
    ...sheet,
    cells: nextCells,
  }
}

export function toggleFreezeFirstRow(sheet: WorkbookSheet): WorkbookSheet {
  return {
    ...sheet,
    frozenRows: sheet.frozenRows ? 0 : 1,
  }
}

export function toggleFreezeFirstColumn(sheet: WorkbookSheet): WorkbookSheet {
  return {
    ...sheet,
    frozenColumns: sheet.frozenColumns ? 0 : 1,
  }
}

export function toggleFilterRow(
  sheet: WorkbookSheet,
  column = 0,
): WorkbookSheet {
  const active = sheet.filterRow === 0
  return {
    ...sheet,
    filterRow: active ? undefined : 0,
    filterColumn: active ? undefined : Math.max(0, Math.floor(column)),
    // Default to Excel-style: the first row is a header and stays visible.
    filterHasHeader: active ? undefined : sheet.filterHasHeader ?? true,
    // Clear the contains query when turning the filter OFF so re-opening the
    // filter starts from an empty field instead of the previous term (#274).
    filterQuery: active ? undefined : sheet.filterQuery,
  }
}

export function setFilterHasHeader(
  sheet: WorkbookSheet,
  hasHeader: boolean,
): WorkbookSheet {
  if ((sheet.filterHasHeader ?? true) === hasHeader) return sheet
  return { ...sheet, filterHasHeader: hasHeader }
}

export function visibleRowsForSheet(sheet: WorkbookSheet): number[] {
  const rows = Array.from({ length: sheet.rowCount }, (_, row) => row)
  if (sheet.filterRow !== 0) return rows
  const filterColumn = Math.max(0, Math.floor(sheet.filterColumn ?? 0))
  const query = sheet.filterQuery?.trim().toLowerCase()
  return rows.filter(row => {
    // The header row stays visible only when the sheet is marked as having one.
    if (row === sheet.filterRow && sheet.filterHasHeader !== false) return true
    const value = displayCellValue(sheet, cellKey(row, filterColumn))
    // With a filter query: contains-match (case-insensitive) on the filter
    // column. Without one: the original non-empty filter.
    if (query) return value.toLowerCase().includes(query)
    return value.trim() !== ''
  })
}

export function createChartFromRange(
  sheet: WorkbookSheet,
  rangeToken: string,
  type: WorkbookChart['type'] = 'bar',
): WorkbookSheet {
  const chart = chartFromRange(sheet, rangeToken, type)
  if (!chart) return sheet
  return {
    ...sheet,
    charts: [...(sheet.charts ?? []), chart],
  }
}

export function removeChart(
  sheet: WorkbookSheet,
  chartId: string,
): WorkbookSheet {
  return {
    ...sheet,
    charts: (sheet.charts ?? []).filter(chart => chart.id !== chartId),
  }
}

/**
 * Chart editor (Phase S4): patch a chart's config. Changing the range or
 * type re-derives the chart's data from the sheet; other fields patch in
 * place. Returns the sheet unchanged when the chart is unknown or a new
 * range yields no plottable data.
 */
export function updateChart(
  sheet: WorkbookSheet,
  chartId: string,
  patch: Partial<
    Pick<
      WorkbookChart,
      'title' | 'type' | 'range' | 'legend' | 'showAxes' | 'xLabel' | 'yLabel'
    >
  >,
): WorkbookSheet {
  const charts = sheet.charts ?? []
  const existing = charts.find(chart => chart.id === chartId)
  if (!existing) return sheet
  let next: WorkbookChart = { ...existing, ...patch }
  const rangeChanged =
    patch.range !== undefined &&
    normalizeRangeToken(patch.range) !== normalizeRangeToken(existing.range)
  const typeChanged = patch.type !== undefined && patch.type !== existing.type
  if (rangeChanged || typeChanged) {
    const range =
      normalizeRangeToken(patch.range ?? existing.range) ?? existing.range
    const derived = chartFromRange(sheet, range, patch.type ?? existing.type)
    if (!derived) return sheet
    next = {
      ...next,
      range,
      labels: derived.labels,
      values: derived.values,
    }
  }
  return {
    ...sheet,
    charts: charts.map(chart => (chart.id === chartId ? next : chart)),
  }
}

/**
 * True when the sheet carries positional metadata (comments, validation,
 * protection, conditional formats). Structural edits on such sheets must run
 * through the model (which shifts the metadata) rather than the Univer
 * bridge, whose snapshot export copies `custom` verbatim and would leave the
 * metadata pointing at the wrong cells.
 */
export function sheetHasPositionalMetadata(sheet: WorkbookSheet): boolean {
  return Boolean(
    (sheet.comments && Object.keys(sheet.comments).length) ||
      sheet.validations?.length ||
      sheet.protection?.lockedRanges.length ||
      sheet.conditionalFormats?.length,
  )
}

/** Append an imported/constructed sheet to the workbook and activate it. */
export function addSheetToWorkbook(
  document: PureSheetsDocument,
  sheet: WorkbookSheet,
): PureSheetsDocument {
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      activeSheetId: sheet.id,
      sheets: [...document.workbook.sheets, sheet],
    },
  })
}

export function displayCellValue(sheet: WorkbookSheet, key: string): string {
  const cell = sheet.cells[key]
  if (!cell) return ''
  return formatCellValue(rawCellValue(sheet, key), cell.style)
}

function chartFromRange(
  sheet: WorkbookSheet,
  rangeToken: string,
  type: WorkbookChart['type'],
): WorkbookChart | null {
  const bounds = parseRange(rangeToken)
  if (!bounds) return null
  const rows: Array<{ label: string; value: number }> = []
  const hasLabelColumn = bounds.end.column > bounds.start.column
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    const labelKey = hasLabelColumn
      ? cellKey(row, bounds.start.column)
      : cellKey(row, bounds.start.column)
    const valueKey = hasLabelColumn
      ? cellKey(row, bounds.start.column + 1)
      : cellKey(row, bounds.start.column)
    const rawText = rawCellValue(sheet, valueKey)
    // A blank cell is not a zero data point (Number('') === 0).
    if (rawText.trim() === '') continue
    const rawValue = Number(rawText)
    if (!Number.isFinite(rawValue)) continue
    rows.push({
      label: hasLabelColumn
        ? displayCellValue(sheet, labelKey) || labelKey
        : valueKey,
      value: rawValue,
    })
  }
  if (!rows.length) return null
  return {
    id: nextChartId(),
    title: `${rangeToken} ${type} chart`,
    type,
    range: rangeToken,
    labels: rows.map(row => row.label),
    values: rows.map(row => row.value),
  }
}

export function isFormulaError(value: string): boolean {
  return /^#[A-Z0-9/?!]+$/i.test(value.trim())
}

// ---------------------------------------------------------------------------
// COMPATIBILITY SHIM (Phase S0, MAIL_SHEETS_MASTER_PLAN.md).
//
// Live editing formulas are evaluated by Univer's formula engine inside
// `UniverSpreadsheetSurface`; this hand-written evaluator is NOT the engine
// behind the visible editor anymore. It remains only for surfaces that read
// the workbook model outside Univer:
//   - the legacy CompatibilityGrid (vitest runs / mount-failure fallback),
//   - displayCellValue consumers: QA history text, chart extraction
//     (chartFromRange), sortRange comparisons, find/VLOOKUP-on-model paths,
//   - reading/converting legacy `.sheets` files.
// Do not add new spreadsheet functions here — extend Univer instead. Delete
// this evaluator only after Phase S2 verifies engine parity (golden tests).
// ---------------------------------------------------------------------------
// Recursion guard. Circular references (e.g. =A1 in A1) previously crashed
// with a stack overflow. Cycles are detected precisely by the set of cells
// currently on the evaluation stack (see rawCellValue) — that set returns
// #REF! the moment a cell re-enters its own evaluation, so a genuinely deep
// but *acyclic* dependency chain evaluates correctly. The depth counter below
// is only a last-ditch stack-overflow backstop set well above any realistic
// chain length; correctness no longer depends on it.
const evaluatingCells = new Set<string>()
let formulaEvaluationDepth = 0
const MAX_FORMULA_EVALUATION_DEPTH = 2048

export function evaluateFormula(sheet: WorkbookSheet, formula: string): string {
  if (formulaEvaluationDepth >= MAX_FORMULA_EVALUATION_DEPTH) return '#REF!'
  formulaEvaluationDepth += 1
  try {
    const expression = formula.trim().replace(/^=/, '')
    if (!expression) return '#FORMULA'
    const result = evaluateExpressionString(sheet, expression)
    return result ?? '#FORMULA'
  } finally {
    formulaEvaluationDepth -= 1
  }
}

/**
 * Resolve one argument of a numeric aggregate to its cell values. Plain cell
 * and range references expand to their raw values; anything else (literals,
 * nested function calls like `AVERAGE(C1:C3)` inside `ROUND(...)`, arithmetic
 * expressions) is evaluated through the full expression parser — previously
 * nested calls silently collapsed to their literal text and were dropped
 * (caught by the engine-agnostic golden suite against Univer's engine).
 */
function numericArgValues(sheet: WorkbookSheet, arg: string): string[] {
  const normalized = arg.replace(/\$/g, '')
  if (/^[A-Z]+\d+(?::[A-Z]+\d+)?$/i.test(normalized)) {
    return expandRange(sheet, arg)
  }
  const evaluated = evaluateExpressionString(sheet, arg)
  return [evaluated ?? unquote(arg)]
}

function applyFormulaFunction(
  sheet: WorkbookSheet,
  fn: string,
  rawArgs: string[],
): string {
  const values = rawArgs
    .flatMap(part => numericArgValues(sheet, part.trim()))
    // A blank cell is not a zero: without this filter, Number('') === 0 made
    // COUNT/AVERAGE/MIN/MAX over sparse ranges count phantom zeros (caught by
    // the engine-agnostic golden suite against Univer's engine).
    .filter(value => value.trim() !== '')
    .map(value => Number(value))
    .filter(value => Number.isFinite(value))
  const textValues = rawArgs.flatMap(part => expandRange(sheet, part.trim()))

  if (fn === 'SUM')
    return String(values.reduce((total, value) => total + value, 0))
  if (fn === 'AVERAGE') {
    return values.length
      ? String(
          values.reduce((total, value) => total + value, 0) / values.length,
        )
      : '#DIV/0!'
  }
  if (fn === 'MIN') return values.length ? String(Math.min(...values)) : '0'
  if (fn === 'MAX') return values.length ? String(Math.max(...values)) : '0'
  if (fn === 'COUNT') return String(values.length)
  if (fn === 'COUNTA')
    return String(textValues.filter(value => value !== '').length)
  if (fn === 'IF') {
    const [condition, truthy = '', falsy = ''] = rawArgs
    return truthyCondition(sheet, condition)
      ? resolveTextArg(sheet, truthy)
      : resolveTextArg(sheet, falsy)
  }
  if (fn === 'ROUND') {
    const [value, places = 0] = values
    return Number.isFinite(value)
      ? String(Number(value.toFixed(places)))
      : '#VALUE'
  }
  if (fn === 'CONCAT')
    return rawArgs.map(arg => resolveTextArg(sheet, arg)).join('')
  if (fn === 'LEFT') {
    const [value = '', count = '1'] = rawArgs
    return resolveTextArg(sheet, value).slice(
      0,
      Math.max(0, Number(resolveTextArg(sheet, count)) || 0),
    )
  }
  if (fn === 'RIGHT') {
    const [value = '', count = '1'] = rawArgs
    const text = resolveTextArg(sheet, value)
    return text.slice(
      Math.max(0, text.length - (Number(resolveTextArg(sheet, count)) || 0)),
    )
  }
  if (fn === 'VLOOKUP') return vlookup(sheet, rawArgs)
  if (fn === 'TODAY') return new Date().toISOString().slice(0, 10)
  if (fn === 'NOW') return new Date().toISOString()
  if (fn === 'COUNTIF') return countIf(sheet, rawArgs)
  if (fn === 'SUMIF') return sumIf(sheet, rawArgs)
  if (fn === 'AVERAGEIF') return averageIf(sheet, rawArgs)
  if (fn === 'RANK') return rank(sheet, rawArgs)
  if (fn === 'IFERROR') {
    const [tryArg = '', fallback = ''] = rawArgs
    const result = resolveTextArg(sheet, tryArg)
    return result.startsWith('#') ? resolveTextArg(sheet, fallback) : result
  }
  return '#NAME?'
}

// ---------------------------------------------------------------------------
// Expression parser: numbers, strings, cell refs (with $), ranges, function
// calls, parentheses, unary minus, + - * / precedence, and & concatenation.
// ---------------------------------------------------------------------------

interface ExpressionToken {
  type: 'number' | 'string' | 'ref' | 'range' | 'func' | 'op'
  text: string
  args?: string
}

function tokenizeExpression(input: string): ExpressionToken[] | null {
  const tokens: ExpressionToken[] = []
  let index = 0
  while (index < input.length) {
    const ch = input[index]
    if (/\s/.test(ch)) {
      index += 1
      continue
    }
    if ('+-*/&()'.includes(ch)) {
      tokens.push({ type: 'op', text: ch })
      index += 1
      continue
    }
    if (ch === '"') {
      const end = input.indexOf('"', index + 1)
      if (end === -1) return null
      tokens.push({ type: 'string', text: input.slice(index + 1, end) })
      index = end + 1
      continue
    }
    const numberMatch = /^\d+(?:\.\d+)?/.exec(input.slice(index))
    if (numberMatch) {
      tokens.push({ type: 'number', text: numberMatch[0] })
      index += numberMatch[0].length
      continue
    }
    const errorMatch = /^#[A-Z0-9/?!]+/i.exec(input.slice(index))
    if (errorMatch) {
      // Error literal (e.g. #REF! left behind by a deleted row) propagates
      tokens.push({ type: 'string', text: errorMatch[0].toUpperCase() })
      index += errorMatch[0].length
      continue
    }
    const rangeMatch = /^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+/i.exec(
      input.slice(index),
    )
    if (rangeMatch) {
      tokens.push({ type: 'range', text: rangeMatch[0].replace(/\$/g, '') })
      index += rangeMatch[0].length
      continue
    }
    const refMatch = /^\$?[A-Z]+\$?\d+(?![A-Z0-9(])/i.exec(input.slice(index))
    if (refMatch) {
      tokens.push({ type: 'ref', text: refMatch[0].replace(/\$/g, '') })
      index += refMatch[0].length
      continue
    }
    const funcMatch = /^([A-Z][A-Z0-9]*)\(/i.exec(input.slice(index))
    if (funcMatch) {
      // Scan to the matching close paren, honouring quotes and nesting
      let depth = 0
      let quoted = false
      let end = -1
      for (
        let scan = index + funcMatch[1].length;
        scan < input.length;
        scan += 1
      ) {
        const scanCh = input[scan]
        if (scanCh === '"') quoted = !quoted
        if (quoted) continue
        if (scanCh === '(') depth += 1
        if (scanCh === ')') {
          depth -= 1
          if (depth === 0) {
            end = scan
            break
          }
        }
      }
      if (end === -1) return null
      tokens.push({
        type: 'func',
        text: funcMatch[1].toUpperCase(),
        args: input.slice(index + funcMatch[1].length + 1, end),
      })
      index = end + 1
      continue
    }
    return null
  }
  return tokens
}

function evaluateExpressionString(
  sheet: WorkbookSheet,
  expression: string,
): string | null {
  const tokenized = tokenizeExpression(expression)
  if (!tokenized || tokenized.length === 0) return null
  const tokens: ExpressionToken[] = tokenized
  let position = 0

  const peek = (): ExpressionToken | undefined => tokens[position]
  const isOp = (text: string): boolean =>
    tokens[position]?.type === 'op' && tokens[position]?.text === text

  function parsePrimary(): string | null {
    const token = peek()
    if (!token) return null
    if (token.type === 'op' && token.text === '(') {
      position += 1
      const inner = parseConcat()
      if (inner === null || !isOp(')')) return null
      position += 1
      return inner
    }
    if (token.type === 'op' && token.text === '-') {
      position += 1
      const operand = parsePrimary()
      if (operand === null) return null
      if (isFormulaError(operand)) return operand
      const numeric = Number(operand)
      return Number.isFinite(numeric) ? String(-numeric) : '#VALUE'
    }
    position += 1
    if (token.type === 'number' || token.type === 'string') return token.text
    if (token.type === 'ref') {
      const positionKey = parseCellKey(token.text)
      return positionKey
        ? rawCellValue(sheet, cellKey(positionKey.row, positionKey.column))
        : '#REF!'
    }
    if (token.type === 'range') {
      // A bare range outside a function argument resolves to its first cell
      return expandRange(sheet, token.text)[0] ?? ''
    }
    if (token.type === 'func') {
      return applyFormulaFunction(
        sheet,
        token.text,
        splitFormulaArguments(token.args ?? ''),
      )
    }
    return null
  }

  function applyNumericOp(left: string, op: string, right: string): string {
    if (isFormulaError(left)) return left
    if (isFormulaError(right)) return right
    const a = Number(left)
    const b = Number(right)
    if (!Number.isFinite(a) || !Number.isFinite(b)) return '#VALUE'
    if (op === '*') return String(a * b)
    if (op === '/') return b === 0 ? '#DIV/0!' : String(a / b)
    if (op === '-') return String(a - b)
    return String(a + b)
  }

  function parseMultiplicative(): string | null {
    let left = parsePrimary()
    if (left === null) return null
    while (isOp('*') || isOp('/')) {
      const op = tokens[position].text
      position += 1
      const right = parsePrimary()
      if (right === null) return null
      left = applyNumericOp(left, op, right)
    }
    return left
  }

  function parseAdditive(): string | null {
    let left = parseMultiplicative()
    if (left === null) return null
    while (isOp('+') || isOp('-')) {
      const op = tokens[position].text
      position += 1
      const right = parseMultiplicative()
      if (right === null) return null
      left = applyNumericOp(left, op, right)
    }
    return left
  }

  function parseConcat(): string | null {
    let left = parseAdditive()
    if (left === null) return null
    while (isOp('&')) {
      position += 1
      const right = parseAdditive()
      if (right === null) return null
      if (isFormulaError(left)) return left
      if (isFormulaError(right)) return right
      left = `${left}${right}`
    }
    return left
  }

  const result = parseConcat()
  if (result === null || position < tokens.length) return null
  return result
}

/**
 * Resolve a criteria argument. Handles quoted literals (">100"), bare values,
 * and concatenated expressions like ">"&G2 (the standard dynamic-criteria
 * idiom), which evaluate to e.g. ">355000".
 */
function resolveCriteriaArg(sheet: WorkbookSheet, arg: string): string {
  const trimmed = arg.trim()
  if (trimmed.includes('&')) {
    const evaluated = evaluateExpressionString(sheet, trimmed)
    if (evaluated !== null) return evaluated
  }
  return unquote(trimmed)
}

function countIf(sheet: WorkbookSheet, args: string[]): string {
  const [rangeArg = '', criteriaArg = ''] = args
  const cellValues = expandRange(sheet, rangeArg.trim())
  const criteria = resolveCriteriaArg(sheet, criteriaArg)
  const count = cellValues.filter(v => matchesCriteria(v, criteria)).length
  return String(count)
}

function sumIf(sheet: WorkbookSheet, args: string[]): string {
  const [rangeArg = '', criteriaArg = '', sumRangeArg] = args
  const rangeKeys = cellKeysInRange(rangeArg.trim())
  const criteria = resolveCriteriaArg(sheet, criteriaArg)
  const sumKeys = sumRangeArg ? cellKeysInRange(sumRangeArg.trim()) : rangeKeys
  let total = 0
  rangeKeys.forEach((key, idx) => {
    if (matchesCriteria(rawCellValue(sheet, key), criteria)) {
      total += Number(rawCellValue(sheet, sumKeys[idx] ?? key)) || 0
    }
  })
  return String(total)
}

function averageIf(sheet: WorkbookSheet, args: string[]): string {
  const [rangeArg = '', criteriaArg = '', avgRangeArg] = args
  const rangeKeys = cellKeysInRange(rangeArg.trim())
  const criteria = resolveCriteriaArg(sheet, criteriaArg)
  const avgKeys = avgRangeArg ? cellKeysInRange(avgRangeArg.trim()) : rangeKeys
  const matchedValues: number[] = []
  rangeKeys.forEach((key, idx) => {
    if (matchesCriteria(rawCellValue(sheet, key), criteria)) {
      const v = Number(rawCellValue(sheet, avgKeys[idx] ?? key))
      if (Number.isFinite(v)) matchedValues.push(v)
    }
  })
  if (!matchedValues.length) return '#DIV/0!'
  return String(matchedValues.reduce((a, b) => a + b, 0) / matchedValues.length)
}

function rank(sheet: WorkbookSheet, args: string[]): string {
  const [valueArg = '', refArg = ''] = args
  // order arg: 0 or omitted = descending, 1 = ascending
  const order = Number(resolveTextArg(sheet, args[2] ?? '0')) || 0
  const target = Number(resolveTextArg(sheet, valueArg))
  if (!Number.isFinite(target)) return '#VALUE'
  const refValues = expandRange(sheet, refArg.replace(/\$/g, '').trim())
    .map(Number)
    .filter(Number.isFinite)
  const rankVal =
    order === 0
      ? refValues.filter(v => v > target).length + 1
      : refValues.filter(v => v < target).length + 1
  return String(rankVal)
}

function matchesCriteria(value: string, criteria: string): boolean {
  // An empty cell must never satisfy a numeric criterion. Previously
  // `Number('') === 0` made every blank cell match `>=0`, `<50`, `=0`, etc.,
  // silently inflating COUNTIF/SUMIF/AVERAGEIF over any sparse column.
  const isEmpty = value.trim() === ''
  const opMatch = /^(>=|<=|<>|>|<)(.+)$/.exec(criteria)
  if (opMatch) {
    if (isEmpty) return false
    const [, op, right] = opMatch
    const l = Number(value),
      r = Number(right)
    if (Number.isFinite(l) && Number.isFinite(r)) {
      switch (op) {
        case '>':
          return l > r
        case '<':
          return l < r
        case '>=':
          return l >= r
        case '<=':
          return l <= r
        case '<>':
          return l !== r
      }
    }
    return false
  }
  // Equality: "" matches only empty cells; otherwise compare numerically when
  // both sides are numbers (so `5` matches `5.0`), else case-insensitive text.
  if (criteria.trim() === '') return isEmpty
  if (isEmpty) return false
  const lv = Number(value),
    rv = Number(criteria)
  if (Number.isFinite(lv) && Number.isFinite(rv)) return lv === rv
  return value.toLowerCase() === criteria.toLowerCase()
}

// Returns cell key strings (A1, B2...) for every cell in a range token
function cellKeysInRange(rangeToken: string): string[] {
  const normalised = rangeToken.replace(/\$/g, '')
  const range = /^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$/i.exec(normalised)
  if (!range) {
    const pos = parseCellKey(normalised)
    return pos ? [cellKey(pos.row, pos.column)] : []
  }
  const start = parseCellKey(range[1])
  const end = parseCellKey(range[2])
  if (!start || !end) return []
  const keys: string[] = []
  for (
    let row = Math.min(start.row, end.row);
    row <= Math.max(start.row, end.row);
    row++
  ) {
    for (
      let col = Math.min(start.column, end.column);
      col <= Math.max(start.column, end.column);
      col++
    ) {
      keys.push(cellKey(row, col))
    }
  }
  return keys
}

export function cloneWithUpdatedSheet(
  document: PureSheetsDocument,
  sheet: WorkbookSheet,
): PureSheetsDocument {
  // Update the workbook FIRST, then refresh the engine snapshot from the
  // result. Spreading `...touchDocument(document)` and overriding `workbook`
  // afterward regenerated `engine.snapshot` from the OLD workbook, leaving the
  // snapshot out of sync with the updated sheet. Bridge-driven edits hid this
  // (the canvas was already correct), but model-only features that rely on the
  // snapshot to reach Univer — conditional formatting — silently lost their
  // data on the surface reload (#270).
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      sheets: document.workbook.sheets.map(item =>
        item.id === sheet.id ? sheet : item,
      ),
    },
  })
}

// Trust boundary for model-produced cell diffs. Keys must parse as real cell
// references and land inside the sheet or a bounded growth margin (generate
// operations legitimately add new rows/columns just past the current grid);
// the batch size is capped so a runaway response cannot flood a sheet.
const MAX_AGENT_CELL_CHANGES = 2000
const AGENT_CELL_GROWTH_MARGIN_ROWS = 200
const AGENT_CELL_GROWTH_MARGIN_COLUMNS = 50
const MAX_AGENT_CELL_VALUE_LENGTH = 32_768

function sanitizeAgentCellChanges(
  cellChanges: Array<{
    cell: string
    value: string
    kind: 'text' | 'number' | 'formula' | 'blank'
  }>,
  sheet: WorkbookSheet,
): Array<{
  cell: string
  value: string
  kind: 'text' | 'number' | 'formula' | 'blank'
}> {
  const maxRow = sheet.rowCount + AGENT_CELL_GROWTH_MARGIN_ROWS
  const maxColumn = sheet.columnCount + AGENT_CELL_GROWTH_MARGIN_COLUMNS
  const seen = new Set<string>()
  const accepted: Array<{
    cell: string
    value: string
    kind: 'text' | 'number' | 'formula' | 'blank'
  }> = []
  for (const change of cellChanges) {
    if (accepted.length >= MAX_AGENT_CELL_CHANGES) break
    const position = parseCellKey(change.cell ?? '')
    if (!position) continue
    if (position.row >= maxRow || position.column >= maxColumn) continue
    if (typeof change.value !== 'string') continue
    if (change.value.length > MAX_AGENT_CELL_VALUE_LENGTH) continue
    const key = cellKey(position.row, position.column)
    if (seen.has(key)) continue
    seen.add(key)
    // Never trust the model's self-declared kind: a response could label a
    // "=SUM(...)" payload as text (bypassing formula handling) or tag a literal
    // value as "formula". Derive it from the value so the agent log and
    // persistence see the truth.
    accepted.push({ ...change, cell: key, kind: deriveCellKind(change.value) })
  }
  return accepted
}

function expandRange(sheet: WorkbookSheet, token: string): string[] {
  const normalizedToken = token.replace(/\$/g, '')
  const range = /^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$/i.exec(normalizedToken)
  if (!range) {
    const position = parseCellKey(normalizedToken)
    if (!position) return [unquote(token)]
    return [rawCellValue(sheet, cellKey(position.row, position.column))]
  }

  const start = parseCellKey(range[1])
  const end = parseCellKey(range[2])
  if (!start || !end) return []
  const values: string[] = []
  for (
    let row = Math.min(start.row, end.row);
    row <= Math.max(start.row, end.row);
    row += 1
  ) {
    for (
      let column = Math.min(start.column, end.column);
      column <= Math.max(start.column, end.column);
      column += 1
    ) {
      values.push(rawCellValue(sheet, cellKey(row, column)))
    }
  }
  return values
}

function parseRange(
  rangeToken: string,
): { start: CellPosition; end: CellPosition } | null {
  const range = /^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$/i.exec(
    rangeToken.trim().replace(/\$/g, ''),
  )
  if (!range) return null
  const start = parseCellKey(range[1])
  const end = parseCellKey(range[2])
  if (!start || !end) return null
  return {
    start: {
      row: Math.min(start.row, end.row),
      column: Math.min(start.column, end.column),
    },
    end: {
      row: Math.max(start.row, end.row),
      column: Math.max(start.column, end.column),
    },
  }
}

function splitFormulaArguments(input: string): string[] {
  const args: string[] = []
  let current = ''
  let quoted = false
  let depth = 0
  for (const character of input) {
    if (character === '"') quoted = !quoted
    if (!quoted && character === '(') depth += 1
    if (!quoted && character === ')') depth = Math.max(0, depth - 1)
    if (character === ',' && !quoted && depth === 0) {
      args.push(current.trim())
      current = ''
    } else {
      current += character
    }
  }
  if (current.trim() || input.endsWith(',')) args.push(current.trim())
  return args
}

function resolveTextArg(sheet: WorkbookSheet, arg: string): string {
  const trimmed = arg.trim()
  const evaluated = evaluateExpressionString(sheet, trimmed)
  if (evaluated !== null) return evaluated
  const expanded = expandRange(sheet, trimmed)
  return expanded[0] ?? ''
}

function truthyCondition(sheet: WorkbookSheet, condition: string): boolean {
  const comparison = /^(.+?)(>=|<=|<>|=|>|<)(.+)$/.exec(condition.trim())
  if (!comparison) {
    const value = resolveTextArg(sheet, condition)
    return Boolean(value) && value !== '0' && value.toLowerCase() !== 'false'
  }
  const left = resolveTextArg(sheet, comparison[1])
  const right = resolveTextArg(sheet, comparison[3])
  const leftNumber = Number(left)
  const rightNumber = Number(right)
  const compareAsNumbers =
    Number.isFinite(leftNumber) && Number.isFinite(rightNumber)
  const a = compareAsNumbers ? leftNumber : left
  const b = compareAsNumbers ? rightNumber : right
  switch (comparison[2]) {
    case '>':
      return compareValues(a, b) > 0
    case '<':
      return compareValues(a, b) < 0
    case '>=':
      return compareValues(a, b) >= 0
    case '<=':
      return compareValues(a, b) <= 0
    case '<>':
      return a !== b
    default:
      return a === b
  }
}

function compareValues(left: number | string, right: number | string): number {
  if (typeof left === 'number' && typeof right === 'number') return left - right
  return String(left).localeCompare(String(right))
}

function vlookup(sheet: WorkbookSheet, args: string[]): string {
  const [lookupArg, rangeArg, columnArg] = args
  const bounds = parseRange(rangeArg ?? '')
  const targetColumnOffset = Math.max(
    0,
    (Number(resolveTextArg(sheet, columnArg ?? '1')) || 1) - 1,
  )
  if (!bounds) return '#VALUE'
  const lookupValue = resolveTextArg(sheet, lookupArg ?? '')
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    if (
      displayCellValue(sheet, cellKey(row, bounds.start.column)) === lookupValue
    ) {
      return displayCellValue(
        sheet,
        cellKey(row, bounds.start.column + targetColumnOffset),
      )
    }
  }
  return '#N/A'
}

function unquote(value: string): string {
  const trimmed = value.trim()
  if (trimmed.startsWith('"') && trimmed.endsWith('"'))
    return trimmed.slice(1, -1)
  return trimmed
}

function touchDocument(document: PureSheetsDocument): PureSheetsDocument {
  return withFreshUniverSnapshot({
    ...document,
    metadata: {
      ...document.metadata,
      updatedAt: new Date().toISOString(),
    },
  })
}

function formatCellValue(value: string, style: SheetCell['style']): string {
  if (!style?.numberFormat || style.numberFormat === 'text') return value
  if (value.startsWith('#')) return value
  const numeric = Number(value)
  const decimals = style.decimals
  if (style.numberFormat === 'number') {
    if (!Number.isFinite(numeric)) return value
    return decimals !== undefined ? numeric.toFixed(decimals) : String(numeric)
  }
  if (style.numberFormat === 'currency') {
    return Number.isFinite(numeric)
      ? new Intl.NumberFormat('en-US', {
          style: 'currency',
          currency: 'USD',
          ...(decimals !== undefined
            ? {
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals,
              }
            : {}),
        }).format(numeric)
      : value
  }
  if (style.numberFormat === 'percent') {
    return Number.isFinite(numeric)
      ? new Intl.NumberFormat('en-US', {
          style: 'percent',
          ...(decimals !== undefined
            ? {
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals,
              }
            : { maximumFractionDigits: 2 }),
        }).format(numeric)
      : value
  }
  if (style.numberFormat === 'date') {
    const date = new Date(value)
    return Number.isNaN(date.getTime())
      ? value
      : date.toISOString().slice(0, 10)
  }
  return value
}

function rawCellValue(sheet: WorkbookSheet, key: string): string {
  const cell = sheet.cells[key]
  if (!cell) return ''
  if (cell.kind !== 'formula') return cell.value
  // Cycle detection: if this cell is already being evaluated further up the
  // stack, following it again would loop forever. Return #REF! instead. A
  // long acyclic chain (A1→A2→…→A100) never re-enters the same key, so it is
  // not affected — unlike the old fixed depth cap, which false-#REF!'d it.
  if (evaluatingCells.has(key)) return '#REF!'
  evaluatingCells.add(key)
  try {
    return evaluateFormula(sheet, cell.value)
  } finally {
    evaluatingCells.delete(key)
  }
}

function clampDimension(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min
  return Math.min(max, Math.max(min, Math.round(value)))
}

function clampIndex(value: number, max: number): number {
  if (!Number.isFinite(value)) return 0
  return Math.min(max, Math.max(0, Math.floor(value)))
}

function shiftDimensions(
  dimensions: Record<string, number> | undefined,
  target: number,
  direction: 1 | -1,
): Record<string, number> | undefined {
  const next: Record<string, number> = {}
  for (const [key, value] of Object.entries(dimensions ?? {})) {
    const index = Number(key)
    if (!Number.isInteger(index)) continue
    if (direction > 0) {
      next[String(index >= target ? index + 1 : index)] = value
    } else if (index < target) {
      next[String(index)] = value
    } else if (index > target) {
      next[String(index - 1)] = value
    }
  }
  return Object.keys(next).length ? next : undefined
}

function nextSheetId(): string {
  sheetIdCounter += 1
  return `sheet-${Date.now()}-${sheetIdCounter}`
}

function nextChartId(): string {
  sheetIdCounter += 1
  return `chart-${Date.now()}-${sheetIdCounter}`
}

function numberLike(value: string): boolean {
  return value !== '' && Number.isFinite(Number(value))
}

// The single source of truth for a cell's kind: derive it from the raw value
// rather than trusting a caller-supplied label. Untrusted producers (the agent
// diff pipeline, persisted files) can claim any kind; deriving here prevents a
// "=SUM(...)" text cell or a "123"-valued "formula" cell from slipping through.
export function deriveCellKind(value: string): CellValueKind {
  const trimmed = value.trim()
  if (!trimmed) return 'blank'
  if (trimmed.startsWith('=')) return 'formula'
  return numberLike(trimmed) ? 'number' : 'text'
}

// ---------------------------------------------------------------------------
// Agent log (direct-write model): agent tool writes apply immediately and are
// recorded here with before/after. The log — not a review queue — is the
// accountability mechanism; approval gates are the shell's permissions.
// ---------------------------------------------------------------------------

const AGENT_LOG_CAP = 200
let agentLogCounter = 0

function nextAgentLogId(): string {
  agentLogCounter += 1
  return `agent-log-${Date.now()}-${agentLogCounter}`
}

export function appendAgentLog(
  document: PureSheetsDocument,
  entry: Omit<AgentLogEntry, 'id' | 'at'>,
): PureSheetsDocument {
  const record: AgentLogEntry = {
    ...entry,
    id: nextAgentLogId(),
    at: new Date().toISOString(),
  }
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      agentLog: [record, ...(document.workbook.agentLog ?? [])].slice(
        0,
        AGENT_LOG_CAP,
      ),
    },
  })
}

export interface DirectAgentWriteResult {
  document: PureSheetsDocument
  applied: number
  skippedLocked: number
}

/**
 * The agent cell-write path: sanitize through the trust boundary above,
 * apply straight to the target sheet (locked cells are skipped, never
 * mutated), and record one agent-log entry carrying every cell's
 * before/after. The log — not a review queue — is the accountability
 * mechanism for agent writes.
 */
export function applyAgentCellChangesDirect(
  document: PureSheetsDocument,
  sheetId: string,
  range: string,
  summary: string,
  changes: AgentCellChange[],
  agentName = 'PureSheets Assistant',
): DirectAgentWriteResult {
  const sheet = document.workbook.sheets.find(item => item.id === sheetId)
  if (!sheet) return { document, applied: 0, skippedLocked: 0 }
  const accepted = sanitizeAgentCellChanges(changes, sheet)
  let nextSheet = sheet
  let skippedLocked = 0
  const logged: AgentLogCellChange[] = []
  for (const change of accepted) {
    if (isCellLocked(nextSheet, change.cell)) {
      skippedLocked += 1
      continue
    }
    const before = nextSheet.cells[change.cell]?.value
    const candidate = upsertCell(nextSheet, change.cell, change.value)
    if (candidate === nextSheet) continue
    nextSheet = candidate
    logged.push({
      cell: change.cell,
      ...(before !== undefined ? { before } : {}),
      after: change.value,
      kind: change.kind,
    })
  }
  if (!logged.length) return { document, applied: 0, skippedLocked }
  let nextDocument = cloneWithUpdatedSheet(document, nextSheet)
  nextDocument = appendAgentLog(nextDocument, {
    agentName,
    tool: 'applyCellChanges',
    summary,
    sheetId: sheet.id,
    sheetName: sheet.name,
    range,
    cells: logged,
  })
  return { document: nextDocument, applied: logged.length, skippedLocked }
}

/**
 * Direct chart creation for agent tools, logged. Returns the unchanged
 * document when the range cannot chart (no labels/values).
 */
export function addChartToSheet(
  document: PureSheetsDocument,
  sheetId: string,
  rangeToken: string,
  type: WorkbookChart['type'],
  summary: string,
  agentName = 'PureSheets Assistant',
): PureSheetsDocument {
  const sheet = document.workbook.sheets.find(item => item.id === sheetId)
  if (!sheet) return document
  const range = normalizeRangeToken(rangeToken)
  if (!range) return document
  const chart = chartFromRange(sheet, range, type)
  if (!chart) return document
  const nextSheet: WorkbookSheet = {
    ...sheet,
    charts: [...(sheet.charts ?? []), chart],
  }
  let nextDocument = cloneWithUpdatedSheet(document, nextSheet)
  nextDocument = appendAgentLog(nextDocument, {
    agentName,
    tool: 'addChart',
    summary,
    sheetId: sheet.id,
    sheetName: sheet.name,
    range,
  })
  return nextDocument
}
