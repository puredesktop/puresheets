import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT } from '../constants'
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from './sheetMetadataNormalize'
import { extractChartData } from './chartEngine'
import { resolveNamedRangesInFormula } from './formulaCatalog'
import { visibleRowsWithFilters } from './dataTransforms'
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
import { createWorkbookUnitId, withFreshUniverSnapshot } from './univerAdapter'

const MAX_RANGE_EXPANSION_CELLS = 50_000

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
      snapshot: { id: createWorkbookUnitId() },
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

/**
 * Rewrite sheet-qualified references (Costs!A1, 'Q1/Q2'!A1) in a formula
 * according to a rename map (`oldName -> newName`), leaving double-quoted
 * string literals untouched.
 */
export function renameSheetReferences(
  formula: string,
  names: Map<string, string>,
): string {
  const renamed = new Map(
    [...names]
      .filter(([from, to]) => from !== to)
      .map(([from, to]) => [from.toLowerCase(), to]),
  )
  if (!renamed.size) return formula
  const needsQuotes = (name: string): boolean => {
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) return true
    if (/^(?:TRUE|FALSE)$/i.test(name)) return true
    // Cell-address-like sheet names (Q1, A1, FY2025, R1C1) must be single-quoted.
    if (/^[A-Za-z]{1,3}[1-9][0-9]*$/i.test(name)) return true
    if (/^R\d+C\d+$/i.test(name)) return true
    return false
  }
  const quote = (name: string) =>
    needsQuotes(name) ? `'${name.replace(/'/g, "''")}'` : name
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, index) =>
      index % 2 === 1
        ? part
        : part.replace(
            /'((?:[^']|'')+)'!|(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)!/g,
            (whole, quoted?: string, bare?: string) => {
              const name =
                quoted !== undefined ? quoted.replace(/''/g, "'") : bare ?? ''
              const target = renamed.get(name.toLowerCase())
              return target ? `${quote(target)}!` : whole
            },
          ),
    )
    .join('')
}

export function renameSheet(
  document: PureSheetsDocument,
  sheetId: string,
  name: string,
): PureSheetsDocument {
  const trimmed = name.trim()
  if (!trimmed) return document
  const existing = document.workbook.sheets.find(sheet => sheet.id === sheetId)
  if (!existing) return document
  if (existing.name === trimmed) return document
  if (
    document.workbook.sheets.some(
      sheet =>
        sheet.id !== sheetId &&
        sheet.name.toLowerCase() === trimmed.toLowerCase(),
    )
  ) {
    return document
  }
  const renameMap = new Map([[existing.name, trimmed]])
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      sheets: document.workbook.sheets.map(sheet => ({
        ...sheet,
        name: sheet.id === sheetId ? trimmed : sheet.name,
        cells: adjustFormulaCells(sheet.cells, formula =>
          renameSheetReferences(formula, renameMap),
        ),
      })),
    },
  })
}

export function duplicateSheet(
  document: PureSheetsDocument,
  sheetId: string,
): PureSheetsDocument {
  const source = document.workbook.sheets.find(sheet => sheet.id === sheetId)
  if (!source) return document
  const existingNames = new Set(
    document.workbook.sheets.map(sheet => sheet.name.toLowerCase()),
  )
  let copyName = `${source.name} copy`
  if (existingNames.has(copyName.toLowerCase())) {
    let suffix = 2
    while (existingNames.has(`${source.name} copy ${suffix}`.toLowerCase())) {
      suffix += 1
    }
    copyName = `${source.name} copy ${suffix}`
  }
  const copy: WorkbookSheet = {
    ...source,
    id: nextSheetId(),
    name: copyName,
    // A duplicate is an ordinary sheet — it must not claim to be an agent
    // draft of the original's source (QA operations only track the original).
    agentDraft: undefined,
    cells: Object.fromEntries(
      Object.entries(source.cells).map(([key, cell]) => [
        key,
        { ...cell, style: cell.style ? { ...cell.style } : undefined },
      ]),
    ),
    columnWidths: source.columnWidths ? { ...source.columnWidths } : undefined,
    rowHeights: source.rowHeights ? { ...source.rowHeights } : undefined,
    merges: source.merges ? [...source.merges] : undefined,
    comments: source.comments ? { ...source.comments } : undefined,
    validations: source.validations?.map(rule => ({
      ...rule,
      options: rule.options ? [...rule.options] : undefined,
    })),
    protection: source.protection
      ? { lockedRanges: [...source.protection.lockedRanges] }
      : undefined,
    conditionalFormats: source.conditionalFormats?.map(rule => ({
      ...rule,
      condition: { ...rule.condition },
      style: { ...rule.style },
    })),
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
      ...document.workbook,
      activeSheetId: copy.id,
      sheets: [...document.workbook.sheets, copy],
    },
  })
}

/** Reorder a visible sheet one step left or right among visible sheets. */
export function moveSheet(
  document: PureSheetsDocument,
  sheetId: string,
  direction: 'left' | 'right',
): PureSheetsDocument {
  const sheets = document.workbook.sheets
  const visibleIndices = sheets
    .map((sheet, index) => (sheet.agentDraft ? -1 : index))
    .filter(index => index >= 0)
  const visiblePos = visibleIndices.findIndex(idx => sheets[idx].id === sheetId)
  if (visiblePos < 0) return document
  const neighborPos = direction === 'left' ? visiblePos - 1 : visiblePos + 1
  if (neighborPos < 0 || neighborPos >= visibleIndices.length) return document
  const fromIndex = visibleIndices[visiblePos]
  const toIndex = visibleIndices[neighborPos]
  const nextSheets = [...sheets]
  const [moved] = nextSheets.splice(fromIndex, 1)
  nextSheets.splice(toIndex, 0, moved)
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      sheets: nextSheets,
    },
  })
}

export function deleteSheet(
  document: PureSheetsDocument,
  sheetId: string,
): PureSheetsDocument {
  const target = document.workbook.sheets.find(sheet => sheet.id === sheetId)
  if (!target) return document
  const remainingVisible = document.workbook.sheets.filter(
    sheet => sheet.id !== sheetId && !sheet.agentDraft,
  )
  if (!target.agentDraft && remainingVisible.length === 0) return document
  const sheets = document.workbook.sheets.filter(sheet => sheet.id !== sheetId)
  if (sheets.length === 0) return document
  return touchDocument({
    ...document,
    workbook: {
      ...document.workbook,
      activeSheetId:
        document.workbook.activeSheetId === sheetId
          ? (remainingVisible[0] ?? sheets[0]).id
          : document.workbook.activeSheetId,
      sheets,
    },
  })
}

export function cellKey(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`
}

export function parseCellKey(key: string): CellPosition | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/i.exec(key.trim())
  if (!match) return null
  const row = Number(match[2]) - 1
  const column = columnIndex(match[1].toUpperCase())
  if (
    !Number.isSafeInteger(row) ||
    row < 0 ||
    row >= MAX_SHEET_ROWS ||
    column < 0 ||
    column >= MAX_SHEET_COLUMNS
  ) {
    return null
  }
  return { row, column }
}

/** Public bounding-box parser for an A1 cell or A1:B2 range token (O(1)). */
export function parseRangeBounds(
  rangeToken: string,
): { start: CellPosition; end: CellPosition } | null {
  const trimmed = rangeToken.trim()
  if (!trimmed) return null
  const single = parseCellKey(trimmed.replace(/\$/g, ''))
  if (single) return { start: single, end: single }
  return parseRange(trimmed)
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

export function cellsInRange(rangeToken: string, maxCells = Infinity): string[] {
  const bounds = parseRange(rangeToken)
  if (!bounds) return []
  const count = (bounds.end.row - bounds.start.row + 1) * (bounds.end.column - bounds.start.column + 1)
  if (!Number.isSafeInteger(count) || count > maxCells) return []
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
  const existing = sheet.cells[key]
  if (!trimmed) {
    if (!existing) return sheet
    const nextCells = { ...sheet.cells }
    delete nextCells[key]
    return { ...sheet, cells: nextCells }
  }
  const kind: CellValueKind = trimmed.startsWith('=')
    ? 'formula'
    : numberLike(trimmed)
    ? 'number'
    : 'text'
  if (existing && existing.value === value && existing.kind === kind) {
    return sheet
  }
  const nextCells = {
    ...sheet.cells,
    [key]: {
      value,
      kind,
      style: existing?.style,
    },
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
  const bounds = parseRangeBounds(rangeToken)
  if (!bounds) return sheet
  let changed = false
  const nextCells = { ...sheet.cells }
  for (const key of Object.keys(sheet.cells)) {
    const pos = parseCellKey(key)
    if (
      !pos ||
      pos.row < bounds.start.row ||
      pos.row > bounds.end.row ||
      pos.column < bounds.start.column ||
      pos.column > bounds.end.column
    ) {
      continue
    }
    if (isCellLocked(sheet, key)) continue
    delete nextCells[key]
    changed = true
  }
  if (!changed) return sheet
  return {
    ...sheet,
    cells: nextCells,
  }
}

export function hasMergeInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): boolean {
  const bounds = parseRangeBounds(rangeToken)
  if (!bounds || !sheet.merges?.length) return false
  return sheet.merges.some(merge => {
    const mergeBounds = parseRangeBounds(merge)
    return Boolean(mergeBounds && boundsIntersect(bounds, mergeBounds))
  })
}

export function mergeRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const normalized = normalizeRangeToken(rangeToken)
  const bounds = normalized ? parseRangeBounds(normalized) : null
  if (!normalized || !bounds) return sheet
  if (
    bounds.start.row === bounds.end.row &&
    bounds.start.column === bounds.end.column
  ) {
    return sheet
  }
  if (isRangeLocked(sheet, normalized)) return sheet
  const existing = sheet.merges ?? []
  if (existing.includes(normalized)) return sheet
  const remaining = existing.filter(merge => {
    const mergeBounds = parseRangeBounds(merge)
    return !mergeBounds || !boundsIntersect(bounds, mergeBounds)
  })
  const topLeftKey = cellKey(bounds.start.row, bounds.start.column)
  let cellsChanged = false
  const nextCells = { ...sheet.cells }
  for (const key of Object.keys(sheet.cells)) {
    if (key === topLeftKey) continue
    const pos = parseCellKey(key)
    if (
      !pos ||
      pos.row < bounds.start.row ||
      pos.row > bounds.end.row ||
      pos.column < bounds.start.column ||
      pos.column > bounds.end.column
    ) {
      continue
    }
    const cell = nextCells[key]
    if (cell?.style) {
      if (cell.value !== '' || cell.kind !== 'blank') {
        nextCells[key] = { value: '', kind: 'blank', style: cell.style }
        cellsChanged = true
      }
    } else {
      delete nextCells[key]
      cellsChanged = true
    }
  }
  return {
    ...sheet,
    cells: cellsChanged ? nextCells : sheet.cells,
    merges: [...remaining, normalized],
  }
}

export function unmergeRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const bounds = parseRangeBounds(rangeToken)
  const existing = sheet.merges ?? []
  if (!bounds || !existing.length) return sheet
  if (isRangeLocked(sheet, rangeToken)) return sheet
  const remaining = existing.filter(merge => {
    const mergeBounds = parseRangeBounds(merge)
    return !mergeBounds || !boundsIntersect(bounds, mergeBounds)
  })
  if (remaining.length === existing.length) return sheet
  return {
    ...sheet,
    merges: remaining.length ? remaining : undefined,
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

/** True when any cell in the range is locked (O(L) bounding-box check). */
export function isRangeLocked(
  sheet: WorkbookSheet,
  rangeToken: string,
): boolean {
  const lockedRanges = sheet.protection?.lockedRanges
  if (!lockedRanges?.length) return false
  const bounds = parseRangeBounds(rangeToken)
  if (!bounds) return false
  return lockedRanges.some(locked => {
    const lockedBounds = parseRangeBounds(locked)
    return Boolean(lockedBounds && boundsIntersect(bounds, lockedBounds))
  })
}

export function boundsIntersect(
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

export function isRowDeletionBlockedByProtection(
  sheet: WorkbookSheet,
  rowIndex: number,
): boolean {
  const lockedRanges = sheet.protection?.lockedRanges
  if (!lockedRanges?.length) return false
  for (const locked of lockedRanges) {
    const bounds = parseRangeBounds(locked)
    if (bounds && bounds.start.row <= rowIndex && bounds.end.row >= rowIndex) {
      return true
    }
  }
  return false
}

export function isColumnDeletionBlockedByProtection(
  sheet: WorkbookSheet,
  columnIndexValue: number,
): boolean {
  const lockedRanges = sheet.protection?.lockedRanges
  if (!lockedRanges?.length) return false
  for (const locked of lockedRanges) {
    const bounds = parseRangeBounds(locked)
    if (
      bounds &&
      bounds.start.column <= columnIndexValue &&
      bounds.end.column >= columnIndexValue
    ) {
      return true
    }
  }
  return false
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
  workbook?: FormulaWorkbook,
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
      const display = displayCellValue(sheet, key, workbook)
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
  // Use a replacer function so `$100` / `$$` / `$&` in replaceText are literal.
  const pattern = new RegExp(escapeRegExp(findText), options.matchCase ? '' : 'i')
  if (!pattern.test(cell.value)) return { sheet, replaced: false }
  const nextValue = cell.value.replace(pattern, () => replaceText)
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
/**
 * A cell's computed value. Pass the workbook so a formula can read other sheets (=Costs!B4, ='Q1 plan'!A1:A9);
 * without it, a reference to another sheet is #REF!.
 */
export function computedCellValue(sheet: WorkbookSheet, key: string, workbook?: FormulaWorkbook): string {
  return inWorkbook(workbook, () => rawCellValue(sheet, key))
}

function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function mergeCellStyle(
  currentStyle: SheetCell['style'],
  patch: Partial<NonNullable<SheetCell['style']>>,
): SheetCell['style'] {
  const style: Record<string, unknown> = { ...currentStyle }
  for (const [k, v] of Object.entries(patch)) {
    if (v === undefined) delete style[k]
    else style[k] = v
  }
  // A chosen format or border replaces what an import carried.
  if (('numberFormat' in patch || 'decimals' in patch) && !('numberPattern' in patch)) {
    delete style.numberPattern
  }
  if ('border' in patch && !('borderSides' in patch)) {
    delete style.borderSides
  }
  return Object.keys(style).length ? (style as SheetCell['style']) : undefined
}

export function setCellStyle(
  sheet: WorkbookSheet,
  key: string,
  patch: Partial<NonNullable<SheetCell['style']>>,
): WorkbookSheet {
  const current = sheet.cells[key] ?? { value: '', kind: 'blank' as const }
  const style = mergeCellStyle(current.style, patch)
  const nextCells = { ...sheet.cells }
  if (!style && current.kind === 'blank' && !current.value) {
    delete nextCells[key]
  } else {
    nextCells[key] = { ...current, style }
  }
  return {
    ...sheet,
    cells: nextCells,
  }
}

export function setRangeStyle(
  sheet: WorkbookSheet,
  rangeToken: string,
  patch: Partial<NonNullable<SheetCell['style']>>,
  maxCells = MAX_RANGE_EXPANSION_CELLS,
): WorkbookSheet {
  const keys = cellsInRange(rangeToken, maxCells)
  if (!keys.length) return sheet
  const nextCells = { ...sheet.cells }
  for (const key of keys) {
    const current = nextCells[key] ?? { value: '', kind: 'blank' as const }
    const style = mergeCellStyle(current.style, patch)
    if (!style && current.kind === 'blank' && !current.value) {
      delete nextCells[key]
    } else {
      nextCells[key] = { ...current, style }
    }
  }
  return {
    ...sheet,
    cells: nextCells,
  }
}

/** Clear all formatting (styles) from cells in the given range while preserving values. */
export function clearRangeStyle(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const bounds = parseRangeBounds(rangeToken)
  if (!bounds) return sheet
  let changed = false
  const nextCells = { ...sheet.cells }
  for (const key of Object.keys(sheet.cells)) {
    const pos = parseCellKey(key)
    if (
      !pos ||
      pos.row < bounds.start.row ||
      pos.row > bounds.end.row ||
      pos.column < bounds.start.column ||
      pos.column > bounds.end.column
    ) {
      continue
    }
    const cell = nextCells[key]
    if (!cell?.style) continue
    changed = true
    if (cell.kind === 'blank' && !cell.value) {
      delete nextCells[key]
    } else {
      nextCells[key] = { value: cell.value, kind: cell.kind }
    }
  }
  if (!changed) return sheet
  return { ...sheet, cells: nextCells }
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
  workbook?: FormulaWorkbook,
): string {
  const bounds = parseRangeBounds(rangeToken)
  if (!bounds) return ''
  const rows: string[] = []
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    const values: string[] = []
    for (
      let column = bounds.start.column;
      column <= bounds.end.column;
      column += 1
    ) {
      values.push(displayCellValue(sheet, cellKey(row, column), workbook))
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
 * Apply `transform` only to formula segments outside double-quoted string
 * literals (`"..."`, with `""` escapes). Single-quoted sheet prefixes
 * (`'Q1'!A1`) remain visible to reference mappers so an out-of-bounds shift
 * replaces the entire cross-sheet reference with `#REF!`.
 */
function mapFormulaOutsideLiterals(
  formula: string,
  transform: (segment: string) => string,
): string {
  let result = ''
  let segment = ''
  let index = 0
  while (index < formula.length) {
    const ch = formula[index]
    if (ch === '"') {
      result += transform(segment)
      segment = ''
      let end = index + 1
      while (end < formula.length) {
        if (formula[end] === '"') {
          if (formula[end + 1] === '"') {
            end += 2
          } else {
            break
          }
        } else {
          end += 1
        }
      }
      result += formula.slice(index, end + 1)
      index = end + 1
    } else {
      segment += ch
      index += 1
    }
  }
  return result + transform(segment)
}

/**
 * Rewrite cell references in a formula (skipping `"..."` strings, matching
 * `'Q1'!A1` and `Sheet!A1` prefixes atomically, and avoiding scientific
 * notation like `1E2`).
 * The callback returns the replacement position, or null to mark `#REF!`.
 */
function mapFormulaRefs(
  formula: string,
  map: (ref: {
    row: number
    column: number
    absRow: boolean
    absColumn: boolean
  }) => { row: number; column: number } | null,
  options: { skipCrossSheet?: boolean } = {},
): string {
  const rewriteSingle = (
    absColumn: string,
    columnText: string,
    absRow: string,
    rowText: string,
    fallback: string,
  ): string => {
    const position = parseCellKey(`${columnText}${rowText}`)
    if (!position) return fallback
    const mapped = map({
      row: position.row,
      column: position.column,
      absRow: absRow === '$',
      absColumn: absColumn === '$',
    })
    if (!mapped || mapped.row < 0 || mapped.column < 0) return '#REF!'
    return `${absColumn}${columnLabel(mapped.column)}${absRow}${mapped.row + 1}`
  }

  return mapFormulaOutsideLiterals(formula, segment =>
    segment.replace(
      /(?:('(?:[^']|'')+'!)|(?<![A-Z0-9_.$:])([A-Z_][A-Z0-9_.]*!)|(?<![A-Z0-9_.$:!']))(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![A-Z0-9_(!.])|'(?:[^']|'')+'/gi,
      (
        token,
        quotedSheetPrefix: string | undefined,
        bareSheetPrefix: string | undefined,
        absCol1: string | undefined,
        col1: string | undefined,
        absRow1: string | undefined,
        row1: string | undefined,
        absCol2: string | undefined,
        col2: string | undefined,
        absRow2: string | undefined,
        row2: string | undefined,
      ) => {
        if (col1 === undefined || row1 === undefined) return token
        const sheetPrefix = quotedSheetPrefix ?? bareSheetPrefix
        if (sheetPrefix && options.skipCrossSheet) return token
        const first = rewriteSingle(
          absCol1 ?? '',
          col1,
          absRow1 ?? '',
          row1,
          `${absCol1 ?? ''}${col1}${absRow1 ?? ''}${row1}`,
        )
        if (col2 === undefined || row2 === undefined) {
          return first === '#REF!' ? '#REF!' : `${sheetPrefix ?? ''}${first}`
        }
        const second = rewriteSingle(
          absCol2 ?? '',
          col2,
          absRow2 ?? '',
          row2,
          `${absCol2 ?? ''}${col2}${absRow2 ?? ''}${row2}`,
        )
        if (first === '#REF!' || second === '#REF!') return '#REF!'
        return `${sheetPrefix ?? ''}${first}:${second}`
      },
    ),
  )
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
  return mapFormulaOutsideLiterals(formula, segment =>
    segment.replace(
      /(?:('(?:[^']|'')+'!)|(?<![A-Z0-9_.$:])([A-Z_][A-Z0-9_.]*!)|(?<![A-Z0-9_.$:!']))(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Z0-9_.$(!:])|'(?:[^']|'')+'/gi,
      (
        token,
        quotedSheetPrefix: string | undefined,
        bareSheetPrefix: string | undefined,
        refToken: string | undefined,
      ) => {
        if (refToken === undefined) return token
        // Cross-sheet references (`Sheet!A1` or `'Q1'!A1`) belong to another sheet.
        if (quotedSheetPrefix || bareSheetPrefix) return token
        if (refToken.includes(':')) {
          return shiftRangeTokenForAxis(refToken, axis, target, delta) ?? '#REF!'
        }
        return mapFormulaRefs(refToken, ref => {
          const refIndex = axis === 'row' ? ref.row : ref.column
          if (delta === -1 && refIndex === target) return null
          const shifted = refIndex >= target ? refIndex + delta : refIndex
          return axis === 'row'
            ? { row: shifted, column: ref.column }
            : { row: ref.row, column: shifted }
        })
      },
    ),
  )
}

/**
 * Shift positional metadata (comments keyed by cell; validation, protection,
 * conditional-format, merge, and chart ranges) with a structural edit,
 * mirroring how the cells themselves move. Rules that lose their entire range
 * are dropped.
 */
function shiftPositionalMetadata(
  sheet: WorkbookSheet,
  axis: 'row' | 'column',
  target: number,
  delta: 1 | -1,
): Pick<
  WorkbookSheet,
  | 'comments'
  | 'validations'
  | 'protection'
  | 'conditionalFormats'
  | 'merges'
  | 'charts'
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
  // A merge that shrinks to one cell is no merge.
  const merges = (sheet.merges ?? []).flatMap(merge => {
    const range = shiftRangeTokenForAxis(merge, axis, target, delta)
    if (!range) return []
    const [start, end] = range.split(':')
    return start === end ? [] : [range]
  })
  const charts = sheet.charts?.map(chart => {
    const shifted = shiftRangeTokenForAxis(chart.range, axis, target, delta)
    return shifted && shifted !== chart.range
      ? { ...chart, range: shifted }
      : chart
  })
  return {
    comments,
    merges: merges.length ? merges : undefined,
    validations: validations.length ? validations : undefined,
    protection: lockedRanges.length ? { lockedRanges } : undefined,
    conditionalFormats: conditionalFormats.length
      ? conditionalFormats
      : undefined,
    charts,
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
  const bounds = parseRangeBounds(rangeToken)
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
  const hasColon = rangeToken.includes(':')
  const [originalStart, originalEnd] = rangeToken.trim().split(':')
  const formatRef = (original: string | undefined, row: number, column: number) => {
    const parts = /^(\$?)[A-Z]+(\$?)\d+$/i.exec((original ?? '').trim())
    return `${parts?.[1] ?? ''}${columnLabel(column)}${parts?.[2] ?? ''}${row + 1}`
  }
  if (!hasColon) {
    return axis === 'row'
      ? formatRef(originalStart, start, bounds.start.column)
      : formatRef(originalStart, bounds.start.row, start)
  }
  return axis === 'row'
    ? `${formatRef(originalStart, start, bounds.start.column)}:${formatRef(
        originalEnd,
        end,
        bounds.end.column,
      )}`
    : `${formatRef(originalStart, bounds.start.row, start)}:${formatRef(
        originalEnd,
        bounds.end.row,
        end,
      )}`
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
  if (isRowDeletionBlockedByProtection(sheet, targetRow)) return sheet
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
  if (isColumnDeletionBlockedByProtection(sheet, targetColumn)) return sheet
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
  workbook?: FormulaWorkbook,
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
      valueRow.push(computedCellValue(sheet, key, workbook))
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
        // Paste-values must never promote `=`-prefixed text into a live formula.
        nextCells[key] = {
          value,
          kind: numberLike(value.trim()) ? 'number' : 'text',
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
  workbook?: FormulaWorkbook,
  options: { hasHeader?: boolean } = {},
): WorkbookSheet {
  const bounds = parseRangeBounds(rangeToken)
  if (!bounds) return sheet
  // Sorting rearranges values — refuse when the range touches locked cells.
  if (isRangeLocked(sheet, rangeToken)) return sheet
  const sortStartRow =
    options.hasHeader && bounds.end.row > bounds.start.row
      ? bounds.start.row + 1
      : bounds.start.row
  if (sortStartRow >= bounds.end.row) return sheet
  const sortColumn = bounds.start.column
  const rows = []
  for (let row = sortStartRow; row <= bounds.end.row; row += 1) {
    const cells = []
    for (
      let column = bounds.start.column;
      column <= bounds.end.column;
      column += 1
    ) {
      const key = cellKey(row, column)
      cells.push({
        key,
        cell: sheet.cells[key],
        comment: sheet.comments?.[key],
      })
    }
    const sortValue = computedCellValue(sheet, cellKey(row, sortColumn), workbook)
    const blank = sortValue.trim() === ''
    rows.push({ row, cells, blank, sortValue })
  }
  // Empty cells are NOT accounted for: rows blank in the sort column keep their
  // position, and only the rows that actually have a value are reordered — they
  // refill the non-blank slots in sorted order (#304).
  const sortedNonBlank = rows
    .filter(entry => !entry.blank)
    .sort((left, right) => {
      const leftValue = left.sortValue
      const rightValue = right.sortValue
      const leftNumber = Number(leftValue)
      const rightNumber = Number(rightValue)
      const comparison =
        leftValue.trim() !== '' &&
        rightValue.trim() !== '' &&
        Number.isFinite(leftNumber) &&
        Number.isFinite(rightNumber)
          ? leftNumber - rightNumber
          : leftValue.localeCompare(rightValue)
      return direction === 'asc' ? comparison : -comparison
    })
  let nextNonBlank = 0
  const targetOrder = rows.map(entry =>
    entry.blank ? entry : sortedNonBlank[nextNonBlank++],
  )
  const nextCells = { ...sheet.cells }
  const nextComments = sheet.comments ? { ...sheet.comments } : undefined
  targetOrder.forEach((sourceRow, targetOffset) => {
    const targetRow = sortStartRow + targetOffset
    const rowDelta = targetRow - sourceRow.row
    sourceRow.cells.forEach(({ cell, comment }, columnOffset) => {
      const targetKey = cellKey(targetRow, bounds.start.column + columnOffset)
      if (cell) {
        // A moved formula behaves like cut/paste: relative row references
        // inside the sorted row band travel with the row; cross-sheet refs and
        // refs outside the sorted band stay anchored.
        nextCells[targetKey] =
          cell.kind === 'formula' && rowDelta !== 0
            ? {
                ...cell,
                value: mapFormulaRefs(
                  cell.value,
                  ref =>
                    ref.absRow ||
                    ref.row < sortStartRow ||
                    ref.row > bounds.end.row
                      ? { row: ref.row, column: ref.column }
                      : { row: ref.row + rowDelta, column: ref.column },
                  { skipCrossSheet: true },
                ),
              }
            : cell
      } else delete nextCells[targetKey]
      if (nextComments) {
        if (comment !== undefined) nextComments[targetKey] = comment
        else delete nextComments[targetKey]
      }
    })
  })
  return {
    ...sheet,
    cells: nextCells,
    ...(nextComments !== undefined
      ? {
          comments: Object.keys(nextComments).length ? nextComments : undefined,
        }
      : {}),
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

export function setFreezePanes(
  sheet: WorkbookSheet,
  frozenRows: number,
  frozenColumns: number,
): WorkbookSheet {
  const nextRows = Math.max(0, Math.min(sheet.rowCount, Math.floor(frozenRows)))
  const nextColumns = Math.max(
    0,
    Math.min(sheet.columnCount, Math.floor(frozenColumns)),
  )
  return {
    ...sheet,
    frozenRows: nextRows,
    frozenColumns: nextColumns,
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

export function visibleRowsForSheet(
  sheet: WorkbookSheet,
  workbook?: FormulaWorkbook,
): number[] {
  if (
    (sheet.columnFilters && Object.keys(sheet.columnFilters).length > 0) ||
    (sheet.hiddenRows && sheet.hiddenRows.length > 0)
  ) {
    return visibleRowsWithFilters(sheet, workbook)
  }
  const rows = Array.from({ length: sheet.rowCount }, (_, row) => row)
  if (sheet.filterRow !== 0) return rows
  const filterColumn = Math.max(0, Math.floor(sheet.filterColumn ?? 0))
  const query = sheet.filterQuery?.trim().toLowerCase()
  return rows.filter(row => {
    // The header row stays visible only when the sheet is marked as having one.
    if (row === sheet.filterRow && sheet.filterHasHeader !== false) return true
    const value = displayCellValue(sheet, cellKey(row, filterColumn), workbook)
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
  workbook?: FormulaWorkbook,
): WorkbookSheet {
  const chart = chartFromRange(sheet, rangeToken, type, workbook)
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
  workbook?: FormulaWorkbook,
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
    const derived = chartFromRange(
      sheet,
      range,
      patch.type ?? existing.type,
      workbook,
    )
    if (!derived) return sheet
    next = {
      ...next,
      range,
      labels: derived.labels,
      values: derived.values,
      series: derived.series,
    }
  }
  return {
    ...sheet,
    charts: charts.map(chart => (chart.id === chartId ? next : chart)),
  }
}

/**
 * True when the sheet carries positional metadata (comments, validation,
 * protection, conditional formats, charts). Structural edits on such sheets
 * must run through the model (which shifts the metadata) rather than the Univer
 * bridge, whose snapshot export copies `custom` verbatim and would leave the
 * metadata pointing at the wrong cells.
 */
export function sheetHasPositionalMetadata(sheet: WorkbookSheet): boolean {
  return Boolean(
    (sheet.comments && Object.keys(sheet.comments).length) ||
      sheet.validations?.length ||
      sheet.protection?.lockedRanges.length ||
      sheet.conditionalFormats?.length ||
      sheet.charts?.length,
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

export function displayCellValue(sheet: WorkbookSheet, key: string, workbook?: FormulaWorkbook): string {
  const cell = sheet.cells[key]
  if (!cell) return ''
  return formatCellValue(inWorkbook(workbook, () => rawCellValue(sheet, key)), cell.style)
}

function chartFromRange(
  sheet: WorkbookSheet,
  rangeToken: string,
  type: WorkbookChart['type'],
  workbook?: FormulaWorkbook,
): WorkbookChart | null {
  const bounds = parseRangeBounds(rangeToken)
  if (!bounds) return null
  const colSpan = bounds.end.column - bounds.start.column + 1
  if (colSpan >= 3) {
    const extracted = extractChartData(sheet, rangeToken, type, formulaWorkbook ?? undefined)
    if (extracted) {
      return {
        id: nextChartId(),
        title: `${rangeToken} ${type} chart`,
        type,
        range: rangeToken,
        labels: extracted.labels,
        values: extracted.values,
        series: extracted.series,
      }
    }
  }
  const rows: Array<{ label: string; value: number }> = []
  const hasLabelColumn = bounds.end.column > bounds.start.column
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    const labelKey = cellKey(row, bounds.start.column)
    const valueKey = hasLabelColumn
      ? cellKey(row, bounds.start.column + 1)
      : cellKey(row, bounds.start.column)
    const rawText = computedCellValue(sheet, valueKey, workbook)
    // A blank cell is not a zero data point (Number('') === 0).
    if (rawText.trim() === '') continue
    const rawValue = Number(rawText)
    if (!Number.isFinite(rawValue)) continue
    rows.push({
      label: hasLabelColumn
        ? displayCellValue(sheet, labelKey, workbook) || labelKey
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

const FORMULA_ERROR_VALUES = new Set([
  '#DIV/0!',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#VALUE',
  '#VALUE!',
  '#FORMULA',
])

export function isFormulaError(value: string): boolean {
  return FORMULA_ERROR_VALUES.has(value.trim().toUpperCase())
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

// Per-state memo of formula results. A cell's result is a pure function of the
// workbook it is read in, the sheet object it lives on and its key — documents
// are immutable, so every edit makes new objects and the old entries simply go
// unused (and are collected with them). Without it every reference re-evaluated
// its whole dependency chain: a cell naming B120 four times evaluated B120's
// chain four times, and a running-total column made that exponential — one cell
// of a 742-formula model took 16 s, and saving (which reads every cell) pinned
// the renderer for minutes while an agent was writing to it.
//
// A result is only kept when it did not depend on HOW it was reached: the cycle
// guard's #REF!, the depth backstop's #REF!, and NOW()/TODAY() all bump
// `uncacheableResults`, and any evaluation that saw a bump is left uncached, so
// cyclic and volatile cells evaluate exactly as they did before. (One edge: a
// chain deeper than the 2048 backstop may now resolve where it used to stop at
// #REF!, because a kept link shortens it.)
const NO_WORKBOOK: object = {}
const formulaResults = new WeakMap<object, WeakMap<WorkbookSheet, Map<string, string>>>()
let uncacheableResults = 0

function formulaResultsFor(sheet: WorkbookSheet): Map<string, string> {
  const root = formulaWorkbook ?? NO_WORKBOOK
  let bySheet = formulaResults.get(root)
  if (!bySheet) {
    bySheet = new WeakMap()
    formulaResults.set(root, bySheet)
  }
  let results = bySheet.get(sheet)
  if (!results) {
    results = new Map()
    bySheet.set(sheet, results)
  }
  return results
}

/** The sheets a formula may name (Sheet2!A1). */
export interface FormulaWorkbook {
  sheets: readonly WorkbookSheet[]
  namedRanges?: readonly import('../types').NamedRange[]
}
/** The workbook of the evaluation under way: set by the public entry points, kept by everything they call. */
let formulaWorkbook: FormulaWorkbook | null = null

function inWorkbook<T>(workbook: FormulaWorkbook | undefined, run: () => T): T {
  if (!workbook) return run()
  const before = formulaWorkbook
  formulaWorkbook = workbook
  try {
    return run()
  } finally {
    formulaWorkbook = before
  }
}

/** A sheet name as a formula writes it: Costs, or 'Q1 plan' with '' for a quote. */
const SHEET_PREFIX = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!`

/**
 * The sheet a reference reads and the reference without its sheet: "Costs!B4" is B4 on Costs. A reference with no
 * sheet reads its own. A sheet that does not exist (or no workbook to look in) is null: #REF!, as in Excel.
 */
function targetOf(sheet: WorkbookSheet, token: string): [WorkbookSheet | null, string] {
  const trimmed = token.trim()
  const bang = trimmed.lastIndexOf('!')
  if (bang <= 0) return [sheet, trimmed]
  const named = trimmed.slice(0, bang)
  const name = named.startsWith("'") && named.endsWith("'") ? named.slice(1, -1).replace(/''/g, "'") : named
  const rest = trimmed.slice(bang + 1)
  if (name.toLowerCase() === sheet.name.toLowerCase()) return [sheet, rest]
  const found = formulaWorkbook?.sheets.find(each => each.name.toLowerCase() === name.toLowerCase()) ?? null
  return [found, rest]
}

/** A plain or sheet-qualified cell or range (A1, $A$1:B9, Costs!B4, 'Q1 plan'!A1:A9). */
const REFERENCE = new RegExp(String.raw`^(?:${SHEET_PREFIX})?\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$`, 'i')

export function evaluateFormula(sheet: WorkbookSheet, formula: string, workbook?: FormulaWorkbook): string {
  return inWorkbook(workbook, () => evaluateFormulaHere(sheet, formula))
}

function evaluateFormulaHere(sheet: WorkbookSheet, formula: string): string {
  if (formulaEvaluationDepth >= MAX_FORMULA_EVALUATION_DEPTH) {
    uncacheableResults += 1
    return '#REF!'
  }
  formulaEvaluationDepth += 1
  try {
    const expanded = formulaWorkbook?.namedRanges?.length
      ? resolveNamedRangesInFormula(
          formula,
          formulaWorkbook.namedRanges,
          formulaWorkbook.sheets,
          sheet.id,
        )
      : formula
    const expression = expanded.trim().replace(/^=/, '')
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
  if (REFERENCE.test(arg.trim())) {
    return expandRange(sheet, arg)
  }
  const evaluated = evaluateExpressionString(sheet, arg)
  return [evaluated ?? unquote(arg)]
}

function formulaArgumentError(
  sheet: WorkbookSheet,
  arg: string,
): string | null {
  const trimmed = arg.trim()
  if (REFERENCE.test(trimmed)) {
    const [target, range] = targetOf(sheet, trimmed)
    if (!target) return '#REF!'
    const keys = cellKeysInRange(range.replace(/\$/g, ''))
    if (keys === null) return '#NUM!'
    for (const key of keys) {
      if (target.cells[key]?.kind !== 'formula') continue
      const value = rawCellValue(target, key)
      if (isFormulaError(value)) return value
    }
    return null
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) return null
  const evaluated = evaluateExpressionString(sheet, trimmed)
  return evaluated !== null && isFormulaError(evaluated) ? evaluated : null
}

function applyFormulaFunction(
  sheet: WorkbookSheet,
  fn: string,
  rawArgs: string[],
): string {
  if (['SUM', 'AVERAGE', 'MIN', 'MAX'].includes(fn)) {
    for (const part of rawArgs) {
      const error = formulaArgumentError(sheet, part)
      if (error) return error
    }
  }
  if (fn === 'IF') {
    const [condition, truthy = '', falsy = ''] = rawArgs
    const conditionResult = truthyCondition(sheet, condition)
    if (typeof conditionResult === 'string') return conditionResult
    return conditionResult
      ? resolveTextArg(sheet, truthy)
      : resolveTextArg(sheet, falsy)
  }
  if (fn === 'ROUND') {
    const valueRaw = rawArgs[0] !== undefined ? resolveTextArg(sheet, rawArgs[0]) : ''
    const placesRaw = rawArgs[1] !== undefined ? resolveTextArg(sheet, rawArgs[1]) : '0'
    if (isFormulaError(valueRaw)) return valueRaw
    if (isFormulaError(placesRaw)) return placesRaw
    const value = Number(valueRaw)
    const places = Math.trunc(Number(placesRaw))
    if (
      valueRaw.trim() === '' ||
      !Number.isFinite(value) ||
      !Number.isFinite(places)
    ) {
      return '#VALUE!'
    }
    if (value === 0) return '0'
    const clampedPlaces = Math.max(-100, Math.min(100, places))
    const shiftDecimal = (abs: number, exp: number): number => {
      if (abs === 0 || exp === 0) return abs
      const [mantissa, rawExp] = String(abs).split(/[eE]/)
      const nextExp = (rawExp ? Number(rawExp) : 0) + exp
      return Number(`${mantissa}e${nextExp}`)
    }
    const sign = value < 0 ? -1 : 1
    const abs = Math.abs(value)
    const shifted = shiftDecimal(abs, clampedPlaces)
    if (!Number.isFinite(shifted)) {
      return clampedPlaces > 0 ? String(value) : '0'
    }
    const roundedMag = shiftDecimal(Math.round(shifted), -clampedPlaces)
    if (!Number.isFinite(roundedMag)) return '#NUM!'
    const rounded = sign * roundedMag
    return String(Object.is(rounded, -0) || rounded === 0 ? 0 : rounded)
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
  if (fn === 'TODAY' || fn === 'NOW') {
    uncacheableResults += 1 // volatile: never memoised
    return fn === 'TODAY' ? new Date().toISOString().slice(0, 10) : new Date().toISOString()
  }
  if (fn === 'COUNTIF') return countIf(sheet, rawArgs)
  if (fn === 'SUMIF') return sumIf(sheet, rawArgs)
  if (fn === 'AVERAGEIF') return averageIf(sheet, rawArgs)
  if (fn === 'RANK') return rank(sheet, rawArgs)
  if (fn === 'IFERROR') {
    const [tryArg = '', fallback = ''] = rawArgs
    const result = resolveTextArg(sheet, tryArg)
    return isFormulaError(result) ? resolveTextArg(sheet, fallback) : result
  }

  const rawValues = rawArgs.flatMap(part =>
    numericArgValues(sheet, part.trim()),
  )
  const values = rawValues
    // A blank cell is not a zero: without this filter, Number('') === 0 made
    // COUNT/AVERAGE/MIN/MAX over sparse ranges count phantom zeros (caught by
    // the engine-agnostic golden suite against Univer's engine).
    .filter(value => value.trim() !== '')
    .map(value => Number(value))
    .filter(value => Number.isFinite(value))

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
  if (fn === 'COUNT') {
    if (rawValues.some(v => v === '#NUM!')) return '#NUM!'
    return String(values.length)
  }
  if (fn === 'COUNTA') {
    const textValues = rawArgs.flatMap(part => expandRange(sheet, part.trim()))
    if (textValues.some(v => v === '#NUM!')) return '#NUM!'
    return String(textValues.filter(value => value !== '').length)
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
    const rangeMatch = new RegExp(String.raw`^(?:${SHEET_PREFIX})?\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+`, 'i').exec(
      input.slice(index),
    )
    if (rangeMatch) {
      tokens.push({ type: 'range', text: rangeMatch[0].replace(/\$/g, '') })
      index += rangeMatch[0].length
      continue
    }
    const refMatch = new RegExp(String.raw`^(?:${SHEET_PREFIX})?\$?[A-Z]+\$?\d+(?![A-Z0-9(])`, 'i').exec(input.slice(index))
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

interface EvalValue {
  text: string
  isBlankRef?: boolean
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

  const coerceNumeric = (val: EvalValue): number | null => {
    if (val.isBlankRef && val.text === '') return 0
    if (val.text.trim() === '') return null
    const numeric = Number(val.text)
    return Number.isFinite(numeric) ? numeric : null
  }

  function parsePrimary(): EvalValue | null {
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
      if (isFormulaError(operand.text)) return { text: operand.text }
      const numeric = coerceNumeric(operand)
      if (numeric === null) return { text: '#VALUE!' }
      const negated = -numeric
      return { text: String(Object.is(negated, -0) ? 0 : negated) }
    }
    position += 1
    if (token.type === 'number' || token.type === 'string') {
      return { text: token.text }
    }
    if (token.type === 'ref') {
      const [target, ref] = targetOf(sheet, token.text)
      const positionKey = target ? parseCellKey(ref) : null
      if (!target || !positionKey) return { text: '#REF!' }
      const key = cellKey(positionKey.row, positionKey.column)
      const cell = target.cells[key]
      const text = rawCellValue(target, key)
      return { text, isBlankRef: !cell || cell.kind === 'blank' }
    }
    if (token.type === 'range') {
      // A bare range outside a function argument resolves to its first cell
      const [target, rangePart] = targetOf(sheet, token.text)
      const firstKey = rangePart.split(':')[0] ?? ''
      const firstPos = target ? parseCellKey(firstKey) : null
      const firstCell =
        target && firstPos
          ? target.cells[cellKey(firstPos.row, firstPos.column)]
          : undefined
      return {
        text: expandRange(sheet, token.text)[0] ?? '',
        isBlankRef: !firstCell || firstCell.kind === 'blank',
      }
    }
    if (token.type === 'func') {
      return {
        text: applyFormulaFunction(
          sheet,
          token.text,
          splitFormulaArguments(token.args ?? ''),
        ),
      }
    }
    return null
  }

  function applyNumericOp(
    left: EvalValue,
    op: string,
    right: EvalValue,
  ): EvalValue {
    if (isFormulaError(left.text)) return { text: left.text }
    if (isFormulaError(right.text)) return { text: right.text }
    const a = coerceNumeric(left)
    const b = coerceNumeric(right)
    if (a === null || b === null) return { text: '#VALUE!' }
    if (op === '/') {
      if (b === 0) return { text: '#DIV/0!' }
      const out = a / b
      return { text: Number.isFinite(out) ? String(out) : '#NUM!' }
    }
    const out = op === '*' ? a * b : op === '-' ? a - b : a + b
    return {
      text: Number.isFinite(out)
        ? String(Object.is(out, -0) ? 0 : out)
        : '#NUM!',
    }
  }

  function parseMultiplicative(): EvalValue | null {
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

  function parseAdditive(): EvalValue | null {
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

  function parseConcat(): EvalValue | null {
    let left = parseAdditive()
    if (left === null) return null
    while (isOp('&')) {
      position += 1
      const right = parseAdditive()
      if (right === null) return null
      if (isFormulaError(left.text)) return { text: left.text }
      if (isFormulaError(right.text)) return { text: right.text }
      left = { text: `${left.text}${right.text}` }
    }
    return left
  }

  const result = parseConcat()
  if (result === null || position < tokens.length) return null
  return result.text
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
  if (cellValues.length === 1 && cellValues[0] === '#NUM!') return '#NUM!'
  const criteria = resolveCriteriaArg(sheet, criteriaArg)
  const count = cellValues.filter(v => matchesCriteria(v, criteria)).length
  return String(count)
}

function sumIf(sheet: WorkbookSheet, args: string[]): string {
  const [rangeArg = '', criteriaArg = '', sumRangeArg] = args
  const [rangeSheet, range] = targetOf(sheet, rangeArg)
  const [sumSheet, sumRange] = sumRangeArg ? targetOf(sheet, sumRangeArg) : [rangeSheet, range]
  if (!rangeSheet || !sumSheet) return '#REF!'
  const rangeKeys = cellKeysInRange(range)
  const sumKeys = sumRangeArg ? cellKeysInRange(sumRange) : rangeKeys
  if (rangeKeys === null || sumKeys === null) return '#NUM!'
  const criteria = resolveCriteriaArg(sheet, criteriaArg)
  let total = 0
  rangeKeys.forEach((key, idx) => {
    if (matchesCriteria(rawCellValue(rangeSheet, key), criteria)) {
      total += Number(rawCellValue(sumSheet, sumKeys[idx] ?? key)) || 0
    }
  })
  return String(total)
}

function averageIf(sheet: WorkbookSheet, args: string[]): string {
  const [rangeArg = '', criteriaArg = '', avgRangeArg] = args
  const [rangeSheet, range] = targetOf(sheet, rangeArg)
  const [avgSheet, avgRange] = avgRangeArg ? targetOf(sheet, avgRangeArg) : [rangeSheet, range]
  if (!rangeSheet || !avgSheet) return '#REF!'
  const rangeKeys = cellKeysInRange(range)
  const avgKeys = avgRangeArg ? cellKeysInRange(avgRange) : rangeKeys
  if (rangeKeys === null || avgKeys === null) return '#NUM!'
  const criteria = resolveCriteriaArg(sheet, criteriaArg)
  const matchedValues: number[] = []
  rangeKeys.forEach((key, idx) => {
    if (matchesCriteria(rawCellValue(rangeSheet, key), criteria)) {
      const v = Number(rawCellValue(avgSheet, avgKeys[idx] ?? key))
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
  const targetRaw = resolveTextArg(sheet, valueArg)
  const target = Number(targetRaw)
  if (targetRaw.trim() === '' || !Number.isFinite(target)) return '#VALUE!'
  const expanded = expandRange(sheet, refArg.replace(/\$/g, '').trim())
  if (expanded.length === 1 && expanded[0] === '#NUM!') return '#NUM!'
  const refValues = expanded
    .filter(value => value.trim() !== '')
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

// Returns cell key strings (A1, B2...) for every cell in a range token,
// or null when the range exceeds MAX_RANGE_EXPANSION_CELLS.
function cellKeysInRange(
  rangeToken: string,
  maxCells = MAX_RANGE_EXPANSION_CELLS,
): string[] | null {
  const normalised = rangeToken.replace(/\$/g, '')
  const range = /^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$/i.exec(normalised)
  if (!range) {
    const pos = parseCellKey(normalised)
    return pos ? [cellKey(pos.row, pos.column)] : []
  }
  const start = parseCellKey(range[1])
  const end = parseCellKey(range[2])
  if (!start || !end) return []
  const minRow = Math.min(start.row, end.row)
  const maxRow = Math.max(start.row, end.row)
  const minCol = Math.min(start.column, end.column)
  const maxCol = Math.max(start.column, end.column)
  const count = (maxRow - minRow + 1) * (maxCol - minCol + 1)
  if (!Number.isSafeInteger(count) || count > maxCells) return null
  const keys: string[] = []
  for (let row = minRow; row <= maxRow; row++) {
    for (let col = minCol; col <= maxCol; col++) {
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
  if (REFERENCE.test(token.trim()) && token.includes('!')) {
    const [target, range] = targetOf(sheet, token)
    return target ? expandRange(target, range) : ['#REF!']
  }
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
  const minRow = Math.min(start.row, end.row)
  const maxRow = Math.max(start.row, end.row)
  const minCol = Math.min(start.column, end.column)
  const maxCol = Math.max(start.column, end.column)
  const count = (maxRow - minRow + 1) * (maxCol - minCol + 1)
  if (!Number.isSafeInteger(count) || count > MAX_RANGE_EXPANSION_CELLS) {
    return ['#NUM!']
  }
  const values: string[] = []
  for (let row = minRow; row <= maxRow; row += 1) {
    for (let column = minCol; column <= maxCol; column += 1) {
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

function truthyCondition(
  sheet: WorkbookSheet,
  condition: string,
): boolean | string {
  const comparison = /^(.+?)(>=|<=|<>|=|>|<)(.+)$/.exec(condition.trim())
  if (!comparison) {
    const error = formulaArgumentError(sheet, condition)
    if (error) return error
    const value = resolveTextArg(sheet, condition)
    return Boolean(value) && value !== '0' && value.toLowerCase() !== 'false'
  }
  const leftError = formulaArgumentError(sheet, comparison[1])
  if (leftError) return leftError
  const rightError = formulaArgumentError(sheet, comparison[3])
  if (rightError) return rightError
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
  const [table, tableRange] = targetOf(sheet, rangeArg ?? '')
  if (!table) return '#REF!'
  const bounds = parseRange(tableRange)
  if (!bounds) return '#VALUE!'
  const rawColText = resolveTextArg(sheet, columnArg ?? '1')
  if (isFormulaError(rawColText)) return rawColText
  const rawColNum = Number(rawColText)
  if (rawColText.trim() === '' || !Number.isFinite(rawColNum)) return '#VALUE!'
  const colIndexNum = Math.floor(rawColNum)
  if (colIndexNum < 1) return '#VALUE!'
  const tableWidth = bounds.end.column - bounds.start.column + 1
  if (colIndexNum > tableWidth) return '#REF!'
  const targetColumnOffset = colIndexNum - 1
  const lookupValue = resolveTextArg(sheet, lookupArg ?? '')
  if (isFormulaError(lookupValue)) return lookupValue
  const lookupLower = lookupValue.toLowerCase()
  const lookupNumber = Number(lookupValue)
  const lookupIsNumber = lookupValue.trim() !== '' && Number.isFinite(lookupNumber)
  for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
    const candidate = rawCellValue(table, cellKey(row, bounds.start.column))
    const candidateNumber = Number(candidate)
    const matched =
      lookupIsNumber &&
      candidate.trim() !== '' &&
      Number.isFinite(candidateNumber)
        ? candidateNumber === lookupNumber
        : candidate.toLowerCase() === lookupLower
    if (matched) {
      return rawCellValue(
        table,
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
  if (value.trim() === '') return value
  if (isFormulaError(value)) return value
  const numeric = Number(value)
  const decimals =
    style.decimals !== undefined && Number.isFinite(style.decimals)
      ? Math.min(8, Math.max(0, Math.floor(style.decimals)))
      : undefined
  if (style.numberFormat === 'number') {
    if (!Number.isFinite(numeric)) return value
    return decimals !== undefined ? numeric.toFixed(decimals) : String(numeric)
  }
  if (style.numberFormat === 'currency' || style.numberFormat === 'accounting') {
    const currencyCode =
      style.currencySymbol === '€'
        ? 'EUR'
        : style.currencySymbol === '£'
        ? 'GBP'
        : style.currencySymbol === '¥'
        ? 'JPY'
        : 'USD'
    return Number.isFinite(numeric)
      ? new Intl.NumberFormat('en-US', {
          style: 'currency',
          currency: currencyCode,
          currencySign: style.numberFormat === 'accounting' ? 'accounting' : 'standard',
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
  if (style.numberFormat === 'scientific') {
    return Number.isFinite(numeric)
      ? numeric.toExponential(decimals ?? 2).toUpperCase()
      : value
  }
  if (style.numberFormat === 'time') {
    const date = new Date(value)
    return Number.isNaN(date.getTime())
      ? value
      : date.toISOString().slice(11, 19)
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
  // Keyed by sheet as well: Costs!A1 and Plan!A1 are different cells.
  const where = `${sheet.id}!${key}`
  if (evaluatingCells.has(where)) {
    uncacheableResults += 1
    return '#REF!'
  }
  // Each cell computes once per workbook state (see formulaResults).
  const results = formulaResultsFor(sheet)
  const known = results.get(key)
  if (known !== undefined) return known
  const uncacheableBefore = uncacheableResults
  evaluatingCells.add(where)
  try {
    const value = evaluateFormulaHere(sheet, cell.value)
    if (uncacheableResults === uncacheableBefore) results.set(key, value)
    return value
  } finally {
    evaluatingCells.delete(where)
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
  let maxRow = sheet.rowCount - 1
  let maxColumn = sheet.columnCount - 1
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
    const pos = parseCellKey(change.cell)
    if (pos) {
      if (pos.row > maxRow) maxRow = pos.row
      if (pos.column > maxColumn) maxColumn = pos.column
    }
    logged.push({
      cell: change.cell,
      ...(before !== undefined ? { before } : {}),
      after: change.value,
      kind: change.kind,
    })
  }
  if (!logged.length) return { document, applied: 0, skippedLocked }
  if (maxRow + 1 > nextSheet.rowCount || maxColumn + 1 > nextSheet.columnCount) {
    nextSheet = {
      ...nextSheet,
      rowCount: Math.max(nextSheet.rowCount, maxRow + 1),
      columnCount: Math.max(nextSheet.columnCount, maxColumn + 1),
    }
  }
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
  const chart = chartFromRange(sheet, range, type, document.workbook)
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
