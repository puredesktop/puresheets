import type {
  CellStyle,
  ColumnFilterRule,
  PureSheetsDocument,
  SheetCell,
  WorkbookSheet,
} from '../types'
import {
  cellKey,
  cellsInRange,
  computedCellValue,
  deriveCellKind,
  displayCellValue,
  isRangeLocked,
  normalizeRangeToken,
  parseCellKey,
  type CopiedCellBlock,
  type FormulaWorkbook,
} from './workbookModel'

function parseBounds(rangeToken: string): {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
} | null {
  const norm = normalizeRangeToken(rangeToken)
  if (!norm) return null
  const [a, b] = norm.split(':')
  const start = parseCellKey(a)
  const end = parseCellKey(b ?? a)
  if (!start || !end) return null
  return {
    startRow: Math.min(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endRow: Math.max(start.row, end.row),
    endColumn: Math.max(start.column, end.column),
  }
}

// ---------------------------------------------------------------------------
// Merge / Unmerge
// ---------------------------------------------------------------------------

export function isRangeMerged(sheet: WorkbookSheet, rangeToken: string): boolean {
  const bounds = parseBounds(rangeToken)
  if (!bounds || !sheet.merges?.length) return false
  return sheet.merges.some(merge => {
    const m = parseBounds(merge)
    if (!m) return false
    return !(
      m.endRow < bounds.startRow ||
      m.startRow > bounds.endRow ||
      m.endColumn < bounds.startColumn ||
      m.startColumn > bounds.endColumn
    )
  })
}

export function mergeRangeInSheet(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const bounds = parseBounds(rangeToken)
  if (!bounds) return sheet
  if (bounds.startRow === bounds.endRow && bounds.startColumn === bounds.endColumn) {
    return sheet
  }
  if (isRangeLocked(sheet, rangeToken)) return sheet
  const canonical = `${cellKey(bounds.startRow, bounds.startColumn)}:${cellKey(
    bounds.endRow,
    bounds.endColumn,
  )}`
  // Remove any overlapping prior merges
  const remaining = (sheet.merges ?? []).filter(merge => {
    const m = parseBounds(merge)
    if (!m) return false
    return (
      m.endRow < bounds.startRow ||
      m.startRow > bounds.endRow ||
      m.endColumn < bounds.startColumn ||
      m.startColumn > bounds.endColumn
    )
  })
  // Keep top-left value; clear values in secondary cells of the merge
  const nextCells = { ...sheet.cells }
  const topLeftKey = cellKey(bounds.startRow, bounds.startColumn)
  for (let r = bounds.startRow; r <= bounds.endRow; r += 1) {
    for (let c = bounds.startColumn; c <= bounds.endColumn; c += 1) {
      const k = cellKey(r, c)
      if (k === topLeftKey) continue
      const existing = nextCells[k]
      if (existing?.style) {
        nextCells[k] = { value: '', kind: 'blank', style: existing.style }
      } else {
        delete nextCells[k]
      }
    }
  }
  return {
    ...sheet,
    cells: nextCells,
    merges: [...remaining, canonical],
  }
}

export function unmergeRangeInSheet(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const bounds = parseBounds(rangeToken)
  if (!bounds || !sheet.merges?.length) return sheet
  const remaining = sheet.merges.filter(merge => {
    const m = parseBounds(merge)
    if (!m) return false
    return (
      m.endRow < bounds.startRow ||
      m.startRow > bounds.endRow ||
      m.endColumn < bounds.startColumn ||
      m.startColumn > bounds.endColumn
    )
  })
  if (remaining.length === sheet.merges.length) return sheet
  return {
    ...sheet,
    merges: remaining.length ? remaining : undefined,
  }
}

// ---------------------------------------------------------------------------
// Clear formatting / Clear all / Freeze panes / Gridlines / Hide-Unhide / Sheet ops
// ---------------------------------------------------------------------------

export function clearFormattingInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  const keys = cellsInRange(rangeToken, 50_000)
  if (!keys.length) return sheet
  const nextCells = { ...sheet.cells }
  let changed = false
  for (const key of keys) {
    const cell = nextCells[key]
    if (!cell?.style) continue
    changed = true
    if (!cell.value && !cell.hyperlink && !cell.checkbox) {
      delete nextCells[key]
    } else {
      const { style: _drop, ...rest } = cell
      nextCells[key] = rest
    }
  }
  return changed ? { ...sheet, cells: nextCells } : sheet
}

export function setFreezePanes(
  sheet: WorkbookSheet,
  frozenRows: number,
  frozenColumns: number,
): WorkbookSheet {
  const r = Math.max(0, Math.floor(frozenRows))
  const c = Math.max(0, Math.floor(frozenColumns))
  return {
    ...sheet,
    frozenRows: r || undefined,
    frozenColumns: c || undefined,
  }
}

export function hideRowsInSheet(
  sheet: WorkbookSheet,
  rows: number[],
): WorkbookSheet {
  const valid = rows.filter(r => Number.isInteger(r) && r >= 0 && r < sheet.rowCount)
  if (!valid.length) return sheet
  const merged = [...new Set([...(sheet.hiddenRows ?? []), ...valid])].sort((a, b) => a - b)
  return { ...sheet, hiddenRows: merged }
}

export function unhideAllRowsInSheet(sheet: WorkbookSheet): WorkbookSheet {
  if (!sheet.hiddenRows?.length) return sheet
  return { ...sheet, hiddenRows: undefined }
}

export function hideColumnsInSheet(
  sheet: WorkbookSheet,
  columns: number[],
): WorkbookSheet {
  const valid = columns.filter(
    c => Number.isInteger(c) && c >= 0 && c < sheet.columnCount,
  )
  if (!valid.length) return sheet
  const merged = [...new Set([...(sheet.hiddenColumns ?? []), ...valid])].sort(
    (a, b) => a - b,
  )
  return { ...sheet, hiddenColumns: merged }
}

export function unhideAllColumnsInSheet(sheet: WorkbookSheet): WorkbookSheet {
  if (!sheet.hiddenColumns?.length) return sheet
  return { ...sheet, hiddenColumns: undefined }
}

export function toggleSheetGridlines(sheet: WorkbookSheet): WorkbookSheet {
  const currentlyVisible = sheet.showGridlines !== false
  return {
    ...sheet,
    showGridlines: currentlyVisible ? false : undefined,
  }
}

export function setSheetTabColor(
  sheet: WorkbookSheet,
  color: string | undefined,
): WorkbookSheet {
  const trimmed = color?.trim()
  return {
    ...sheet,
    tabColor: trimmed || undefined,
  }
}

export function setSheetHidden(
  document: PureSheetsDocument,
  sheetId: string,
  hidden: boolean,
): PureSheetsDocument {
  const visible = document.workbook.sheets.filter(
    s => !s.agentDraft && !s.hidden,
  )
  if (hidden && visible.length <= 1 && visible[0]?.id === sheetId) {
    return document // keep at least one visible sheet
  }
  const sheets = document.workbook.sheets.map(s =>
    s.id === sheetId ? { ...s, hidden: hidden ? true : undefined } : s,
  )
  let activeSheetId = document.workbook.activeSheetId
  if (hidden && activeSheetId === sheetId) {
    const fallback = sheets.find(s => !s.agentDraft && !s.hidden)
    if (fallback) activeSheetId = fallback.id
  }
  return {
    ...document,
    workbook: { ...document.workbook, activeSheetId, sheets },
  }
}

export function moveSheetInWorkbook(
  document: PureSheetsDocument,
  sheetId: string,
  targetIndex: number,
): PureSheetsDocument {
  const sheets = [...document.workbook.sheets]
  const fromIndex = sheets.findIndex(s => s.id === sheetId)
  if (fromIndex < 0) return document
  const clamped = Math.max(0, Math.min(sheets.length - 1, targetIndex))
  if (clamped === fromIndex) return document
  const [moved] = sheets.splice(fromIndex, 1)
  sheets.splice(clamped, 0, moved)
  return {
    ...document,
    workbook: { ...document.workbook, sheets },
  }
}

// ---------------------------------------------------------------------------
// Hyperlinks & Checkboxes (Smart Cells)
// ---------------------------------------------------------------------------

export function setCellHyperlink(
  sheet: WorkbookSheet,
  key: string,
  url: string,
  label?: string,
): WorkbookSheet {
  const trimmedUrl = url.trim()
  const existing = sheet.cells[key]
  if (!trimmedUrl) {
    if (!existing?.hyperlink) return sheet
    const { hyperlink: _drop, ...rest } = existing
    return { ...sheet, cells: { ...sheet.cells, [key]: rest } }
  }
  const value = label?.trim() || existing?.value || trimmedUrl
  const style: CellStyle = {
    ...existing?.style,
    underline: true,
    textColor: existing?.style?.textColor ?? '#1d6fa5',
  }
  return {
    ...sheet,
    cells: {
      ...sheet.cells,
      [key]: {
        value,
        kind: deriveCellKind(value),
        style,
        hyperlink: trimmedUrl,
      },
    },
  }
}

export function toggleCheckboxInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  if (isRangeLocked(sheet, rangeToken)) return sheet
  const keys = cellsInRange(rangeToken, 5_000)
  if (!keys.length) return sheet
  const allChecked = keys.every(k => sheet.cells[k]?.checkbox)
  const nextCells = { ...sheet.cells }
  for (const key of keys) {
    const existing = nextCells[key]
    if (allChecked) {
      if (existing) {
        const { checkbox: _drop, ...rest } = existing
        nextCells[key] = rest
      }
    } else {
      const normalized =
        existing?.value?.trim().toUpperCase() === 'TRUE' ? 'TRUE' : 'FALSE'
      nextCells[key] = {
        ...existing,
        value: normalized,
        kind: 'text',
        checkbox: true,
        style: { ...existing?.style, align: existing?.style?.align ?? 'center' },
      }
    }
  }
  return { ...sheet, cells: nextCells }
}

// ---------------------------------------------------------------------------
// Header-Aware Multi-Level Sort (Wave 3)
// ---------------------------------------------------------------------------

export interface SortKey {
  column: number
  direction: 'asc' | 'desc'
}

export function sortRangeMulti(
  sheet: WorkbookSheet,
  rangeToken: string,
  keys: SortKey[],
  hasHeader = true,
  workbook?: FormulaWorkbook,
): WorkbookSheet {
  const bounds = parseBounds(rangeToken)
  if (!bounds || !keys.length) return sheet
  if (isRangeLocked(sheet, rangeToken)) return sheet
  const dataStartRow = hasHeader ? bounds.startRow + 1 : bounds.startRow
  if (dataStartRow > bounds.endRow) return sheet

  const rows: Array<{
    row: number
    cells: Array<SheetCell | undefined>
    allBlank: boolean
  }> = []

  for (let row = dataStartRow; row <= bounds.endRow; row += 1) {
    const cells: Array<SheetCell | undefined> = []
    let anyNonBlank = false
    for (let col = bounds.startColumn; col <= bounds.endColumn; col += 1) {
      const k = cellKey(row, col)
      cells.push(sheet.cells[k])
      if (displayCellValue(sheet, k, workbook).trim() !== '') anyNonBlank = true
    }
    rows.push({ row, cells, allBlank: !anyNonBlank })
  }

  const nonBlank = rows.filter(r => !r.allBlank)
  const blank = rows.filter(r => r.allBlank)

  nonBlank.sort((a, b) => {
    for (const sortKey of keys) {
      const leftVal = computedCellValue(
        sheet,
        cellKey(a.row, sortKey.column),
        workbook,
      ).trim()
      const rightVal = computedCellValue(
        sheet,
        cellKey(b.row, sortKey.column),
        workbook,
      ).trim()
      if (leftVal === '' && rightVal === '') continue
      if (leftVal === '') return 1
      if (rightVal === '') return -1
      const leftNum = Number(leftVal)
      const rightNum = Number(rightVal)
      const cmp =
        Number.isFinite(leftNum) && Number.isFinite(rightNum)
          ? leftNum - rightNum
          : leftVal.localeCompare(rightVal, undefined, { numeric: true, sensitivity: 'base' })
      if (cmp !== 0) return sortKey.direction === 'asc' ? cmp : -cmp
    }
    return a.row - b.row
  })

  const ordered = [...nonBlank, ...blank]
  const nextCells = { ...sheet.cells }
  ordered.forEach((entry, idx) => {
    const targetRow = dataStartRow + idx
    entry.cells.forEach((cell, colOffset) => {
      const targetKey = cellKey(targetRow, bounds.startColumn + colOffset)
      if (cell) nextCells[targetKey] = cell
      else delete nextCells[targetKey]
    })
  })

  return { ...sheet, cells: nextCells }
}

// ---------------------------------------------------------------------------
// Multi-Column Smart Filter & Instant Column Profile (Wave 3)
// ---------------------------------------------------------------------------

export function setColumnFilter(
  sheet: WorkbookSheet,
  column: number,
  rule: ColumnFilterRule | null,
): WorkbookSheet {
  const next = { ...(sheet.columnFilters ?? {}) }
  const colKey = String(Math.max(0, Math.floor(column)))
  if (
    !rule ||
    (!rule.selectedValues && !rule.operator)
  ) {
    delete next[colKey]
  } else {
    next[colKey] = rule
  }
  const hasFilters = Object.keys(next).length > 0
  return {
    ...sheet,
    columnFilters: hasFilters ? next : undefined,
    filterRow: hasFilters ? 0 : sheet.filterRow,
    filterHasHeader: sheet.filterHasHeader ?? true,
  }
}

export function clearAllColumnFilters(sheet: WorkbookSheet): WorkbookSheet {
  return {
    ...sheet,
    columnFilters: undefined,
    filterRow: undefined,
    filterColumn: undefined,
    filterQuery: undefined,
  }
}

export function activeFilterColumnIndices(sheet: WorkbookSheet): number[] {
  const cols = new Set<number>()
  if (sheet.filterRow === 0 && sheet.filterColumn !== undefined) {
    cols.add(sheet.filterColumn)
  }
  for (const key of Object.keys(sheet.columnFilters ?? {})) {
    const idx = Number(key)
    if (Number.isInteger(idx) && idx >= 0) cols.add(idx)
  }
  return [...cols].sort((a, b) => a - b)
}

function cellMatchesColumnRule(
  rawDisplay: string,
  rawComputed: string,
  rule: ColumnFilterRule,
): boolean {
  if (rule.selectedValues) {
    const label = rawDisplay.trim() === '' ? '(Blank)' : rawDisplay
    if (!rule.selectedValues.includes(label)) return false
  }
  if (!rule.operator) return true
  const trimmed = rawDisplay.trim()
  const lower = trimmed.toLowerCase()
  const val = (rule.value ?? '').trim()
  const valLower = val.toLowerCase()
  if (rule.operator === 'blank') return trimmed === ''
  if (rule.operator === 'not-blank') return trimmed !== ''
  if (rule.operator === 'contains') return lower.includes(valLower)
  if (rule.operator === 'not-contains') return !lower.includes(valLower)
  if (rule.operator === 'equals') {
    const n1 = Number(rawComputed)
    const n2 = Number(val)
    if (rawComputed.trim() !== '' && val !== '' && Number.isFinite(n1) && Number.isFinite(n2)) {
      return n1 === n2
    }
    return lower === valLower
  }
  const num = Number(rawComputed)
  const target = Number(val)
  if (rawComputed.trim() === '' || !Number.isFinite(num) || !Number.isFinite(target)) {
    return false
  }
  if (rule.operator === 'gt') return num > target
  if (rule.operator === 'lt') return num < target
  if (rule.operator === 'between') {
    const upper = Number(rule.value2 ?? '')
    if (!Number.isFinite(upper)) return false
    return num >= Math.min(target, upper) && num <= Math.max(target, upper)
  }
  return true
}

/**
 * Compute all visible row indices honoring:
 * 1) explicitly hidden rows (`sheet.hiddenRows`)
 * 2) legacy single-column filter (`filterRow === 0` + `filterQuery`)
 * 3) multi-column smart filters (`sheet.columnFilters`)
 */
export function visibleRowsWithFilters(
  sheet: WorkbookSheet,
  workbook?: FormulaWorkbook,
): number[] {
  const hiddenSet = new Set(sheet.hiddenRows ?? [])
  const hasMulti =
    sheet.columnFilters && Object.keys(sheet.columnFilters).length > 0
  const hasLegacy = sheet.filterRow === 0 && !hasMulti
  const hasHeader = sheet.filterHasHeader !== false
  const filterEntries = Object.entries(sheet.columnFilters ?? {})
  const legacyCol = Math.max(0, Math.floor(sheet.filterColumn ?? 0))
  const legacyQuery = sheet.filterQuery?.trim().toLowerCase()

  const visible: number[] = []
  for (let row = 0; row < sheet.rowCount; row += 1) {
    if (hiddenSet.has(row)) continue
    if (row === 0 && hasHeader && (hasMulti || hasLegacy)) {
      visible.push(row)
      continue
    }
    if (hasMulti) {
      let ok = true
      for (const [colStr, rule] of filterEntries) {
        const col = Number(colStr)
        const key = cellKey(row, col)
        const disp = displayCellValue(sheet, key, workbook)
        const comp = computedCellValue(sheet, key, workbook)
        if (!cellMatchesColumnRule(disp, comp, rule)) {
          ok = false
          break
        }
      }
      if (ok) visible.push(row)
      continue
    }
    if (hasLegacy) {
      const val = displayCellValue(sheet, cellKey(row, legacyCol), workbook)
      if (legacyQuery ? val.toLowerCase().includes(legacyQuery) : val.trim() !== '') {
        visible.push(row)
      }
      continue
    }
    visible.push(row)
  }
  return visible
}

export interface ColumnProfile {
  column: number
  header: string
  totalRows: number
  nonBlankCount: number
  blankCount: number
  uniqueCount: number
  valueCounts: Array<{ label: string; count: number }>
  numericStats?: {
    count: number
    min: number
    median: number
    max: number
    mean: number
    bins: number[]
  }
}

export function profileColumn(
  sheet: WorkbookSheet,
  column: number,
  hasHeader = true,
  workbook?: FormulaWorkbook,
): ColumnProfile {
  // Determine last used row across the sheet so we don't scan 1000 blank rows
  let maxRow = 0
  for (const key of Object.keys(sheet.cells)) {
    const pos = parseCellKey(key)
    if (pos && pos.row > maxRow) maxRow = pos.row
  }
  const header = hasHeader
    ? displayCellValue(sheet, cellKey(0, column), workbook).trim() ||
      `Column ${column + 1}`
    : `Column ${column + 1}`
  const startRow = hasHeader ? 1 : 0
  const counts = new Map<string, number>()
  const numbers: number[] = []
  let blankCount = 0
  let nonBlankCount = 0

  for (let row = startRow; row <= maxRow; row += 1) {
    const key = cellKey(row, column)
    const disp = displayCellValue(sheet, key, workbook)
    const comp = computedCellValue(sheet, key, workbook).trim()
    if (disp.trim() === '') {
      blankCount += 1
      counts.set('(Blank)', (counts.get('(Blank)') ?? 0) + 1)
    } else {
      nonBlankCount += 1
      counts.set(disp, (counts.get(disp) ?? 0) + 1)
      const num = Number(comp)
      if (comp !== '' && Number.isFinite(num)) numbers.push(num)
    }
  }

  const valueCounts = [...counts.entries()]
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => {
      if (a.label === '(Blank)') return 1
      if (b.label === '(Blank)') return -1
      return b.count - a.count || a.label.localeCompare(b.label)
    })

  let numericStats: ColumnProfile['numericStats']
  if (numbers.length > 0) {
    const sorted = [...numbers].sort((a, b) => a - b)
    const min = sorted[0]
    const max = sorted[sorted.length - 1]
    const mid = Math.floor(sorted.length / 2)
    const median =
      sorted.length % 2 === 1
        ? sorted[mid]
        : (sorted[mid - 1] + sorted[mid]) / 2
    const mean = numbers.reduce((acc, n) => acc + n, 0) / numbers.length
    const binCount = 8
    const bins = new Array<number>(binCount).fill(0)
    const span = max - min
    for (const n of sorted) {
      const idx =
        span === 0
          ? 0
          : Math.min(binCount - 1, Math.floor(((n - min) / span) * binCount))
      bins[idx] += 1
    }
    numericStats = { count: numbers.length, min, median, max, mean, bins }
  }

  return {
    column,
    header,
    totalRows: Math.max(0, maxRow - startRow + 1),
    nonBlankCount,
    blankCount,
    uniqueCount: [...counts.keys()].filter(k => k !== '(Blank)').length,
    valueCounts,
    numericStats,
  }
}

// ---------------------------------------------------------------------------
// Data Cleanup & Paste Transpose (Wave 3)
// ---------------------------------------------------------------------------

export function removeDuplicatesInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
  keyColumns?: number[],
  hasHeader = true,
  workbook?: FormulaWorkbook,
): { sheet: WorkbookSheet; removedCount: number; keptCount: number } {
  const bounds = parseBounds(rangeToken)
  if (!bounds || isRangeLocked(sheet, rangeToken)) {
    return { sheet, removedCount: 0, keptCount: 0 }
  }
  const cols =
    keyColumns && keyColumns.length
      ? keyColumns
      : Array.from(
          { length: bounds.endColumn - bounds.startColumn + 1 },
          (_, i) => bounds.startColumn + i,
        )
  const startRow = hasHeader ? bounds.startRow + 1 : bounds.startRow
  const seen = new Set<string>()
  const keptRows: Array<Array<SheetCell | undefined>> = []
  let removedCount = 0

  for (let row = startRow; row <= bounds.endRow; row += 1) {
    const rowCells: Array<SheetCell | undefined> = []
    let allBlank = true
    for (let col = bounds.startColumn; col <= bounds.endColumn; col += 1) {
      const c = sheet.cells[cellKey(row, col)]
      rowCells.push(c)
      if (c && c.value.trim() !== '') allBlank = false
    }
    if (allBlank) {
      keptRows.push(rowCells)
      continue
    }
    const sig = cols
      .map(col =>
        computedCellValue(sheet, cellKey(row, col), workbook).trim().toLowerCase(),
      )
      .join('\u0001')
    if (seen.has(sig)) {
      removedCount += 1
    } else {
      seen.add(sig)
      keptRows.push(rowCells)
    }
  }

  if (removedCount === 0) {
    return { sheet, removedCount: 0, keptCount: keptRows.length }
  }

  const nextCells = { ...sheet.cells }
  for (let offset = 0; offset <= bounds.endRow - startRow; offset += 1) {
    const targetRow = startRow + offset
    const source = keptRows[offset]
    for (let col = bounds.startColumn; col <= bounds.endColumn; col += 1) {
      const k = cellKey(targetRow, col)
      const cell = source?.[col - bounds.startColumn]
      if (cell) nextCells[k] = cell
      else delete nextCells[k]
    }
  }
  return {
    sheet: { ...sheet, cells: nextCells },
    removedCount,
    keptCount: keptRows.length,
  }
}

export function splitTextToColumns(
  sheet: WorkbookSheet,
  rangeToken: string,
  delimiter = ',',
): { sheet: WorkbookSheet; rowsSplit: number; maxParts: number } {
  const bounds = parseBounds(rangeToken)
  if (!bounds || !delimiter || isRangeLocked(sheet, rangeToken)) {
    return { sheet, rowsSplit: 0, maxParts: 0 }
  }
  const sourceCol = bounds.startColumn
  const nextCells = { ...sheet.cells }
  let rowsSplit = 0
  let maxParts = 0
  let neededCols = sheet.columnCount

  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    const srcKey = cellKey(row, sourceCol)
    const srcCell = sheet.cells[srcKey]
    if (!srcCell || !srcCell.value.includes(delimiter)) continue
    const parts = srcCell.value.split(delimiter).map(part => part.trim())
    if (parts.length <= 1) continue
    rowsSplit += 1
    maxParts = Math.max(maxParts, parts.length)
    neededCols = Math.max(neededCols, sourceCol + parts.length)
    parts.forEach((part, idx) => {
      const targetKey = cellKey(row, sourceCol + idx)
      nextCells[targetKey] = {
        value: part,
        kind: deriveCellKind(part),
        style: idx === 0 ? srcCell.style : nextCells[targetKey]?.style,
      }
    })
  }

  if (rowsSplit === 0) return { sheet, rowsSplit: 0, maxParts: 0 }
  return {
    sheet: { ...sheet, columnCount: neededCols, cells: nextCells },
    rowsSplit,
    maxParts,
  }
}

export function trimWhitespaceInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): { sheet: WorkbookSheet; trimmedCount: number } {
  if (isRangeLocked(sheet, rangeToken)) return { sheet, trimmedCount: 0 }
  const keys = cellsInRange(rangeToken, 50_000)
  const nextCells = { ...sheet.cells }
  let trimmedCount = 0
  for (const key of keys) {
    const cell = nextCells[key]
    if (!cell || cell.kind === 'formula') continue
    const cleaned = cell.value.trim().replace(/\s+/g, ' ')
    if (cleaned !== cell.value) {
      trimmedCount += 1
      nextCells[key] = {
        ...cell,
        value: cleaned,
        kind: deriveCellKind(cleaned),
      }
    }
  }
  return trimmedCount
    ? { sheet: { ...sheet, cells: nextCells }, trimmedCount }
    : { sheet, trimmedCount: 0 }
}

export function transposeCopiedBlock(block: CopiedCellBlock): CopiedCellBlock {
  const transposed: CopiedCellBlock['cells'] = []
  const computedValues: string[][] = []
  for (let c = 0; c < block.columns; c += 1) {
    const row: CopiedCellBlock['cells'][number] = []
    const valRow: string[] = []
    for (let r = 0; r < block.rows; r += 1) {
      row.push(block.cells[r]?.[c])
      valRow.push(block.computedValues[r]?.[c] ?? '')
    }
    transposed.push(row)
    computedValues.push(valRow)
  }
  return {
    rows: block.columns,
    columns: block.rows,
    sourceStart: block.sourceStart,
    cells: transposed,
    computedValues,
  }
}
