import type {
  CellStyle,
  PivotAggregate,
  PivotDefinition,
  PivotFieldRef,
  SheetCell,
  WorkbookSheet,
} from '../types'
import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT } from '../constants'
import {
  cellKey,
  columnLabel,
  computedCellValue,
  deriveCellKind,
  displayCellValue,
  normalizeRangeToken,
  parseCellKey,
  type FormulaWorkbook,
} from './workbookModel'

export interface InferredPivotField extends PivotFieldRef {
  isNumeric: boolean
  distinctCount: number
  sampleValues: string[]
}

function parseBounds(rangeToken: string) {
  const norm = normalizeRangeToken(rangeToken)
  if (!norm) return null
  const [a, b] = norm.split(':')
  const s = parseCellKey(a)
  const e = parseCellKey(b ?? a)
  if (!s || !e) return null
  return {
    startRow: Math.min(s.row, e.row),
    startColumn: Math.min(s.column, e.column),
    endRow: Math.max(s.row, e.row),
    endColumn: Math.max(s.column, e.column),
  }
}

export function inferPivotFields(
  source: WorkbookSheet,
  sourceRange: string,
  workbook?: FormulaWorkbook,
): InferredPivotField[] {
  const bounds = parseBounds(sourceRange)
  if (!bounds) return []
  const fields: InferredPivotField[] = []
  for (let col = bounds.startColumn; col <= bounds.endColumn; col += 1) {
    const header =
      displayCellValue(source, cellKey(bounds.startRow, col), workbook).trim() ||
      `Col ${columnLabel(col)}`
    let numericCount = 0
    let nonBlank = 0
    const distinct = new Set<string>()
    for (let row = bounds.startRow + 1; row <= bounds.endRow; row += 1) {
      const comp = computedCellValue(source, cellKey(row, col), workbook).trim()
      if (!comp) continue
      nonBlank += 1
      distinct.add(comp)
      if (Number.isFinite(Number(comp))) numericCount += 1
    }
    fields.push({
      columnIndex: col,
      label: header,
      isNumeric: nonBlank > 0 && numericCount / nonBlank >= 0.7,
      distinctCount: distinct.size,
      sampleValues: [...distinct].slice(0, 6),
    })
  }
  return fields
}

export function computePivotSourceSignature(
  source: WorkbookSheet,
  sourceRange: string,
  workbook?: FormulaWorkbook,
): string {
  const bounds = parseBounds(sourceRange)
  if (!bounds) return ''
  let hash = 2166136261
  const maxCells = 50_000
  let count = 0
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let col = bounds.startColumn; col <= bounds.endColumn; col += 1) {
      if (++count > maxCells) break
      const val = computedCellValue(source, cellKey(row, col), workbook)
      for (let i = 0; i < val.length; i += 1) {
        hash ^= val.charCodeAt(i)
        hash = Math.imul(hash, 16777619)
      }
      hash ^= 0x1f
      hash = Math.imul(hash, 16777619)
    }
  }
  return (hash >>> 0).toString(36)
}

export function isPivotStale(
  pivotSheet: WorkbookSheet,
  sourceSheet: WorkbookSheet | undefined,
  workbook?: FormulaWorkbook,
): boolean {
  if (!pivotSheet.pivot) return false
  if (!sourceSheet) return true
  if (!pivotSheet.pivot.sourceSignature) return false
  const current = computePivotSourceSignature(
    sourceSheet,
    pivotSheet.pivot.sourceRange,
    workbook,
  )
  return current !== pivotSheet.pivot.sourceSignature
}

function aggregateValues(
  nums: number[],
  rawCount: number,
  agg: PivotAggregate,
  grandSum: number,
): number {
  if (agg === 'count') return rawCount
  if (!nums.length) return 0
  if (agg === 'sum') return nums.reduce((a, b) => a + b, 0)
  if (agg === 'avg') return nums.reduce((a, b) => a + b, 0) / nums.length
  if (agg === 'min') return Math.min(...nums)
  if (agg === 'max') return Math.max(...nums)
  if (agg === 'percent-of-total') {
    const sum = nums.reduce((a, b) => a + b, 0)
    return grandSum === 0 ? 0 : sum / grandSum
  }
  return 0
}

function aggLabel(agg: PivotAggregate): string {
  switch (agg) {
    case 'sum':
      return 'Sum'
    case 'count':
      return 'Count'
    case 'avg':
      return 'Avg'
    case 'min':
      return 'Min'
    case 'max':
      return 'Max'
    case 'percent-of-total':
      return '% Total'
  }
}

const HEADER_STYLE: CellStyle = {
  bold: true,
  fillColor: '#eef1f5',
  borderSides: ['bottom'],
}

const TOTAL_STYLE: CellStyle = {
  bold: true,
  fillColor: '#f6f7f9',
  borderSides: ['top', 'bottom'],
}

const VALUE_NUMBER_STYLE: CellStyle = {
  align: 'right',
  numberFormat: 'number',
  decimals: 2,
}

const VALUE_PERCENT_STYLE: CellStyle = {
  align: 'right',
  numberFormat: 'percent',
  decimals: 1,
}

const VALUE_COUNT_STYLE: CellStyle = {
  align: 'right',
  numberFormat: 'number',
  decimals: 0,
}

function valueStyle(agg: PivotAggregate, isTotal = false): CellStyle {
  const base =
    agg === 'percent-of-total'
      ? VALUE_PERCENT_STYLE
      : agg === 'count'
      ? VALUE_COUNT_STYLE
      : VALUE_NUMBER_STYLE
  return isTotal ? { ...base, ...TOTAL_STYLE } : base
}

function roundClean(n: number): string {
  if (!Number.isFinite(n)) return '0'
  return String(Number(n.toFixed(6)))
}

/**
 * Pure pivot computation: produces a normal, protected derived WorkbookSheet
 * with row grouping, optional column cross-tab, multi-value aggregates, filters,
 * and grand totals.
 */
export function computePivotSheet(
  source: WorkbookSheet,
  def: PivotDefinition,
  targetMeta: { id: string; name: string },
  workbook?: FormulaWorkbook,
): WorkbookSheet {
  const bounds = parseBounds(def.sourceRange)
  const cells: Record<string, SheetCell> = {}
  const signature = computePivotSourceSignature(source, def.sourceRange, workbook)
  const refreshedDef: PivotDefinition = {
    ...def,
    refreshedAt: new Date().toISOString(),
    sourceSignature: signature,
  }

  if (!bounds || (!def.rows.length && !def.values.length)) {
    cells['A1'] = {
      value: 'Configure row and value fields in the Pivot inspector.',
      kind: 'text',
      style: { italic: true, textColor: '#666c75' },
    }
    return {
      id: targetMeta.id,
      name: targetMeta.name,
      rowCount: DEFAULT_ROW_COUNT,
      columnCount: DEFAULT_COLUMN_COUNT,
      cells,
      pivot: refreshedDef,
    }
  }

  // Extract filtered data rows
  const records: Array<Map<number, string>> = []
  const maxRows = Math.min(bounds.endRow, bounds.startRow + 25_000)
  for (let row = bounds.startRow + 1; row <= maxRows; row += 1) {
    const rec = new Map<number, string>()
    let anyNonBlank = false
    for (let col = bounds.startColumn; col <= bounds.endColumn; col += 1) {
      const v = computedCellValue(source, cellKey(row, col), workbook).trim()
      rec.set(col, v)
      if (v !== '') anyNonBlank = true
    }
    if (!anyNonBlank) continue
    if (def.filters?.length) {
      const pass = def.filters.every(f => {
        const val = rec.get(f.columnIndex) || '(Blank)'
        return f.selectedValues.includes(val)
      })
      if (!pass) continue
    }
    records.push(rec)
  }

  const rowFields = def.rows
  const colFields = def.columns ?? []
  const valueFields = def.values.length
    ? def.values
    : [{ columnIndex: bounds.startColumn, label: 'Rows', aggregate: 'count' as PivotAggregate }]

  // Collect distinct row-group tuples and column-group tuples
  const rowKeys: string[] = []
  const rowTupleMap = new Map<string, string[]>()
  const colKeys: string[] = []
  const colTupleMap = new Map<string, string[]>()

  const buckets = new Map<string, Array<Map<number, string>>>()
  const rowTotals = new Map<string, Array<Map<number, string>>>()
  const colTotals = new Map<string, Array<Map<number, string>>>()

  for (const rec of records) {
    const rTuple = rowFields.length
      ? rowFields.map(f => rec.get(f.columnIndex) || '(Blank)')
      : ['All']
    const rKey = rTuple.join('\u0001')
    if (!rowTupleMap.has(rKey)) {
      rowTupleMap.set(rKey, rTuple)
      rowKeys.push(rKey)
    }

    const cTuple = colFields.length
      ? colFields.map(f => rec.get(f.columnIndex) || '(Blank)')
      : ['']
    const cKey = cTuple.join('\u0001')
    if (!colTupleMap.has(cKey)) {
      colTupleMap.set(cKey, cTuple)
      colKeys.push(cKey)
    }

    const bucketKey = `${rKey}\u0002${cKey}`
    const bucket = buckets.get(bucketKey) ?? []
    bucket.push(rec)
    buckets.set(bucketKey, bucket)

    const rBucket = rowTotals.get(rKey) ?? []
    rBucket.push(rec)
    rowTotals.set(rKey, rBucket)

    const cBucket = colTotals.get(cKey) ?? []
    cBucket.push(rec)
    colTotals.set(cKey, cBucket)
  }

  rowKeys.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
  colKeys.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))

  // Grand sum per value field (for percent-of-total)
  const grandSums = valueFields.map(vf => {
    let sum = 0
    for (const rec of records) {
      const n = Number(rec.get(vf.columnIndex) ?? '')
      if (Number.isFinite(n)) sum += n
    }
    return sum
  })

  const rowHeaderColCount = Math.max(1, rowFields.length)
  const hasCrossTab = colFields.length > 0

  // Row 0: Header row
  const headers: string[] = rowFields.length
    ? rowFields.map(f => f.label)
    : ['Summary']

  const dataColumnSpecs: Array<{
    cKey: string | null // null = grand total across columns
    vfIndex: number
    label: string
  }> = []

  for (const cKey of colKeys) {
    const cLabel = (colTupleMap.get(cKey) ?? []).filter(Boolean).join(' / ')
    valueFields.forEach((vf, vfIndex) => {
      const base = `${aggLabel(vf.aggregate)} of ${vf.label}`
      dataColumnSpecs.push({
        cKey,
        vfIndex,
        label: hasCrossTab ? `${cLabel} · ${base}` : base,
      })
    })
  }

  if (hasCrossTab) {
    valueFields.forEach((vf, vfIndex) => {
      dataColumnSpecs.push({
        cKey: null,
        vfIndex,
        label: `Total · ${aggLabel(vf.aggregate)} of ${vf.label}`,
      })
    })
  }

  headers.forEach((h, c) => {
    cells[cellKey(0, c)] = { value: h, kind: 'text', style: HEADER_STYLE }
  })
  dataColumnSpecs.forEach((spec, idx) => {
    cells[cellKey(0, rowHeaderColCount + idx)] = {
      value: spec.label,
      kind: 'text',
      style: { ...HEADER_STYLE, align: 'right' },
    }
  })

  // Data rows
  rowKeys.forEach((rKey, rIdx) => {
    const row = rIdx + 1
    const tuple = rowTupleMap.get(rKey) ?? ['']
    tuple.forEach((label, c) => {
      cells[cellKey(row, c)] = {
        value: label,
        kind: deriveCellKind(label),
        style: rIdx % 2 === 1 ? { fillColor: '#fafbfc' } : undefined,
      }
    })
    dataColumnSpecs.forEach((spec, idx) => {
      const vf = valueFields[spec.vfIndex]
      const recs =
        spec.cKey === null
          ? rowTotals.get(rKey) ?? []
          : buckets.get(`${rKey}\u0002${spec.cKey}`) ?? []
      const nums = recs
        .map(rec => Number(rec.get(vf.columnIndex) ?? ''))
        .filter(Number.isFinite)
      const val = aggregateValues(nums, recs.length, vf.aggregate, grandSums[spec.vfIndex])
      const text = roundClean(val)
      cells[cellKey(row, rowHeaderColCount + idx)] = {
        value: text,
        kind: 'number',
        style: {
          ...valueStyle(vf.aggregate, spec.cKey === null),
          ...(rIdx % 2 === 1 && spec.cKey !== null ? { fillColor: '#fafbfc' } : {}),
        },
      }
    })
  })

  // Grand total row
  const totalRow = rowKeys.length + 1
  for (let c = 0; c < rowHeaderColCount; c += 1) {
    cells[cellKey(totalRow, c)] = {
      value: c === 0 ? 'Grand Total' : '',
      kind: c === 0 ? 'text' : 'blank',
      style: TOTAL_STYLE,
    }
  }
  dataColumnSpecs.forEach((spec, idx) => {
    const vf = valueFields[spec.vfIndex]
    const recs =
      spec.cKey === null ? records : colTotals.get(spec.cKey) ?? []
    const nums = recs
      .map(rec => Number(rec.get(vf.columnIndex) ?? ''))
      .filter(Number.isFinite)
    const val = aggregateValues(nums, recs.length, vf.aggregate, grandSums[spec.vfIndex])
    cells[cellKey(totalRow, rowHeaderColCount + idx)] = {
      value: roundClean(val),
      kind: 'number',
      style: valueStyle(vf.aggregate, true),
    }
  })

  const totalCols = rowHeaderColCount + dataColumnSpecs.length
  const usedRange = `A1:${cellKey(totalRow, Math.max(0, totalCols - 1))}`

  return {
    id: targetMeta.id,
    name: targetMeta.name,
    rowCount: Math.max(DEFAULT_ROW_COUNT, totalRow + 10),
    columnCount: Math.max(DEFAULT_COLUMN_COUNT, totalCols + 4),
    cells,
    frozenRows: 1,
    protection: { lockedRanges: [usedRange] },
    pivot: refreshedDef,
  }
}
