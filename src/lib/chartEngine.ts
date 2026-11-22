import type {
  ChartSeries,
  WorkbookChart,
  WorkbookChartType,
  WorkbookSheet,
} from '../types'
import {
  cellKey,
  columnLabel,
  computedCellValue,
  displayCellValue,
  normalizeRangeToken,
  parseCellKey,
  type FormulaWorkbook,
} from './workbookModel'

export const CHART_SERIES_COLORS = [
  '#2f6ea2',
  '#377d5a',
  '#b06b1f',
  '#8a5aa8',
  '#a8385a',
  '#3f7f8f',
  '#6c7a52',
  '#946f43',
]

function parseBounds(rangeToken: string) {
  const norm = normalizeRangeToken(rangeToken)
  if (!norm) return null
  const [a, b] = norm.split(':')
  const s = parseCellKey(a)
  const e = parseCellKey(b ?? a)
  if (!s || !e) return null
  return {
    normalized: norm,
    startRow: Math.min(s.row, e.row),
    startColumn: Math.min(s.column, e.column),
    endRow: Math.max(s.row, e.row),
    endColumn: Math.max(s.column, e.column),
  }
}

function parseNumeric(raw: string): number | null {
  const cleaned = raw.trim().replace(/[$,%]/g, '').replace(/,/g, '')
  if (!cleaned) return null
  const n = Number(cleaned)
  return Number.isFinite(n) ? n : null
}

/**
 * Extract single- or multi-series chart data from any selected range.
 * - 1 column: values only, labels = cell keys (A1, A2, …)
 * - 2+ columns: col 0 = category labels (or numeric X for scatter), cols 1..N = series.
 *   Automatically detects whether row 0 is a header row (non-numeric value cols).
 */
export function extractChartData(
  sheet: WorkbookSheet,
  rangeToken: string,
  type: WorkbookChartType,
  workbook?: FormulaWorkbook,
): {
  range: string
  title: string
  labels: string[]
  values: number[]
  series: ChartSeries[]
} | null {
  const bounds = parseBounds(rangeToken)
  if (!bounds) return null
  const colSpan = bounds.endColumn - bounds.startColumn + 1
  const rowSpan = bounds.endRow - bounds.startRow + 1

  if (colSpan === 1) {
    // Check if first row is a text header
    const firstRaw = computedCellValue(
      sheet,
      cellKey(bounds.startRow, bounds.startColumn),
      workbook,
    )
    const hasHeader = rowSpan > 1 && parseNumeric(firstRaw) === null && firstRaw.trim() !== ''
    const startRow = hasHeader ? bounds.startRow + 1 : bounds.startRow
    const labels: string[] = []
    const values: number[] = []
    for (let row = startRow; row <= bounds.endRow; row += 1) {
      const k = cellKey(row, bounds.startColumn)
      const n = parseNumeric(computedCellValue(sheet, k, workbook))
      if (n === null) continue
      labels.push(k)
      values.push(n)
    }
    if (!values.length) return null
    const seriesName = hasHeader
      ? displayCellValue(sheet, cellKey(bounds.startRow, bounds.startColumn), workbook).trim()
      : `Col ${columnLabel(bounds.startColumn)}`
    return {
      range: bounds.normalized,
      title: hasHeader ? seriesName : `${sheet.name} (${bounds.normalized})`,
      labels,
      values,
      series: [{ name: seriesName, values, color: CHART_SERIES_COLORS[0] }],
    }
  }

  // Multi-column range: check whether first row is header
  let firstRowHasNumbersInValueCols = false
  for (let col = bounds.startColumn + 1; col <= bounds.endColumn; col += 1) {
    const raw = computedCellValue(sheet, cellKey(bounds.startRow, col), workbook)
    if (parseNumeric(raw) !== null) {
      firstRowHasNumbersInValueCols = true
      break
    }
  }
  const hasHeader = rowSpan > 1 && !firstRowHasNumbersInValueCols
  const dataStartRow = hasHeader ? bounds.startRow + 1 : bounds.startRow

  const valueCols: Array<{ col: number; name: string }> = []
  for (let col = bounds.startColumn + 1; col <= bounds.endColumn; col += 1) {
    const headerText = hasHeader
      ? displayCellValue(sheet, cellKey(bounds.startRow, col), workbook).trim()
      : ''
    valueCols.push({
      col,
      name: headerText || `Series ${col - bounds.startColumn}`,
    })
  }

  const labels: string[] = []
  const seriesBuffers: number[][] = valueCols.map(() => [])

  for (let row = dataStartRow; row <= bounds.endRow; row += 1) {
    const rowLabel =
      displayCellValue(sheet, cellKey(row, bounds.startColumn), workbook).trim() ||
      `Row ${row + 1}`
    const parsedRow = valueCols.map(({ col }) =>
      parseNumeric(computedCellValue(sheet, cellKey(row, col), workbook)),
    )
    if (type === 'scatter' && parseNumeric(rowLabel) === null) continue
    if (parsedRow.every(v => v === null)) continue
    labels.push(rowLabel)
    parsedRow.forEach((v, idx) => {
      seriesBuffers[idx].push(v ?? 0)
    })
  }

  if (!labels.length) return null

  const series: ChartSeries[] = valueCols
    .map((vc, idx) => ({
      name: vc.name,
      values: seriesBuffers[idx],
      color: CHART_SERIES_COLORS[idx % CHART_SERIES_COLORS.length],
    }))
    .filter(s => s.values.length === labels.length)

  if (!series.length) return null

  const title =
    series.length === 1 && hasHeader
      ? series[0].name
      : `${sheet.name} (${bounds.normalized})`

  return {
    range: bounds.normalized,
    title,
    labels,
    values: series[0].values,
    series,
  }
}

/**
 * Recompute all charts on a sheet from current cell/formula values so charts
 * stay live-linked whenever data changes.
 */
export function recomputeSheetCharts(
  sheet: WorkbookSheet,
  workbook?: FormulaWorkbook,
): WorkbookSheet {
  if (!sheet.charts?.length) return sheet
  let changed = false
  const nextCharts = sheet.charts.map(chart => {
    const extracted = extractChartData(sheet, chart.range, chart.type, workbook)
    if (!extracted) return chart
    const sameLabels =
      extracted.labels.length === chart.labels.length &&
      extracted.labels.every((l, i) => l === chart.labels[i])
    const sameValues =
      extracted.values.length === chart.values.length &&
      extracted.values.every((v, i) => v === chart.values[i])
    const sameSeries =
      (chart.series?.length ?? 0) === extracted.series.length &&
      extracted.series.every(
        (s, i) =>
          chart.series?.[i]?.name === s.name &&
          chart.series?.[i]?.values.length === s.values.length &&
          chart.series?.[i]?.values.every((v, j) => v === s.values[j]),
      )
    if (sameLabels && sameValues && sameSeries) return chart
    changed = true
    return {
      ...chart,
      labels: extracted.labels,
      values: extracted.values,
      series: extracted.series,
    }
  })
  return changed ? { ...sheet, charts: nextCharts } : sheet
}

export function chartSeriesOf(chart: WorkbookChart): ChartSeries[] {
  if (chart.series?.length) return chart.series
  return [
    {
      name: chart.title || 'Series 1',
      values: chart.values,
      color: CHART_SERIES_COLORS[0],
    },
  ]
}
