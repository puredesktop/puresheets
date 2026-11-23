/**
 * PureSheets Live Data Publisher & Cross-App Transform (Items 1 & 3).
 *
 * Extracts evaluated KPIs and tabular ranges from the active workbook sheet
 * and publishes them as a named LiveDataset (local + PureServer PureCloud)
 * so PureSlides, PureCanvas, and PureSite can bind & refresh live numbers in 1 click.
 */
import {
  publishDataset,
  sendTransformEnvelope,
  type CloudBridgeLike,
  type LiveDataset,
  type LiveMetric,
  type SuiteAppId,
  type TransformEnvelope,
} from '@purescience/platform-ui/suite'
import { cellKey, displayCellValue, type FormulaWorkbook } from './workbookModel'
import type { PureSheetsDocument, WorkbookSheet } from '../types'

type WorkbookInput = PureSheetsDocument | FormulaWorkbook

function resolveWorkbook(input: WorkbookInput): { formulaWorkbook: FormulaWorkbook; title?: string } {
  if ('workbook' in input) {
    return { formulaWorkbook: input.workbook, title: input.metadata.title }
  }
  return { formulaWorkbook: input }
}

function slugKey(raw: string, fallback: string): string {
  const s = raw
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 48)
  return s || fallback
}

export function extractSheetDataset(
  workbook: WorkbookInput,
  sheet: WorkbookSheet,
  options: { id?: string; name?: string; sourceDoc?: string } = {},
): Omit<LiveDataset, 'updatedAt'> {
  const { formulaWorkbook, title } = resolveWorkbook(workbook)
  let maxRow = -1
  let maxCol = -1
  for (const key of Object.keys(sheet.cells)) {
    const match = /^([A-Z]+)(\d+)$/i.exec(key)
    if (!match) continue
    const colLetters = match[1]!.toUpperCase()
    let col = 0
    for (let i = 0; i < colLetters.length; i++) {
      col = col * 26 + (colLetters.charCodeAt(i) - 64)
    }
    const r = Number(match[2]) - 1
    const c = col - 1
    if (r > maxRow) maxRow = r
    if (c > maxCol) maxCol = c
  }

  const rowLimit = Math.min(Math.max(maxRow + 1, 0), 50)
  const colLimit = Math.min(Math.max(maxCol + 1, 0), 16)
  const grid: string[][] = []
  for (let r = 0; r < rowLimit; r++) {
    const rowVals: string[] = []
    let hasValue = false
    for (let c = 0; c < colLimit; c++) {
      const k = cellKey(r, c)
      const val = sheet.cells[k] ? displayCellValue(sheet, k, formulaWorkbook).trim() : ''
      if (val) hasValue = true
      rowVals.push(val)
    }
    if (hasValue) grid.push(rowVals)
  }

  const headers = grid[0] ?? []
  const rows = grid.slice(1)
  const metrics: LiveMetric[] = []
  const seenKeys = new Set<string>()

  // Extract KPI pairs from 2+ column rows where col 0 is a label and col 1..last has a numeric/metric value
  for (let i = 0; i < rows.length && metrics.length < 12; i++) {
    const row = rows[i]!
    const label = row[0]?.trim()
    const value = [...row.slice(1)].reverse().find(v => v.trim().length > 0)?.trim()
    if (label && value) {
      let key = slugKey(label, `metric_${i + 1}`)
      if (seenKeys.has(key)) key = `${key}_${i + 1}`
      seenKeys.add(key)
      metrics.push({ key, label, value })
    }
  }

  const wbTitle = title || 'Workbook'
  const name = options.name?.trim() || `${wbTitle} — ${sheet.name}`
  const id = options.id || `ds-${slugKey(name, 'sheet')}`

  return {
    id,
    name,
    sourceApp: 'puresheets',
    sourceDoc: options.sourceDoc || `${wbTitle}.sheet`,
    rangeRef: `${sheet.name}!A1:${String.fromCharCode(64 + Math.max(1, colLimit))}${Math.max(1, rowLimit)}`,
    metrics,
    table: { headers, rows },
    updatedBy: 'PureSheets',
    scope: 'local',
  }
}

export async function publishSheetLiveData(
  workbook: WorkbookInput,
  sheet: WorkbookSheet,
  bridge?: CloudBridgeLike,
  options?: { id?: string; name?: string; sourceDoc?: string },
): Promise<LiveDataset> {
  const ds = extractSheetDataset(workbook, sheet, options)
  return publishDataset(ds, bridge)
}

export async function sendSheetToApp(
  workbook: WorkbookInput,
  sheet: WorkbookSheet,
  targetApp: SuiteAppId,
  bridge?: CloudBridgeLike,
): Promise<TransformEnvelope> {
  const ds = await publishSheetLiveData(workbook, sheet, bridge)
  return sendTransformEnvelope(
    {
      sourceApp: 'puresheets',
      targetApp,
      title: ds.name,
      subtitle: `Live bound from ${ds.sourceDoc} (${ds.rangeRef})`,
      datasetId: ds.id,
      metrics: ds.metrics,
      table: ds.table,
      sections: [
        {
          heading: `${sheet.name} Highlights`,
          bullets: ds.metrics.slice(0, 6).map(m => `${m.label}: ${m.value}`),
        },
      ],
    },
    bridge,
  )
}
