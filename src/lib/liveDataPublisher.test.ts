// @vitest-environment happy-dom
import { describe, expect, it, beforeEach } from 'vitest'
import { createBlankWorkbook, upsertCell } from './workbookModel'
import { extractSheetDataset, sendSheetToApp } from './liveDataPublisher'
import { listPendingTransforms } from '@purescience/platform-ui/suite'

describe('PureSheets Live Data Publisher & Transform Pipeline', () => {
  beforeEach(() => {
    window.localStorage.clear()
  })

  it('extracts evaluated metrics and table rows from a PureSheets workbook', () => {
    let doc = createBlankWorkbook('FY26 Model')
    let sheet = doc.workbook.sheets[0]!
    sheet = upsertCell(sheet, 'A1', 'Region')
    sheet = upsertCell(sheet, 'B1', 'Q4 Revenue')
    sheet = upsertCell(sheet, 'A2', 'North America')
    sheet = upsertCell(sheet, 'B2', '2450')
    sheet = upsertCell(sheet, 'A3', 'Europe')
    sheet = upsertCell(sheet, 'B3', '1480')
    sheet = upsertCell(sheet, 'A4', 'Total')
    sheet = upsertCell(sheet, 'B4', '=SUM(B2:B3)')
    doc = { ...doc, workbook: { ...doc.workbook, sheets: [sheet] } }

    const ds = extractSheetDataset(doc, sheet)
    expect(ds.name).toBe('FY26 Model — Sheet 1')
    expect(ds.table.headers).toEqual(['Region', 'Q4 Revenue'])
    expect(ds.table.rows).toHaveLength(3)
    expect(ds.metrics.find(m => m.key === 'total')?.value).toBe('3930')
  })

  it('publishes live dataset and queues a TransformEnvelope for PureSlides', async () => {
    let doc = createBlankWorkbook('Board KPIs')
    let sheet = doc.workbook.sheets[0]!
    sheet = upsertCell(sheet, 'A1', 'Metric')
    sheet = upsertCell(sheet, 'B1', 'Value')
    sheet = upsertCell(sheet, 'A2', 'ARR')
    sheet = upsertCell(sheet, 'B2', '$4.8M')
    doc = { ...doc, workbook: { ...doc.workbook, sheets: [sheet] } }

    const env = await sendSheetToApp(doc, sheet, 'pureslides')
    expect(env.datasetId).toBeTruthy()
    expect(listPendingTransforms('pureslides')).toHaveLength(1)
  })
})
