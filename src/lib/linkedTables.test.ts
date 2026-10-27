import { expect, it } from 'vitest'
import { createBlankWorkbook, upsertCell, setCellStyle } from './workbookModel'
import { workbookLinkedSnapshot } from './linkedTables'
import {
  serializeSheetsDocument,
  serializeSheetsHtmlDocument,
  loadSheetsDocument,
} from './sheetsDocument'
const ref = {
  version: 1 as const,
  path: '/Budget.sheets',
  file: 'workbook.json',
  tableId: 'sheet-1',
  range: 'A1:B2',
}
it('publishes calculated, formatted values and preserves stable sheet identity', () => {
  const doc = createBlankWorkbook('Budget')
  let sheet = doc.workbook.sheets[0]
  sheet = upsertCell(sheet, 'A1', 'Research')
  sheet = upsertCell(sheet, 'B1', '1000')
  sheet = upsertCell(sheet, 'B2', '=B1*2')
  sheet = setCellStyle(sheet, 'B2', { numberFormat: 'currency', bold: true })
  doc.workbook.sheets[0] = sheet
  expect(workbookLinkedSnapshot(doc, ref).rows[1][1]).toEqual({
    text: '$2,000.00',
    bold: true,
  })
  const raw = serializeSheetsDocument(doc),
    saved = JSON.parse(raw)
  expect(saved.linkedTables.tables[0].id).toBe('sheet-1')
  expect(saved.linkedTables.tables[0].cells.B2.text).toBe('$2,000.00')
  expect(loadSheetsDocument(raw).workbook.sheets[0].cells.B2.value).toBe(
    '=B1*2',
  )
  expect(serializeSheetsHtmlDocument(doc)).toContain('data-linked-table-source')
})
it('refreshes projections after cell edits and sheet renames, excluding legacy agent drafts', () => {
  const doc = createBlankWorkbook('Budget')
  doc.workbook.sheets[0] = upsertCell(doc.workbook.sheets[0], 'B1', '12')
  doc.workbook.sheets[0].name = 'Updated name'
  expect(workbookLinkedSnapshot(doc, ref).title).toContain('Updated name')
  expect(workbookLinkedSnapshot(doc, ref).rows[0][1].text).toBe('12')
})
