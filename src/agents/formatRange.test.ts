// @vitest-environment happy-dom
import { expect, it } from 'vitest'
import { formatRangeHandler } from './handlers'
import type { SheetsAgentToolContext } from './catalog'
import { createBlankWorkbook, lockRange } from '../lib/workbookModel'
import { loadSheetsDocument, serializeSheetsDocument } from '../lib/sheetsDocument'
import { withFreshUniverSnapshot } from '../lib/univerAdapter'

function fixture() {
  let document = createBlankWorkbook('Readable table')
  document.workbook.sheets[0].cells.A1 = { value: 'A long procedure description', kind: 'text' }
  const context: SheetsAgentToolContext = {
    get document() { return document }, filePath: '', selectedRange: 'A1:B2',
    setDocument(update) { document = update(document) },
  }
  return context
}

it('formats readable prose without changing values and preserves it through snapshot save/reopen', async () => {
  const context = fixture()
  await formatRangeHandler(context, { range: 'A1:B2', wrap: 'wrap', columnWidth: 240, rowHeight: 96 })
  const sheet = context.document!.workbook.sheets[0]
  expect(sheet.cells.A1).toMatchObject({ value: 'A long procedure description', style: { wrap: 'wrap' } })
  expect(sheet.columnWidths).toMatchObject({ 0: 240, 1: 240 })
  expect(sheet.rowHeights).toMatchObject({ 0: 96, 1: 96 })
  expect(context.document!.workbook.agentLog?.[0].tool).toBe('formatRange')
  const reopened = loadSheetsDocument(serializeSheetsDocument(withFreshUniverSnapshot(context.document!))).workbook.sheets[0]
  expect(reopened.cells.A1.style?.wrap).toBe('wrap')
  expect(reopened.columnWidths?.['0']).toBe(240)
  expect(reopened.rowHeights?.['0']).toBe(96)
  for (const wrap of ['clip', 'overflow'] as const) {
    await formatRangeHandler(context, { range: 'A1', wrap })
    expect(loadSheetsDocument(serializeSheetsDocument(withFreshUniverSnapshot(context.document!))).workbook.sheets[0].cells.A1.style?.wrap).toBe(wrap)
  }
})

it('rejects invalid sizes and refuses whole-row/column resizing on protected sheets', async () => {
  const context = fixture()
  const before = JSON.stringify(context.document)
  for (const args of [{ columnWidth: 1000 }, { rowHeight: NaN }, { wrap: 'bogus' }]) {
    await formatRangeHandler(context, { range: 'A1:B2', ...args })
    expect(JSON.stringify(context.document)).toBe(before)
  }
  context.document!.workbook.sheets[0] = lockRange(context.document!.workbook.sheets[0], 'Z10')
  const locked = JSON.stringify(context.document)
  await formatRangeHandler(context, { range: 'A1:B2', columnWidth: 240 })
  expect(JSON.stringify(context.document)).toBe(locked)
})
