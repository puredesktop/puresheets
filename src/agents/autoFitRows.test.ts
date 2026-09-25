// @vitest-environment happy-dom
import { expect, it, vi } from 'vitest'
import { autoFitRowsHandler } from './handlers'
import { createBlankWorkbook, lockRange } from '../lib/workbookModel'
import { createUniverSnapshot, documentFromUniverSnapshot, withFreshUniverSnapshot } from '../lib/univerAdapter'
import type { SheetsAgentToolContext } from './catalog'

it('targets bounded rows on the requested sheet and refuses unavailable/protected fitting', async () => {
  let document = createBlankWorkbook('Fit rows')
  const fit = vi.fn(async () => true)
  const context: SheetsAgentToolContext = { get document() { return document }, filePath: '', selectedRange: 'A1',
    setDocument(update) { document = update(document) }, autoFitRows: fit }
  await autoFitRowsHandler(context, { startRow: 2, rowCount: 3 })
  expect(fit).toHaveBeenCalledWith(document.workbook.sheets[0].id, [1, 2, 3])
  expect(document.workbook.agentLog?.[0].tool).toBe('autoFitRows')
  fit.mockClear()
  for (const args of [{ startRow: 0, rowCount: 1 }, { startRow: 1, rowCount: 101 }, { startRow: 2.5, rowCount: 1 }])
    await autoFitRowsHandler(context, args)
  expect(fit).not.toHaveBeenCalled()
  document.workbook.sheets[0] = lockRange(document.workbook.sheets[0], 'Z10')
  await autoFitRowsHandler(context, { startRow: 1, rowCount: 1 })
  expect(fit).not.toHaveBeenCalled()
  document = createBlankWorkbook()
  fit.mockResolvedValue(false)
  const result = await autoFitRowsHandler(context, { startRow: 1, rowCount: 1 })
  expect(JSON.stringify(result)).toContain('unavailable')
  expect(document.workbook.agentLog ?? []).toHaveLength(0)
})

it('retains measured auto height across model-driven reload but respects forced height', () => {
  const original = createBlankWorkbook()
  const snapshot = createUniverSnapshot(original)
  snapshot.sheets['sheet-1'].rowData = { 0: { h: 24, ah: 85, ia: 1 }, 1: { h: 32, ah: 90, ia: 0 } }
  const loaded = documentFromUniverSnapshot(snapshot, original)
  expect(loaded.workbook.sheets[0].rowHeights).toEqual({ 0: 85, 1: 32 })
  const rebuilt = withFreshUniverSnapshot(loaded)
  expect(rebuilt.engine.snapshot).toMatchObject({ sheets: { 'sheet-1': { rowData: { 0: { h: 85 } } } } })
})
