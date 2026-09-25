import { expect, it } from 'vitest'
import { createBlankWorkbook } from './workbookModel'
import { createUniverSnapshot, documentFromEditorSnapshot, documentFromUniverSnapshot } from './univerAdapter'

it('keeps the chrome title while accepting new editor cells and sheet names', () => {
  const current = createBlankWorkbook('Renamed workbook')
  const snapshot = createUniverSnapshot(current)
  snapshot.name = 'Old mounted title'
  snapshot.sheets['sheet-1'].name = 'Revenue'
  snapshot.sheets['sheet-1'].cellData = { 0: { 0: { v: 'month' } }, 1: { 0: { v: 'Jan' } } }
  const result = documentFromEditorSnapshot(snapshot, current)
  expect(result.metadata.title).toBe('Renamed workbook')
  expect(result.workbook.sheets[0].name).toBe('Revenue')
  expect(result.workbook.sheets[0].cells.A2.value).toBe('Jan')
  expect(snapshot.name).toBe('Old mounted title')
  // Loading external snapshots keeps its existing title semantics.
  expect(documentFromUniverSnapshot(snapshot, current).metadata.title).toBe('Old mounted title')
})
