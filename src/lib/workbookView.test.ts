// @vitest-environment happy-dom
import { expect, it } from 'vitest'
import { addSheet, createBlankWorkbook } from './workbookModel'
import { createUniverSnapshot } from './univerAdapter'
import { requestedSheetId } from './workbookView'

it('restores an explicitly added sheet and ignores invalid or absent saved selection', () => {
  const doc = addSheet(createBlankWorkbook(), 'Procedure')
  const snapshot = createUniverSnapshot(doc)
  expect(requestedSheetId(snapshot)).toBe(doc.workbook.sheets[1].id)
  snapshot.custom = { puresheets: { activeSheetId: 'missing' } }
  expect(requestedSheetId(snapshot)).toBeNull()
  snapshot.custom = {}
  expect(requestedSheetId(snapshot)).toBeNull()
})
