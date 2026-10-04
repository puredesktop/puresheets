import { expect, it } from 'vitest'
import { cellsInRange } from './workbookModel'
it('checks huge selection size before allocating keys', () => {
  expect(cellsInRange('A1:XFD1048576', 20000)).toEqual([])
  expect(cellsInRange('A1:A1000000000000000000000000', 20000)).toEqual([])
})
it('keeps exact-limit ranges, ordering and reversed endpoints', () => {
  expect(cellsInRange('B2:A1', 4)).toEqual(['A1', 'B1', 'A2', 'B2'])
  expect(cellsInRange('B2:A1', 3)).toEqual([])
  expect(cellsInRange('A1:B2')).toEqual(['A1', 'B1', 'A2', 'B2'])
})
