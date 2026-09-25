/**
 * Phase S3 — data features at the model layer: comments, protection,
 * validation, conditional formatting, filter query, find/replace.
 */
import { describe, expect, it } from 'vitest'
import {
  clearRange,
  copyRangeCells,
  createBlankWorkbook,
  deleteRow,
  fillRange,
  findMatches,
  insertColumn,
  insertRow,
  isCellLocked,
  isRangeLocked,
  lockRange,
  pasteCellBlock,
  replaceAllInSheet,
  replaceInRange,
  setCellComment,
  setFilterQuery,
  sortRange,
  toggleFilterRow,
  unlockRange,
  upsertCell,
  visibleRowsForSheet,
} from './workbookModel'
import {
  addConditionalFormatRule,
  addValidationRule,
  clearConditionalFormatsInRange,
  clearValidationInRange,
  conditionMatches,
  conditionalStyleForCell,
  invalidCellFlagsForSheet,
  invalidCellsForSheet,
  validateValue,
  validationRuleForCell,
} from './sheetDataFeatures'
import type { WorkbookSheet } from '../types'

function sheetWith(cells: Record<string, string>): WorkbookSheet {
  let sheet = createBlankWorkbook('Data features').workbook.sheets[0]
  for (const [key, value] of Object.entries(cells)) {
    sheet = upsertCell(sheet, key, value)
  }
  return sheet
}

// ---------------------------------------------------------------------------
// Comments
// ---------------------------------------------------------------------------

describe('cell comments', () => {
  it('sets, replaces, and removes notes', () => {
    let sheet = sheetWith({ A1: 'x' })
    sheet = setCellComment(sheet, 'A1', 'check this')
    expect(sheet.comments).toEqual({ A1: 'check this' })
    sheet = setCellComment(sheet, 'A1', 'checked ✓')
    expect(sheet.comments).toEqual({ A1: 'checked ✓' })
    sheet = setCellComment(sheet, 'A1', '')
    expect(sheet.comments).toBeUndefined()
  })

  it('comment keys shift with row and column structure edits', () => {
    let sheet = sheetWith({ A2: 'watch', B3: 'other' })
    sheet = setCellComment(sheet, 'A2', 'moves down')
    sheet = setCellComment(sheet, 'B3', 'stays put relative')
    const inserted = insertRow(sheet, 1) // above row 2
    expect(inserted.comments).toEqual({
      A3: 'moves down',
      B4: 'stays put relative',
    })
    const widened = insertColumn(inserted, 0) // before column A
    expect(widened.comments).toEqual({
      B3: 'moves down',
      C4: 'stays put relative',
    })
    const collapsed = deleteRow(widened, 2) // delete the commented row 3
    expect(collapsed.comments).toEqual({ C3: 'stays put relative' })
  })
})

// ---------------------------------------------------------------------------
// Protection
// ---------------------------------------------------------------------------

describe('protected ranges', () => {
  const base = lockRange(
    sheetWith({ B2: 'locked value', C2: 'free value' }),
    'B1:B4',
  )

  it('reports locked cells and ranges', () => {
    expect(isCellLocked(base, 'B2')).toBe(true)
    expect(isCellLocked(base, 'C2')).toBe(false)
    expect(isRangeLocked(base, 'A1:C3')).toBe(true)
    expect(isRangeLocked(base, 'C1:D4')).toBe(false)
  })

  it('upsert, clear, paste, fill, sort, and replace respect locks', () => {
    // upsert is a no-op on the locked cell
    expect(upsertCell(base, 'B2', 'overwrite')).toBe(base)
    // clear skips locked cells but clears the rest
    const cleared = clearRange(base, 'B2:C2')
    expect(cleared.cells.B2.value).toBe('locked value')
    expect(cleared.cells.C2).toBeUndefined()
    // paste skips locked targets
    const block = copyRangeCells(base, 'C2:C2')
    if (!block) throw new Error('expected block')
    const pasted = pasteCellBlock(base, 'B2', block, 'all')
    expect(pasted.cells.B2.value).toBe('locked value')
    // fill skips locked targets
    let fillSheet = lockRange(sheetWith({ A1: '1', A2: '2' }), 'A4:A4')
    fillSheet = fillRange(fillSheet, 'A1:A2', 'A3:A5')
    expect(fillSheet.cells.A3.value).toBe('3')
    expect(fillSheet.cells.A4).toBeUndefined()
    expect(fillSheet.cells.A5.value).toBe('5')
    // sort refuses ranges touching locks
    expect(sortRange(base, 'B1:C4')).toBe(base)
    // replace skips locked cells
    const replaced = replaceAllInSheet(base, 'value', 'thing')
    expect(replaced.sheet.cells.B2.value).toBe('locked value')
    expect(replaced.sheet.cells.C2.value).toBe('free thing')
    expect(replaced.count).toBe(1)
  })

  it('unlock removes intersecting locked ranges', () => {
    const unlocked = unlockRange(base, 'B2:B2')
    expect(unlocked.protection).toBeUndefined()
    expect(isCellLocked(unlocked, 'B2')).toBe(false)
    // unlocking a non-intersecting range changes nothing
    expect(unlockRange(base, 'D1:D4')).toBe(base)
  })
})

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

describe('data validation', () => {
  it('list rules flag values outside the allowed set', () => {
    let sheet = sheetWith({ A1: 'north', A2: 'south', A3: 'NOPE' })
    sheet = addValidationRule(sheet, {
      range: 'A1:A4',
      kind: 'list',
      options: ['north', 'south'],
    })
    const rule = validationRuleForCell(sheet, 'A3')
    if (!rule) throw new Error('expected rule for A3')
    expect(validateValue(rule, 'north').valid).toBe(true)
    expect(validateValue(rule, 'NOPE').valid).toBe(false)
    expect(validateValue(rule, '').valid).toBe(true) // blanks always valid
    const flagged = invalidCellsForSheet(sheet)
    expect(Array.from(flagged.keys())).toEqual(['A3'])
    expect(flagged.get('A3')).toContain('not in the allowed list')
  })

  it('number-range rules flag non-numbers and out-of-range numbers', () => {
    let sheet = sheetWith({ B1: '5', B2: '50', B3: 'abc' })
    sheet = addValidationRule(sheet, {
      range: 'B1:B3',
      kind: 'number-range',
      min: 0,
      max: 10,
    })
    const flagged = invalidCellsForSheet(sheet)
    expect(flagged.has('B1')).toBe(false)
    expect(flagged.get('B2')).toBe('Value must be at most 10')
    expect(flagged.get('B3')).toBe('Value must be a number')
  })

  it('date rules flag unparseable dates', () => {
    let sheet = sheetWith({ C1: '2026-07-05', C2: 'not a date' })
    sheet = addValidationRule(sheet, { range: 'C1:C2', kind: 'date' })
    const flagged = invalidCellsForSheet(sheet)
    expect(flagged.has('C1')).toBe(false)
    expect(flagged.get('C2')).toBe('Value must be a date')
  })

  it('rules validate computed formula results, replace per range, and clear', () => {
    let sheet = sheetWith({ A1: '8', A2: '=A1*2' })
    sheet = addValidationRule(sheet, {
      range: 'A2:A2',
      kind: 'number-range',
      max: 10,
    })
    expect(invalidCellsForSheet(sheet).get('A2')).toBe(
      'Value must be at most 10',
    )
    // Re-adding on the same range replaces the earlier rule.
    sheet = addValidationRule(sheet, {
      range: 'A2:A2',
      kind: 'number-range',
      max: 100,
    })
    expect(sheet.validations).toHaveLength(1)
    expect(invalidCellsForSheet(sheet).size).toBe(0)
    sheet = clearValidationInRange(sheet, 'A1:A5')
    expect(sheet.validations).toBeUndefined()
  })

  it('invalidCellFlagsForSheet pairs each flagged cell with its computed value and message', () => {
    // The production surface paints from these flags: cell + current computed
    // value (formula results included) + message.
    let sheet = sheetWith({ A1: '5', A2: '=A1*10', A3: 'nope' })
    sheet = addValidationRule(sheet, {
      range: 'A1:A3',
      kind: 'number-range',
      min: 0,
      max: 10,
    })
    const flags = invalidCellFlagsForSheet(sheet)
    expect(flags).toEqual([
      { cell: 'A2', value: '50', message: 'Value must be at most 10' },
      { cell: 'A3', value: 'nope', message: 'Value must be a number' },
    ])
  })
})

// ---------------------------------------------------------------------------
// Conditional formatting
// ---------------------------------------------------------------------------

describe('conditional formatting', () => {
  it('evaluates all five condition kinds', () => {
    expect(conditionMatches({ kind: 'greater', value: '10' }, '11')).toBe(true)
    expect(conditionMatches({ kind: 'greater', value: '10' }, '10')).toBe(false)
    expect(conditionMatches({ kind: 'less', value: '10' }, '9.5')).toBe(true)
    expect(conditionMatches({ kind: 'equal', value: '5' }, '5.0')).toBe(true)
    expect(conditionMatches({ kind: 'equal', value: 'ok' }, 'OK')).toBe(true)
    expect(
      conditionMatches({ kind: 'between', value: '5', value2: '10' }, '7'),
    ).toBe(true)
    expect(
      conditionMatches({ kind: 'between', value: '10', value2: '5' }, '7'),
    ).toBe(true) // bounds normalize
    expect(
      conditionMatches({ kind: 'text-contains', value: 'err' }, 'NO ERRORS'),
    ).toBe(true)
    expect(conditionMatches({ kind: 'greater', value: '10' }, '')).toBe(false)
  })

  it('applies styles from matching rules to computed values, merging in order', () => {
    let sheet = sheetWith({ A1: '4', A2: '20', A3: '=A2*2' })
    sheet = addConditionalFormatRule(sheet, {
      range: 'A1:A3',
      condition: { kind: 'greater', value: '10' },
      style: { fillColor: '#fff4bf' },
    })
    sheet = addConditionalFormatRule(sheet, {
      range: 'A1:A3',
      condition: { kind: 'greater', value: '30' },
      style: { textColor: '#982d25', bold: true },
    })
    expect(conditionalStyleForCell(sheet, 'A1')).toBeUndefined()
    expect(conditionalStyleForCell(sheet, 'A2')).toEqual({
      fillColor: '#fff4bf',
    })
    // A3 = 40: both rules match and merge
    expect(conditionalStyleForCell(sheet, 'A3')).toEqual({
      fillColor: '#fff4bf',
      textColor: '#982d25',
      bold: true,
    })
    const cleared = clearConditionalFormatsInRange(sheet, 'A1:A3')
    expect(cleared.conditionalFormats).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// Filter query + find/replace
// ---------------------------------------------------------------------------

describe('filter query', () => {
  it('narrows visible rows to contains-matches on the filter column', () => {
    let sheet = sheetWith({
      A1: 'Region',
      A2: 'North',
      A3: 'South',
      A4: 'Northeast',
    })
    sheet = toggleFilterRow(sheet, 0)
    // default filter: non-empty rows only
    expect(visibleRowsForSheet(sheet).slice(0, 4)).toEqual([0, 1, 2, 3])
    sheet = setFilterQuery(sheet, 'north')
    expect(visibleRowsForSheet(sheet)).toEqual([0, 1, 3])
    sheet = setFilterQuery(sheet, '')
    expect(sheet.filterQuery).toBeUndefined()
  })
})

describe('find and replace', () => {
  const sheet = sheetWith({
    A1: 'North',
    A2: 'north star',
    B1: 'SOUTH',
    B2: '=A1',
  })

  it('findMatches is case-insensitive by default and ordered', () => {
    expect(findMatches(sheet, 'north')).toEqual(['A1', 'A2', 'B2'])
    expect(findMatches(sheet, 'north', { matchCase: true })).toEqual(['A2'])
  })

  it('replaceInRange honors matchCase', () => {
    const sensitive = replaceInRange(sheet, 'A1:A2', 'north', 'west', {
      matchCase: true,
    })
    expect(sensitive.count).toBe(1)
    expect(sensitive.sheet.cells.A1.value).toBe('North')
    expect(sensitive.sheet.cells.A2.value).toBe('west star')
    const insensitive = replaceInRange(sheet, 'A1:A2', 'north', 'west', {
      matchCase: false,
    })
    expect(insensitive.count).toBe(2)
    expect(insensitive.sheet.cells.A1.value).toBe('west')
  })

  it('replaceAllInSheet counts every occurrence across the sheet', () => {
    const result = replaceAllInSheet(sheet, 'south', 'east')
    expect(result.count).toBe(1)
    expect(result.sheet.cells.B1.value).toBe('east')
  })
})