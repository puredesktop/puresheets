import { describe, expect, it } from 'vitest'
import { applyRows, auditFormula, cleanedRows, goalSeek, pivotRows, readSelection } from './workbookAnalysis'
import { addSheet, cloneWithUpdatedSheet, createBlankWorkbook, lockRange, upsertCell } from './workbookModel'
import { loadSheetsDocument, serializeSheetsDocument } from './sheetsDocument'

function workbookWith(values: Record<string, string>) {
  const doc = createBlankWorkbook('Analysis')
  const sheet = Object.entries(values).reduce((sheet, [key, value]) => upsertCell(sheet, key, value), doc.workbook.sheets[0])
  return { doc: cloneWithUpdatedSheet(doc, sheet), sheet }
}

it('previews duplicates by selected column and splits quoted text without changing its digits', () => {
  const { doc, sheet } = workbookWith({
    A1: 'Ticket', B1: 'Notes', A2: '101', B2: '"A,B",004', A3: '101', B3: 'C,D',
  })
  const all = readSelection(sheet, 'A1:B3', doc.workbook)
  expect(cleanedRows(all, { mode: 'dedupe', columns: [0], separator: '', header: true })).toHaveLength(2)
  const column = readSelection(sheet, 'B1:B3', doc.workbook)
  const split = cleanedRows(column, { mode: 'split', columns: [], separator: ',', header: true })
  expect(split[1]).toEqual([{ value: 'A,B', kind: 'text' }, { value: '004', kind: 'text' }])
  expect(sheet.cells.B2.value).toBe('"A,B",004')
})

it('groups and updates a saved pivot without replacing an adjacent occupied cell', () => {
  const { doc, sheet } = workbookWith({ A1: 'Region', B1: 'Revenue', A2: 'South', B2: '50', A3: 'South', B3: '200', A4: 'North', B4: '100' })
  const rows = pivotRows(readSelection(sheet, 'A1:B4', doc.workbook), { row: 0, column: null, value: 1, summary: 'sum' })
  expect(rows[1][1].value).toBe('250')
  const added = addSheet(doc, 'Pivot')
  const output = applyRows(added.workbook.sheets.at(-1)!, rows)
  const saved = cloneWithUpdatedSheet(added, { ...output, pivot: { sourceSheetId: sheet.id, range: 'A1:B4', row: 0, column: null, value: 1, summary: 'sum', output: 'A1', rows: 3, columns: 2 } })
  const loaded = loadSheetsDocument(serializeSheetsDocument(saved)).workbook.sheets[1]
  expect(loaded.pivot?.sourceSheetId).toBe(sheet.id)
  expect(loaded.cells.B2.value).toBe('250')
  expect(() => applyRows(upsertCell(output, 'A4', 'Other data'), [...rows, rows[1]], 'A1', loaded.pivot)).toThrow('already contains data')
})

describe('auditing and Goal Seek', () => {
  it('traces named ranges across sheets and ignores references in string literals', () => {
    const { doc, sheet } = workbookWith({ A1: '10' })
    const next = addSheet(doc, 'Forecast')
    const forecast = upsertCell(next.workbook.sheets[1], 'B2', '=IF(TRUE,TaxRate,"A8")')
    const workbook = { ...next.workbook, sheets: [sheet, forecast], namedRanges: [{ name: 'TaxRate', range: 'A1', sheetId: sheet.id }] }
    expect(auditFormula(workbook, forecast, 'B2').precedents).toEqual([{ sheetId: sheet.id, range: 'A1' }])
    expect(auditFormula(workbook, sheet, 'A1').dependents).toEqual([{ sheetId: forecast.id, range: 'B2' }])
  })
  it('previews the input for a target and leaves the original intact', () => {
    const { doc, sheet } = workbookWith({ A1: '10', B1: '=A1*4-9' })
    expect(goalSeek(doc.workbook, sheet, 'B1', 'A1', 51).value).toBeCloseTo(15, 7)
    expect(sheet.cells.A1.value).toBe('10')
    expect(() => goalSeek(doc.workbook, lockRange(sheet, 'A1:A1'), 'B1', 'A1', 51)).toThrow('unlocked')
  })
})
