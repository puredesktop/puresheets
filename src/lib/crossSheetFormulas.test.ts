/** Formulas that read other sheets (=Costs!B4, ='Q1 plan'!A1:A3), as the agent, exports and the fallback grid see them. */
import { describe, expect, it } from 'vitest'
import { computedCellValue, createBlankSheet, displayCellValue, evaluateFormula, upsertCell } from './workbookModel'
import type { WorkbookSheet } from '../types'

function named(id: string, name: string, cells: Record<string, string>): WorkbookSheet {
  let sheet = createBlankSheet(id, name)
  for (const [key, value] of Object.entries(cells)) sheet = upsertCell(sheet, key, value)
  return sheet
}

const costs = named('s-costs', 'Costs', { A1: 'Rent', B1: '1200', A2: 'Staff', B2: '5300', B3: '=SUM(B1:B2)' })
const plan = named('s-plan', 'Q1 plan', { A1: '10', A2: '20', A3: '30', B1: 'x', B2: 'y', B3: 'z' })
const summary = named('s-sum', 'Summary', {
  A1: '=Costs!B3',
  A2: "='Q1 plan'!A2*2",
  A3: "=SUM('Q1 plan'!A1:A3)",
  A4: '=Costs!B1+Costs!B2',
  A5: '=Missing!A1',
  A6: "=VLOOKUP(20,'Q1 plan'!A1:B3,2)",
  A7: "=SUMIF('Q1 plan'!A1:A3,\">15\",'Q1 plan'!A1:A3)",
  A8: '=A1*2',
  A9: '=costs!$B$1',
})
const workbook = { sheets: [costs, plan, summary] }

describe('formulas that read other sheets', () => {
  it('reads a cell on another sheet, including one that is itself a formula', () => {
    expect(computedCellValue(summary, 'A1', workbook)).toBe('6500')
    expect(computedCellValue(summary, 'A4', workbook)).toBe('6500')
    expect(computedCellValue(summary, 'A8', workbook)).toBe('13000')
  })

  it("reads quoted sheet names with spaces, ranges in functions, and lookups", () => {
    expect(computedCellValue(summary, 'A2', workbook)).toBe('40')
    expect(computedCellValue(summary, 'A3', workbook)).toBe('60')
    expect(computedCellValue(summary, 'A6', workbook)).toBe('y')
    expect(computedCellValue(summary, 'A7', workbook)).toBe('50')
  })

  it('matches sheet names without regard to case and ignores $', () => {
    expect(computedCellValue(summary, 'A9', workbook)).toBe('1200')
  })

  it('a sheet that does not exist is #REF!, as in Excel, not #FORMULA', () => {
    expect(computedCellValue(summary, 'A5', workbook)).toBe('#REF!')
    expect(displayCellValue(summary, 'A5', workbook)).toBe('#REF!')
  })

  it('without the workbook, another sheet is #REF! and the sheet’s own name still works', () => {
    expect(computedCellValue(summary, 'A1')).toBe('#REF!')
    expect(evaluateFormula(costs, '=Costs!B1*2')).toBe('2400')
  })

  it('the same cell on two sheets is not taken for a cycle', () => {
    const left = named('s-l', 'Left', { A1: '=Right!A1+1' })
    const right = named('s-r', 'Right', { A1: '5' })
    expect(computedCellValue(left, 'A1', { sheets: [left, right] })).toBe('6')
  })

  it('a real cycle across sheets is #REF!', () => {
    const left = named('s-l', 'Left', { A1: '=Right!A1' })
    const right = named('s-r', 'Right', { A1: '=Left!A1' })
    expect(computedCellValue(left, 'A1', { sheets: [left, right] })).toBe('#REF!')
  })
})
