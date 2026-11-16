import { expect, it } from 'vitest'
import { LocaleType, Univer, UniverInstanceType } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import '@univerjs/sheets/facade'
import '@univerjs/engine-formula/facade'
import '@univerjs/sheets-formula/facade'
import { createUniverSnapshot } from './univerAdapter'
import { cloneWithUpdatedSheet, createBlankWorkbook, upsertCell } from './workbookModel'

it('calculates a workbook-scoped named range in the live Univer engine', async () => {
  const base = createBlankWorkbook()
  const sheet = upsertCell(upsertCell(upsertCell(base.workbook.sheets[0], 'A1', '10'), 'A2', '20'), 'B1', '=SUM(Revenue)')
  const document = cloneWithUpdatedSheet({ ...base, workbook: { ...base.workbook, namedRanges: [{ name: 'Revenue', sheetId: sheet.id, range: 'A1:A2' }] } }, sheet)
  const univer = new Univer({ locale: LocaleType.EN_US })
  try {
    univer.registerPlugin(UniverFormulaEnginePlugin)
    univer.registerPlugin(UniverSheetsPlugin)
    univer.registerPlugin(UniverSheetsFormulaPlugin)
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, createUniverSnapshot(document))
    const api = FUniver.newAPI(univer)
    const book = api.getActiveWorkbook()!
    expect(book.getDefinedNames().map(name => name.getName())).toContain('Revenue')
    const formula = api.getFormula()
    const settled = formula.onCalculationResultApplied(3000)
    formula.executeCalculation()
    await settled
    expect(book.getActiveSheet().getRange('B1').getValue()).toBe(30)
  } finally { univer.dispose() }
})
