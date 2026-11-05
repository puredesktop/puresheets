/**
 * An agent building a financial model froze the sheet: every write was followed by work that read every formula
 * (autosave's linked tables, the formula bar, the range summary, agent reads), and the evaluator re-walked each
 * reference's whole dependency chain on every read. Running totals (cash = last month's cash + this month's flow)
 * made that exponential: one cell of a 742-formula model took 16 s, and the renderer sat at 100% CPU for minutes.
 *
 * The workbook below is built like that model: 11 sheets, 36 months of chained rows, yearly sums of those rows,
 * a funding-need cell that reads the running total's MIN six times over, and cross-sheet references both ways.
 */
import { describe, expect, it } from 'vitest'
import {
  applyAgentCellChangesDirect,
  cellKey,
  computedCellValue,
  createBlankSheet,
  createBlankWorkbook,
  displayCellValue,
  upsertCell,
} from './workbookModel'
import { invalidCellFlagsForSheet } from './sheetDataFeatures'
import { workbookFilesForSave } from './workbookFiles'
import { workbookVersion } from './sheetsAgentContext'
import { withFreshUniverSnapshot } from './univerAdapter'
import type { PureSheetsDocument, WorkbookSheet } from '../types'

const MONTHS = 36
const FIRST = 56 // first month row, as in the model
const LAST = FIRST + MONTHS - 1

function sheetWith(id: string, name: string, cells: Record<string, string>): WorkbookSheet {
  let sheet: WorkbookSheet = { ...createBlankSheet(id, name), rowCount: 1000, columnCount: 26 }
  for (const [key, value] of Object.entries(cells)) sheet = upsertCell(sheet, key, value)
  return sheet
}

function financialModel(): PureSheetsDocument {
  const assumptions: Record<string, string> = {
    A1: 'Assumption', B1: 'Value',
    A6: 'Price per seat (annual)', B6: '42000',
    A8: 'Renewal rate', B8: '0.85',
    A9: 'COGS share', B9: '0.18',
    A10: 'Sales cost per new logo', B10: '6500',
    A11: 'Overhead multiplier', B11: '1.25',
    A12: 'Second-year renewal', B12: '0.7',
    A15: 'Salary', B15: '96000',
    A16: 'Tools per FTE', B16: '350',
    A17: 'Rent', B17: '4200',
    A18: 'Other', B18: '1800',
    A20: 'Other (year 3)', B20: '2600',
    A21: 'Buffer', B21: '0.15',
    A22: 'Raise rounding', B22: '50000',
    A24: 'Market growth', B24: '1.18',
  }
  const monthly: Record<string, string> = {
    A55: 'Month', C55: 'New logos', D55: 'Renewing', E55: 'Churned', F55: 'Active', G55: 'Billings', H55: 'Revenue',
    I55: 'COGS', J55: 'Gross profit', K55: 'FTE', L55: 'Payroll', M55: 'Opex', N55: 'Sales cost', O55: 'Total opex',
    P55: 'Operating result', Q55: 'Cash flow', R55: 'Cumulative cash',
    B33: '6', C33: '9', D33: '14',
  }
  for (let year = 0; year < 3; year++) for (let m = 0; m < 12; m++) monthly[cellKey(38 + m, 1 + year)] = String(2 + m + year * 3)
  for (let m = 0; m < 12; m++) monthly[cellKey(38 + m, 3)] = `=B${39 + m}*Assumptions!$B$24`
  for (let row = FIRST; row <= LAST; row++) {
    const year = Math.floor((row - FIRST) / 12)
    const p = row - 1
    monthly[`A${row}`] = String(row - FIRST + 1)
    monthly[`C${row}`] = `=${'BCD'[year]}${39 + ((row - FIRST) % 12)}`
    monthly[`D${row}`] = row - 12 >= FIRST ? `=C${row - 12}*Assumptions!$B$8` : '0'
    monthly[`E${row}`] = row - 24 >= FIRST ? `=C${row - 12}*(1-Assumptions!$B$8)+C${row - 24}*(Assumptions!$B$8-Assumptions!$B$12)` : row - 12 >= FIRST ? `=C${row - 12}*(1-Assumptions!$B$8)` : '0'
    monthly[`F${row}`] = row === FIRST ? `=C${row}-E${row}` : `=F${p}+C${row}-E${row}`
    monthly[`G${row}`] = `=(C${row}+D${row})*Assumptions!$B$6`
    monthly[`H${row}`] = `=F${row}*Assumptions!$B$6/12`
    monthly[`I${row}`] = `=H${row}*Assumptions!$B$9`
    monthly[`J${row}`] = `=H${row}-I${row}`
    monthly[`K${row}`] = `=$${'BCD'[year]}$33`
    monthly[`L${row}`] = `=K${row}*Assumptions!$B$15/12*Assumptions!$B$11`
    monthly[`M${row}`] = `=(K${row}*Assumptions!$B$16+Assumptions!$B$17+Assumptions!$B$${year === 2 ? 20 : 18})*Assumptions!$B$11`
    monthly[`N${row}`] = `=C${row}*Assumptions!$B$10`
    monthly[`O${row}`] = `=L${row}+M${row}+N${row}`
    monthly[`P${row}`] = `=J${row}-O${row}`
    monthly[`Q${row}`] = `=G${row}-I${row}-O${row}`
    monthly[`R${row}`] = row === FIRST ? `=Q${row}` : `=R${p}+Q${row}`
  }
  // Yearly totals of every money column, then a total across the years.
  'GHIJLMNOPQ'.split('').forEach((column, index) => {
    const row = 95 + index
    for (let year = 0; year < 3; year++) {
      const from = FIRST + year * 12
      monthly[cellKey(row - 1, 1 + year)] = `=SUM(${column}${from}:${column}${from + 11})`
    }
    monthly[`E${row}`] = `=SUM(B${row}:D${row})`
  })
  Object.assign(monthly, {
    B117: `=-MIN(R${FIRST}:R${LAST})`,
    B118: `=COUNTIF(R${FIRST}:R${LAST},"<0")`,
    B119: `=COUNTIF(Q${FIRST}:Q${LAST},"<0")`,
    B120: '=B117*(1+Assumptions!B21)',
    B121: '=IF(ROUND(B120/Assumptions!B22,0)*Assumptions!B22<B120,(ROUND(B120/Assumptions!B22,0)+1)*Assumptions!B22,ROUND(B120/Assumptions!B22,0)*Assumptions!B22)',
    B122: `=B121+R${FIRST + 11}`,
    B123: '=B122/B121',
    B124: `=IF(B121>E96,B121-E96,0)`,
  })
  const annual: Record<string, string> = { A1: 'Year', B1: 'Revenue', C1: 'Opex', D1: 'Cash', E1: 'Raise' }
  for (let year = 0; year < 3; year++) {
    const column = 'BCD'[year]
    annual[`A${year + 2}`] = String(year + 1)
    annual[`B${year + 2}`] = `='Monthly Model'!${column}95`
    annual[`C${year + 2}`] = `='Monthly Model'!${column}101`
    annual[`D${year + 2}`] = `='Monthly Model'!R${FIRST + year * 12 + 11}`
    annual[`E${year + 2}`] = `='Monthly Model'!B121*(${year}+1)/3`
  }
  // Cash after the raise, month by month, read across from the model.
  const runway: Record<string, string> = { A1: 'Month', B1: 'Cash before raise', C1: 'Cash after raise' }
  for (let month = 1; month <= MONTHS; month++) {
    runway[`A${month + 1}`] = String(month)
    runway[`B${month + 1}`] = `='Monthly Model'!R${FIRST + month - 1}`
    runway[`C${month + 1}`] = `=B${month + 1}+'Monthly Model'!$B$121`
  }
  const document = createBlankWorkbook('Generated financial model')
  const sheets = [
    sheetWith('assumptions', 'Assumptions', assumptions),
    sheetWith('market', 'Market Sizing', { A1: 'Segment', B1: '120000', B2: '=B1*Assumptions!B24', B3: "='Annual Summary'!B4/B2" }),
    sheetWith('monthly', 'Monthly Model', monthly),
    sheetWith('annual', 'Annual Summary', annual),
    sheetWith('runway', 'Cash Flow and Runway', runway),
    ...['Staffing and Opex', 'Scenarios', 'Sensitivity', 'Unit Economics', 'Funding and Roadmap', 'Sources and Limitations'].map(
      (name, index) => sheetWith(`notes-${index}`, name, { A1: name, A2: 'Notes the agent has not written yet.' }),
    ),
  ]
  return withFreshUniverSnapshot({
    ...document,
    workbook: { ...document.workbook, sheets, activeSheetId: 'monthly' },
  })
}

function formulaCount(document: PureSheetsDocument): number {
  return document.workbook.sheets.reduce(
    (total, sheet) => total + Object.values(sheet.cells).filter(cell => cell.kind === 'formula').length,
    0,
  )
}

/** What one agent write costs: the handler's version check and write, then what the shell recomputes after it. */
async function agentWriteAndRecompute(document: PureSheetsDocument, index: number): Promise<PureSheetsDocument> {
  await workbookVersion(document, '/Drafts/model.sheets')
  const row = FIRST + (index % MONTHS)
  const { document: next, applied } = applyAgentCellChangesDirect(
    document,
    'monthly',
    `C${FIRST}:C${LAST}`,
    `New-logo plan for month ${row - FIRST + 1}`,
    [{ cell: `C${row}`, value: String(3 + (index % 7)), kind: 'number' }],
    'Data analyst',
  )
  expect(applied).toBe(1)
  const monthly = next.workbook.sheets.find(sheet => sheet.id === 'monthly')!
  for (const sheet of next.workbook.sheets) invalidCellFlagsForSheet(sheet, next.workbook) // validation flags
  displayCellValue(monthly, 'B121', next.workbook) // the formula bar on the selected cell
  for (let row = FIRST; row <= LAST; row++) computedCellValue(monthly, `R${row}`, next.workbook) // range summary
  workbookFilesForSave(next, { boundPath: '/Drafts/model.sheets', isPackage: true }) // autosave: every cell
  return next
}

describe('agent writes on a large financial model', () => {
  it('is the size of the model that froze', () => {
    const document = financialModel()
    expect(document.workbook.sheets).toHaveLength(11)
    expect(formulaCount(document)).toBeGreaterThan(690)
  })

  it('computes the model correctly', () => {
    const document = financialModel()
    const monthly = document.workbook.sheets.find(sheet => sheet.id === 'monthly')!
    const annual = document.workbook.sheets.find(sheet => sheet.id === 'annual')!
    // Month 1: 2 new logos, nothing renewing or churning.
    expect(computedCellValue(monthly, `F${FIRST}`, document.workbook)).toBe('2')
    expect(computedCellValue(monthly, `H${FIRST}`, document.workbook)).toBe('7000')
    // The funding need is the deepest cash dip plus the buffer, rounded up to the next 50,000.
    const dip = Number(computedCellValue(monthly, 'B117', document.workbook))
    const raise = Number(computedCellValue(monthly, 'B121', document.workbook))
    expect(dip).toBeGreaterThan(0)
    expect(raise % 50000).toBe(0)
    expect(raise).toBeGreaterThanOrEqual(dip * 1.15)
    expect(raise - dip * 1.15).toBeLessThan(50000)
    expect(computedCellValue(annual, 'E4', document.workbook)).toBe(String(raise))
  })

  it('takes a burst of 50 agent writes, each with the full recompute, within budget', async () => {
    const BUDGET_MS = 5000 // was: a single full read of the model did not finish in minutes
    let document = financialModel()
    const start = performance.now()
    for (let index = 0; index < 50; index++) document = await agentWriteAndRecompute(document, index)
    const elapsed = performance.now() - start
    expect(document.workbook.agentLog?.length).toBe(50)
    expect(elapsed).toBeLessThan(BUDGET_MS)
  }, 60_000)

  it('gives the same answers as a fresh read after each write', async () => {
    let document = financialModel()
    for (let index = 0; index < 5; index++) {
      document = await agentWriteAndRecompute(document, index)
      const monthly = document.workbook.sheets.find(sheet => sheet.id === 'monthly')!
      // A copy of the state shares no cached results with it.
      const fresh = { sheets: document.workbook.sheets.map(sheet => ({ ...sheet })) }
      const freshMonthly = fresh.sheets.find(sheet => sheet.id === 'monthly')!
      for (const key of ['B117', 'B121', 'E96', `R${LAST}`])
        expect(computedCellValue(monthly, key, document.workbook)).toBe(computedCellValue(freshMonthly, key, fresh))
    }
  })
})
