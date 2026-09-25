/**
 * Phase S2 — engine-agnostic golden formula suite.
 *
 * One table of golden cases runs against BOTH engines that evaluate formulas
 * in PureSheets:
 *  - the legacy shim (`evaluateFormula` in workbookModel.ts) used by the
 *    compatibility grid / QA / chart paths, and
 *  - Univer's real formula engine (headless: engine-formula + sheets +
 *    sheets-formula plugins), which powers live editing.
 *
 * Every case shares one fixture grid; each case's formula is placed in its
 * own scratch cell so the Univer run evaluates the whole table in a single
 * headless workbook. Where the two engines legitimately differ in
 * representation (dates, error spelling), a case provides a per-engine
 * expectation.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { LocaleType, Univer, UniverInstanceType } from '@univerjs/core'
import type { ICellData } from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { FUniver } from '@univerjs/core/facade'
import '@univerjs/sheets/facade'
import '@univerjs/engine-formula/facade'
import '@univerjs/sheets-formula/facade'
import { createBlankWorkbook, evaluateFormula, upsertCell } from './workbookModel'
import type { WorkbookSheet } from '../types'

// ---------------------------------------------------------------------------
// Shared fixture grid (columns A–E). Blanks are deliberate: A4 is empty inside
// A1:A5, and B4 is empty inside B1:B5.
// ---------------------------------------------------------------------------
const FIXTURE_CELLS: Record<string, string> = {
  A1: '10',
  A2: '20',
  A3: '30',
  A5: '40', // A4 intentionally blank
  B1: 'apples',
  B2: 'bananas',
  B3: 'apples',
  B5: 'cherries', // B4 intentionally blank
  C1: '5',
  C2: '15',
  C3: '25',
  D1: 'north',
  D2: 'south',
  D3: 'north',
  E1: '100',
  E2: '200',
  E3: '300',
}

type Expectation = string | RegExp | ((value: string) => boolean)

interface GoldenCase {
  name: string
  formula: string
  expected: Expectation
  /** Override when Univer's representation legitimately differs. */
  univerExpected?: Expectation
}

const GOLDEN_CASES: GoldenCase[] = [
  // --- the 80/20 set --------------------------------------------------------
  { name: 'SUM over a range', formula: '=SUM(A1:A3)', expected: '60' },
  { name: 'SUM skips blank cells', formula: '=SUM(A1:A5)', expected: '100' },
  { name: 'AVERAGE over a range', formula: '=AVERAGE(A1:A3)', expected: '20' },
  {
    name: 'AVERAGE ignores blanks (no phantom zeros)',
    formula: '=AVERAGE(A1:A5)',
    expected: '25',
  },
  { name: 'MIN', formula: '=MIN(A1:A3)', expected: '10' },
  {
    name: 'MIN ignores blanks (blank is not zero)',
    formula: '=MIN(A1:A5)',
    expected: '10',
  },
  { name: 'MAX', formula: '=MAX(A1:A5)', expected: '40' },
  {
    name: 'COUNT counts numbers only',
    formula: '=COUNT(A1:B5)',
    expected: '4',
  },
  {
    name: 'COUNTA counts non-empty cells',
    formula: '=COUNTA(B1:B5)', // B4 blank → 4 of 5 filled
    expected: '4',
  },
  {
    name: 'COUNTIF with text criterion',
    formula: '=COUNTIF(B1:B5,"apples")',
    expected: '2',
  },
  {
    name: 'COUNTIF with comparison criterion',
    formula: '=COUNTIF(A1:A5,">15")',
    expected: '3',
  },
  {
    name: 'COUNTIF never matches blanks numerically',
    formula: '=COUNTIF(A1:A5,">=0")',
    expected: '4',
  },
  {
    name: 'SUMIF with criteria range and sum range',
    formula: '=SUMIF(D1:D3,"north",E1:E3)',
    expected: '400',
  },
  {
    name: 'AVERAGEIF with criteria range and average range',
    formula: '=AVERAGEIF(D1:D3,"north",E1:E3)',
    expected: '200',
  },
  { name: 'IF true branch', formula: '=IF(A1>5,"big","small")', expected: 'big' },
  { name: 'IF false branch', formula: '=IF(A1>50,"big","small")', expected: 'small' },
  { name: 'ROUND to 2 places', formula: '=ROUND(2.567,2)', expected: '2.57' },
  { name: 'ROUND to integer', formula: '=ROUND(2.4,0)', expected: '2' },
  {
    name: 'CONCAT joins values',
    formula: '=CONCAT(B1,D2)',
    expected: 'applessouth',
  },
  { name: 'LEFT', formula: '=LEFT(B2,3)', expected: 'ban' },
  { name: 'RIGHT', formula: '=RIGHT(B2,3)', expected: 'nas' },
  {
    name: 'VLOOKUP exact match',
    formula: '=VLOOKUP("south",D1:E3,2,FALSE)',
    expected: '200',
  },
  {
    name: 'TODAY produces a date',
    formula: '=TODAY()',
    // Legacy: ISO date string. Univer: a positive date serial number.
    expected: value => /^\d{4}-\d{2}-\d{2}$/.test(value),
    univerExpected: value => Number(value) > 40000,
  },
  {
    name: 'NOW produces a date-time',
    formula: '=NOW()',
    expected: value => /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(value),
    univerExpected: value => Number(value) > 40000,
  },
  // --- operators, references, ranges ---------------------------------------
  { name: 'arithmetic with precedence', formula: '=A1+A2*2', expected: '50' },
  { name: 'parentheses', formula: '=(A1+A2)*2', expected: '60' },
  { name: 'unary minus', formula: '=-A1+5', expected: '-5' },
  { name: 'text concatenation operator', formula: '=B1&"-"&D1', expected: 'apples-north' },
  {
    name: 'absolute references evaluate like relative ones',
    formula: '=SUM($A$1:$A$3)',
    expected: '60',
  },
  {
    name: 'mixed absolute/relative reference',
    formula: '=$A$1+C1',
    expected: '15',
  },
  {
    name: 'range reference inside nested function',
    formula: '=ROUND(AVERAGE(C1:C3),0)',
    expected: '15',
  },
  // --- error surfacing -------------------------------------------------------
  {
    name: 'division by zero',
    formula: '=A1/0',
    expected: '#DIV/0!',
  },
  {
    name: 'unknown function name',
    formula: '=NOTAREALFN(A1)',
    expected: /^#NAME\??/,
    univerExpected: /^#NAME\??/,
  },
  {
    name: 'VLOOKUP miss',
    formula: '=VLOOKUP("missing",D1:E3,2,FALSE)',
    expected: '#N/A',
  },
  {
    name: 'AVERAGE of empty cells errors',
    formula: '=AVERAGE(A4:A4)',
    expected: '#DIV/0!',
  },
]

// Scratch cells hold one formula per case, in column H (clear of the fixture).
const SCRATCH_COLUMN = 7 // H
function scratchKey(index: number): string {
  return `H${index + 1}`
}

function matches(value: string, expectation: Expectation): boolean {
  if (typeof expectation === 'function') return expectation(value)
  if (expectation instanceof RegExp) return expectation.test(value)
  // Numeric-aware comparison: '2.57' vs 2.57 vs '2.5700000000000003'.
  const left = Number(value)
  const right = Number(expectation)
  if (
    value.trim() !== '' &&
    expectation.trim() !== '' &&
    Number.isFinite(left) &&
    Number.isFinite(right)
  ) {
    return Math.abs(left - right) < 1e-9
  }
  return value === expectation
}

function describeExpectation(expectation: Expectation): string {
  if (typeof expectation === 'function') return '<predicate>'
  return String(expectation)
}

// ---------------------------------------------------------------------------
// Engine 1 — legacy shim (workbookModel.evaluateFormula)
// ---------------------------------------------------------------------------
function legacySheet(): WorkbookSheet {
  let sheet = createBlankWorkbook('Golden').workbook.sheets[0]
  for (const [key, value] of Object.entries(FIXTURE_CELLS)) {
    sheet = upsertCell(sheet, key, value)
  }
  return sheet
}

describe('golden formulas — legacy shim engine', () => {
  const sheet = legacySheet()
  for (const goldenCase of GOLDEN_CASES) {
    it(goldenCase.name, () => {
      const value = evaluateFormula(sheet, goldenCase.formula)
      expect(
        matches(value, goldenCase.expected),
        `${goldenCase.formula} → ${value} (expected ${describeExpectation(
          goldenCase.expected,
        )})`,
      ).toBe(true)
    })
  }

  it('circular references surface an error instead of hanging', () => {
    let sheet = createBlankWorkbook('Cycle').workbook.sheets[0]
    sheet = upsertCell(sheet, 'A1', '=B1')
    sheet = upsertCell(sheet, 'B1', '=A1')
    expect(evaluateFormula(sheet, '=A1')).toMatch(/^#/)
  })
})

// ---------------------------------------------------------------------------
// Engine 2 — Univer formula engine (headless)
// ---------------------------------------------------------------------------
type UniverFacadeApi = {
  getFormula?: () => {
    executeCalculation?: () => void
    onCalculationResultApplied?: (timeout?: number) => Promise<void>
  }
  getActiveWorkbook?: () => {
    getActiveSheet?: () => {
      getRange?: (token: string) => {
        getValue?: () => unknown
        getDisplayValue?: () => string
      }
    } | null
  } | null
}

function univerCellData(): Record<number, Record<number, ICellData>> {
  const cellData: Record<number, Record<number, ICellData>> = {}
  const put = (row: number, column: number, cell: ICellData): void => {
    cellData[row] = { ...cellData[row], [column]: cell }
  }
  for (const [key, value] of Object.entries(FIXTURE_CELLS)) {
    const match = /^([A-Z]+)(\d+)$/.exec(key)
    if (!match) continue
    const column = match[1].charCodeAt(0) - 65
    const row = Number(match[2]) - 1
    const numeric = Number(value)
    put(
      row,
      column,
      value !== '' && Number.isFinite(numeric)
        ? { v: numeric, t: 2 }
        : { v: value, t: 1 },
    )
  }
  GOLDEN_CASES.forEach((goldenCase, index) => {
    put(index, SCRATCH_COLUMN, { f: goldenCase.formula })
  })
  return cellData
}

describe('golden formulas — Univer engine (headless)', () => {
  let univer: Univer | null = null

  const resultsPromise = (async () => {
    univer = new Univer({ locale: LocaleType.EN_US })
    univer.registerPlugin(UniverFormulaEnginePlugin)
    univer.registerPlugin(UniverSheetsPlugin)
    univer.registerPlugin(UniverSheetsFormulaPlugin)
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'golden',
      name: 'golden',
      appVersion: '0.25.0',
      locale: LocaleType.EN_US,
      styles: {},
      sheetOrder: ['g1'],
      sheets: {
        g1: {
          id: 'g1',
          name: 'Golden',
          rowCount: Math.max(GOLDEN_CASES.length + 2, 40),
          columnCount: 12,
          cellData: univerCellData(),
        },
      },
    })
    const api = FUniver.newAPI(univer) as unknown as UniverFacadeApi
    api.getFormula?.()?.executeCalculation?.()
    await api
      .getFormula?.()
      ?.onCalculationResultApplied?.(10000)
      ?.catch(() => undefined)
    const worksheet = api.getActiveWorkbook?.()?.getActiveSheet?.()
    const values = new Map<number, string>()
    GOLDEN_CASES.forEach((_goldenCase, index) => {
      const range = worksheet?.getRange?.(scratchKey(index))
      const raw = range?.getValue?.()
      const display = range?.getDisplayValue?.()
      // Prefer the raw value (stable numerics); fall back to display (errors).
      const value =
        raw === null || raw === undefined || raw === ''
          ? display ?? ''
          : String(raw)
      values.set(index, value)
    })
    return values
  })()

  afterAll(() => {
    univer?.dispose()
    univer = null
  })

  GOLDEN_CASES.forEach((goldenCase, index) => {
    it(goldenCase.name, async () => {
      const values = await resultsPromise
      const value = values.get(index) ?? ''
      const expectation = goldenCase.univerExpected ?? goldenCase.expected
      expect(
        matches(value, expectation),
        `${goldenCase.formula} → ${value} (expected ${describeExpectation(
          expectation,
        )})`,
      ).toBe(true)
    })
  })
})
