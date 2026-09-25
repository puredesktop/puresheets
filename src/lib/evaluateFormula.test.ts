/**
 * Tests for evaluateFormula fixes:
 * 1. Arithmetic expressions (G2/H2, C2+D2+E2+F2, G2*0.1)
 * 2. COUNTIF / SUMIF / AVERAGEIF / RANK
 * 3. Cascaded IF formulas that depend on arithmetic (the "Bronze bug")
 */
import { describe, expect, it } from 'vitest'
import {
  createBlankSheet,
  evaluateFormula,
  upsertCell,
} from './workbookModel'
import type { WorkbookSheet } from '../types'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function sheet(cells: Record<string, { value: string; kind: 'text' | 'number' | 'formula' }>): WorkbookSheet {
  let s = createBlankSheet('test-id', 'Test')
  for (const [key, cell] of Object.entries(cells)) {
    s = upsertCell(s, key, cell.value)
  }
  return s
}

// ---------------------------------------------------------------------------
// Fix 1: Arithmetic expressions
// ---------------------------------------------------------------------------

describe('arithmetic expressions in evaluateFormula', () => {
  // Dataset A: sales ratios
  it('Dataset A — division: G2/H2 (355000/340000)', () => {
    const s = sheet({
      G2: { value: '355000', kind: 'number' },
      H2: { value: '340000', kind: 'number' },
    })
    const result = evaluateFormula(s, '=G2/H2')
    expect(Number(result)).toBeCloseTo(355000 / 340000)
  })

  it('Dataset A — addition chain: C2+D2+E2+F2 (82000+91000+87000+95000)', () => {
    const s = sheet({
      C2: { value: '82000', kind: 'number' },
      D2: { value: '91000', kind: 'number' },
      E2: { value: '87000', kind: 'number' },
      F2: { value: '95000', kind: 'number' },
    })
    expect(evaluateFormula(s, '=C2+D2+E2+F2')).toBe('355000')
  })

  it('Dataset A — multiply by constant: G2*0.1', () => {
    const s = sheet({ G2: { value: '355000', kind: 'number' } })
    expect(evaluateFormula(s, '=G2*0.1')).toBe('35500')
  })

  // Dataset B: temperature conversion
  it('Dataset B — Celsius to Fahrenheit: B3*1.8+32 (25°C → 77°F)', () => {
    const s = sheet({ B3: { value: '25', kind: 'number' } })
    expect(evaluateFormula(s, '=B3*1.8+32')).toBe('77')
  })

  it('Dataset B — profit margin: (B5-C5)/B5 (revenue 1000, cost 650)', () => {
    const s = sheet({
      B5: { value: '1000', kind: 'number' },
      C5: { value: '650', kind: 'number' },
    })
    // (1000-650)/1000 — parenthesized arithmetic evaluates through the
    // expression tokenizer.
    expect(evaluateFormula(s, '=(B5-C5)/B5')).toBe('0.35')
    expect(evaluateFormula(s, '=B5-C5')).toBe('350')
  })

  it('Dataset B — simple subtraction: D4-E4 (500-200)', () => {
    const s = sheet({
      D4: { value: '500', kind: 'number' },
      E4: { value: '200', kind: 'number' },
    })
    expect(evaluateFormula(s, '=D4-E4')).toBe('300')
  })

  // Dataset C: absolute refs ($G$2:$G$11 stripped to G2)
  it('Dataset C — absolute ref stripped: $G$2/$H$2', () => {
    const s = sheet({
      G2: { value: '120000', kind: 'number' },
      H2: { value: '100000', kind: 'number' },
    })
    const result = evaluateFormula(s, '=$G$2/$H$2')
    expect(Number(result)).toBeCloseTo(1.2)
  })

  it('Dataset C — multiply with mixed absolute: $G$3*0.15', () => {
    const s = sheet({ G3: { value: '200000', kind: 'number' } })
    expect(evaluateFormula(s, '=$G$3*0.15')).toBe('30000')
  })

  it('Dataset C — division by zero returns #DIV/0!', () => {
    const s = sheet({
      A1: { value: '100', kind: 'number' },
      B1: { value: '0', kind: 'number' },
    })
    expect(evaluateFormula(s, '=A1/B1')).toBe('#DIV/0!')
  })
})

// ---------------------------------------------------------------------------
// Fix 1b: Cascaded IF — the "Bronze bug"
// Dataset: IF(J>=1.2, Platinum, IF(J>=1, Gold, IF(J>=0.8, Silver, Bronze)))
// where J is itself an arithmetic formula
// ---------------------------------------------------------------------------

describe('IF formulas cascading through arithmetic cell refs (Bronze bug)', () => {
  function makeEmployeeRow(
    s: WorkbookSheet,
    row: number,
    q1: number, q2: number, q3: number, q4: number, target: number,
  ): WorkbookSheet {
    const r = String(row)
    s = upsertCell(s, `C${r}`, String(q1))
    s = upsertCell(s, `D${r}`, String(q2))
    s = upsertCell(s, `E${r}`, String(q3))
    s = upsertCell(s, `F${r}`, String(q4))
    s = upsertCell(s, `H${r}`, String(target))
    s = upsertCell(s, `G${r}`, `=C${r}+D${r}+E${r}+F${r}`)
    s = upsertCell(s, `I${r}`, `=G${r}/4`)
    s = upsertCell(s, `J${r}`, `=G${r}/H${r}`)
    s = upsertCell(s, `K${r}`, `=IF(J${r}>=1.2,"Platinum",IF(J${r}>=1,"Gold",IF(J${r}>=0.8,"Silver","Bronze")))`)
    s = upsertCell(s, `L${r}`, `=IF(J${r}>=1.2,G${r}*0.15,IF(J${r}>=1,G${r}*0.1,IF(J${r}>=0.8,G${r}*0.05,0)))`)
    return s
  }

  // Dataset A: Platinum employee (attainment = 1.44, quarterly 120+130+125+140=515, target 357)
  it('Dataset A — Platinum: attainment 1.44 → Tier=Platinum, Bonus=15%', () => {
    let s = createBlankSheet('id', 'Test')
    s = makeEmployeeRow(s, 2, 120000, 130000, 125000, 140000, 357000)
    expect(evaluateFormula(s, '=G2')).toBe('515000')
    expect(evaluateFormula(s, '=J2')).toContain('1.44') // ~1.4426
    expect(evaluateFormula(s, '=K2')).toBe('Platinum')
    expect(Number(evaluateFormula(s, '=L2'))).toBeCloseTo(515000 * 0.15)
  })

  // Dataset B: Gold employee (attainment = 1.044, 82+91+87+95=355, target 340)
  it('Dataset B — Gold: attainment 1.044 → Tier=Gold, Bonus=10%', () => {
    let s = createBlankSheet('id', 'Test')
    s = makeEmployeeRow(s, 2, 82000, 91000, 87000, 95000, 340000)
    expect(evaluateFormula(s, '=G2')).toBe('355000')
    expect(evaluateFormula(s, '=K2')).toBe('Gold')
    expect(Number(evaluateFormula(s, '=L2'))).toBeCloseTo(355000 * 0.1)
  })

  // Dataset C: Silver employee (attainment ~0.983, 70+75+72+78=295, target 300)
  it('Dataset C — Silver: attainment 0.983 → Tier=Silver, Bonus=5%', () => {
    let s = createBlankSheet('id', 'Test')
    s = makeEmployeeRow(s, 2, 70000, 75000, 72000, 78000, 300000)
    expect(evaluateFormula(s, '=G2')).toBe('295000')
    expect(evaluateFormula(s, '=K2')).toBe('Silver')
    expect(Number(evaluateFormula(s, '=L2'))).toBeCloseTo(295000 * 0.05)
  })

  // Dataset D: Bronze employee (attainment ~0.714, 50+55+60+65=230, target 322)
  it('Dataset D — Bronze: attainment 0.714 → Tier=Bronze, Bonus=0', () => {
    let s = createBlankSheet('id', 'Test')
    s = makeEmployeeRow(s, 2, 50000, 55000, 60000, 65000, 322000)
    expect(evaluateFormula(s, '=G2')).toBe('230000')
    expect(evaluateFormula(s, '=K2')).toBe('Bronze')
    expect(evaluateFormula(s, '=L2')).toBe('0')
  })

  // Dataset E: Avg Quarter formula on different row
  it('Dataset E — Avg Quarter: I5=G5/4 where G5 is itself a SUM formula', () => {
    let s = createBlankSheet('id', 'Test')
    s = makeEmployeeRow(s, 5, 90000, 88000, 92000, 94000, 360000)
    expect(evaluateFormula(s, '=G5')).toBe('364000')
    expect(evaluateFormula(s, '=I5')).toBe('91000')
  })
})

// ---------------------------------------------------------------------------
// Fix 2: COUNTIF / SUMIF / AVERAGEIF / RANK
// ---------------------------------------------------------------------------

describe('COUNTIF', () => {
  // Dataset A: department abbreviations
  it('Dataset A — counts ENG in B2:B5 (3 out of 4)', () => {
    const s = sheet({
      B2: { value: 'ENG', kind: 'text' },
      B3: { value: 'SLS', kind: 'text' },
      B4: { value: 'ENG', kind: 'text' },
      B5: { value: 'ENG', kind: 'text' },
    })
    expect(evaluateFormula(s, '=COUNTIF(B2:B5,"ENG")')).toBe('3')
  })

  // Dataset B: numeric threshold
  it('Dataset B — counts values >100 in A1:A4', () => {
    const s = sheet({
      A1: { value: '50', kind: 'number' },
      A2: { value: '150', kind: 'number' },
      A3: { value: '200', kind: 'number' },
      A4: { value: '80', kind: 'number' },
    })
    expect(evaluateFormula(s, '=COUNTIF(A1:A4,">100")')).toBe('2')
  })

  // Dataset C: case-insensitive text match
  it('Dataset C — case-insensitive: counts "mkt" matching "MKT"', () => {
    const s = sheet({
      C1: { value: 'MKT', kind: 'text' },
      C2: { value: 'SLS', kind: 'text' },
      C3: { value: 'MKT', kind: 'text' },
    })
    expect(evaluateFormula(s, '=COUNTIF(C1:C3,"mkt")')).toBe('2')
  })
})

describe('SUMIF', () => {
  // Dataset A: sum bonuses by department
  it('Dataset A — sum column L where column B = ENG', () => {
    const s = sheet({
      B2: { value: 'ENG', kind: 'text' },
      B3: { value: 'SLS', kind: 'text' },
      B4: { value: 'ENG', kind: 'text' },
      L2: { value: '35500', kind: 'number' },
      L3: { value: '51500', kind: 'number' },
      L4: { value: '36400', kind: 'number' },
    })
    expect(evaluateFormula(s, '=SUMIF(B2:B4,"ENG",L2:L4)')).toBe('71900')
  })

  // Dataset B: sum where numeric range > threshold
  it('Dataset B — sum values > 100 in A1:A4', () => {
    const s = sheet({
      A1: { value: '50', kind: 'number' },
      A2: { value: '150', kind: 'number' },
      A3: { value: '200', kind: 'number' },
      A4: { value: '80', kind: 'number' },
    })
    expect(evaluateFormula(s, '=SUMIF(A1:A4,">100")')).toBe('350')
  })

  // Dataset C: three departments, verify each
  it('Dataset C — SLS total bonus 153300 from three SLS rows', () => {
    const s = sheet({
      B2: { value: 'SLS', kind: 'text' },
      B3: { value: 'ENG', kind: 'text' },
      B4: { value: 'SLS', kind: 'text' },
      B5: { value: 'SLS', kind: 'text' },
      L2: { value: '51500', kind: 'number' },
      L3: { value: '36400', kind: 'number' },
      L4: { value: '46800', kind: 'number' },
      L5: { value: '55000', kind: 'number' },
    })
    expect(evaluateFormula(s, '=SUMIF(B2:B5,"SLS",L2:L5)')).toBe('153300')
  })
})

describe('AVERAGEIF', () => {
  // Dataset A: average attainment by department
  it('Dataset A — average J where B=ENG (three ENG rows)', () => {
    const s = sheet({
      B2: { value: 'ENG', kind: 'text' },
      B3: { value: 'SLS', kind: 'text' },
      B4: { value: 'ENG', kind: 'text' },
      B5: { value: 'ENG', kind: 'text' },
      J2: { value: '1.0441', kind: 'number' },
      J3: { value: '1.1444', kind: 'number' },
      J4: { value: '1.0111', kind: 'number' },
      J5: { value: '1.0540', kind: 'number' },
    })
    const result = Number(evaluateFormula(s, '=AVERAGEIF(B2:B5,"ENG",J2:J5)'))
    const expected = (1.0441 + 1.0111 + 1.0540) / 3
    expect(result).toBeCloseTo(expected, 3)
  })

  // Dataset B: no matches → #DIV/0!
  it('Dataset B — no matching rows returns #DIV/0!', () => {
    const s = sheet({
      B2: { value: 'ENG', kind: 'text' },
      J2: { value: '1.0', kind: 'number' },
    })
    expect(evaluateFormula(s, '=AVERAGEIF(B2:B2,"SLS",J2:J2)')).toBe('#DIV/0!')
  })

  // Dataset C: MKT average
  it('Dataset C — MKT average attainment (two MKT rows)', () => {
    const s = sheet({
      B2: { value: 'MKT', kind: 'text' },
      B3: { value: 'ENG', kind: 'text' },
      B4: { value: 'MKT', kind: 'text' },
      J2: { value: '0.9833', kind: 'number' },
      J3: { value: '1.0111', kind: 'number' },
      J4: { value: '1.0156', kind: 'number' },
    })
    const result = Number(evaluateFormula(s, '=AVERAGEIF(B2:B4,"MKT",J2:J4)'))
    expect(result).toBeCloseTo((0.9833 + 1.0156) / 2, 3)
  })
})

describe('RANK', () => {
  // Dataset A: rank employee by Annual Total descending
  it('Dataset A — highest value (550000) ranks 1st', () => {
    const s = sheet({
      G2: { value: '355000', kind: 'number' },
      G3: { value: '515000', kind: 'number' },
      G4: { value: '295000', kind: 'number' },
      G5: { value: '364000', kind: 'number' },
      G6: { value: '550000', kind: 'number' },
    })
    expect(evaluateFormula(s, '=RANK(G6,$G$2:$G$6,0)')).toBe('1')
  })

  // Dataset B: ascending rank
  it('Dataset B — lowest value (295000) ranks 1st ascending', () => {
    const s = sheet({
      G2: { value: '355000', kind: 'number' },
      G3: { value: '515000', kind: 'number' },
      G4: { value: '295000', kind: 'number' },
    })
    expect(evaluateFormula(s, '=RANK(G4,$G$2:$G$4,1)')).toBe('1')
  })

  // Dataset C: middle rank descending
  it('Dataset C — 515000 is 2nd out of 5 descending', () => {
    const s = sheet({
      G2: { value: '355000', kind: 'number' },
      G3: { value: '515000', kind: 'number' },
      G4: { value: '295000', kind: 'number' },
      G5: { value: '364000', kind: 'number' },
      G6: { value: '550000', kind: 'number' },
    })
    expect(evaluateFormula(s, '=RANK(G3,$G$2:$G$6,0)')).toBe('2')
  })
})

// ---------------------------------------------------------------------------
// Integration: full employee row with all formula columns
// ---------------------------------------------------------------------------

describe('full employee row integration (all formula columns evaluated)', () => {
  function buildSheet(rows: Array<{ q1: number; q2: number; q3: number; q4: number; target: number }>): WorkbookSheet {
    let s = createBlankSheet('id', 'Sheet1')
    rows.forEach(({ q1, q2, q3, q4, target }, i) => {
      const r = i + 2
      s = upsertCell(s, `C${r}`, String(q1))
      s = upsertCell(s, `D${r}`, String(q2))
      s = upsertCell(s, `E${r}`, String(q3))
      s = upsertCell(s, `F${r}`, String(q4))
      s = upsertCell(s, `H${r}`, String(target))
      s = upsertCell(s, `G${r}`, `=C${r}+D${r}+E${r}+F${r}`)
      s = upsertCell(s, `I${r}`, `=G${r}/4`)
      s = upsertCell(s, `J${r}`, `=G${r}/$H${r}`)
      s = upsertCell(s, `K${r}`, `=IF(J${r}>=1.2,"Platinum",IF(J${r}>=1,"Gold",IF(J${r}>=0.8,"Silver","Bronze")))`)
      s = upsertCell(s, `L${r}`, `=IF(J${r}>=1.2,G${r}*0.15,IF(J${r}>=1,G${r}*0.1,IF(J${r}>=0.8,G${r}*0.05,0)))`)
    })
    return s
  }

  it('Dataset A — 5-row mixed tier table evaluates all tiers correctly', () => {
    const s = buildSheet([
      { q1: 120000, q2: 130000, q3: 125000, q4: 140000, target: 357000 }, // Platinum 1.44
      { q1: 82000,  q2: 91000,  q3: 87000,  q4: 95000,  target: 340000 }, // Gold 1.044
      { q1: 70000,  q2: 75000,  q3: 72000,  q4: 78000,  target: 300000 }, // Silver 0.983
      { q1: 50000,  q2: 55000,  q3: 60000,  q4: 65000,  target: 322000 }, // Bronze 0.714
      { q1: 95000,  q2: 98000,  q3: 97000,  q4: 100000, target: 370000 }, // Gold 1.054
    ])
    expect(evaluateFormula(s, '=K2')).toBe('Platinum')
    expect(evaluateFormula(s, '=K3')).toBe('Gold')
    expect(evaluateFormula(s, '=K4')).toBe('Silver')
    expect(evaluateFormula(s, '=K5')).toBe('Bronze')
    expect(evaluateFormula(s, '=K6')).toBe('Gold')
  })

  it('Dataset B — product pricing sheet: margin formula chain', () => {
    let s = createBlankSheet('id', 'Sheet1')
    // Revenue, COGS, Margin = (Revenue-COGS)/Revenue, Grade = IF(Margin>=0.4,"A",IF(Margin>=0.25,"B","C"))
    const products = [
      { rev: 1000, cogs: 550 }, // margin 0.45 → A
      { rev: 800,  cogs: 600 }, // margin 0.25 → B
      { rev: 500,  cogs: 450 }, // margin 0.10 → C
    ]
    products.forEach(({ rev, cogs }, i) => {
      const r = i + 2
      s = upsertCell(s, `B${r}`, String(rev))
      s = upsertCell(s, `C${r}`, String(cogs))
      s = upsertCell(s, `D${r}`, `=B${r}-C${r}`)            // gross profit
      s = upsertCell(s, `E${r}`, `=D${r}/B${r}`)            // margin
      s = upsertCell(s, `F${r}`, `=IF(E${r}>=0.4,"A",IF(E${r}>=0.25,"B","C"))`)
    })
    expect(evaluateFormula(s, '=D2')).toBe('450')
    expect(Number(evaluateFormula(s, '=E2'))).toBeCloseTo(0.45)
    expect(evaluateFormula(s, '=F2')).toBe('A')
    expect(evaluateFormula(s, '=F3')).toBe('B')
    expect(evaluateFormula(s, '=F4')).toBe('C')
  })

  it('Dataset C — COUNTIF/SUMIF over formula-driven Tier column matches Univer output', () => {
    const s = buildSheet([
      { q1: 120000, q2: 130000, q3: 125000, q4: 140000, target: 357000 }, // Platinum
      { q1: 82000,  q2: 91000,  q3: 87000,  q4: 95000,  target: 340000 }, // Gold
      { q1: 70000,  q2: 75000,  q3: 72000,  q4: 78000,  target: 300000 }, // Silver
      { q1: 50000,  q2: 55000,  q3: 60000,  q4: 65000,  target: 322000 }, // Bronze
      { q1: 95000,  q2: 98000,  q3: 97000,  q4: 100000, target: 370000 }, // Gold
    ])
    // Two Gold rows (rows 3 and 6)
    expect(evaluateFormula(s, '=COUNTIF(K2:K6,"Gold")')).toBe('2')
    expect(evaluateFormula(s, '=COUNTIF(K2:K6,"Platinum")')).toBe('1')
    expect(evaluateFormula(s, '=COUNTIF(K2:K6,"Bronze")')).toBe('1')

    // Sum bonuses for Gold tier
    const goldBonus = 355000 * 0.1 + 390000 * 0.1
    expect(Number(evaluateFormula(s, '=SUMIF(K2:K6,"Gold",L2:L6)'))).toBeCloseTo(goldBonus)
  })
})

// ---------------------------------------------------------------------------
// Expression parser upgrades: cycles, absolute refs in functions, parentheses,
// function calls inside arithmetic, & concatenation, unary minus
// ---------------------------------------------------------------------------

describe('expression parser correctness', () => {
  it('circular references return #REF! instead of crashing', () => {
    let s = createBlankSheet('id', 'Test')
    s = upsertCell(s, 'A1', '=A1')
    s = upsertCell(s, 'B1', '=C1')
    s = upsertCell(s, 'C1', '=B1')
    expect(evaluateFormula(s, '=A1')).toBe('#REF!')
    expect(evaluateFormula(s, '=B1')).toBe('#REF!')
  })

  it('absolute references work inside range functions', () => {
    let s = createBlankSheet('id', 'Test')
    s = upsertCell(s, 'A1', '1')
    s = upsertCell(s, 'A2', '2')
    s = upsertCell(s, 'A3', '3')
    expect(evaluateFormula(s, '=SUM($A$1:$A$3)')).toBe('6')
    expect(evaluateFormula(s, '=AVERAGE($A$1:A3)')).toBe('2')
  })

  it('function calls compose with arithmetic and parentheses', () => {
    let s = createBlankSheet('id', 'Test')
    s = upsertCell(s, 'A1', '1')
    s = upsertCell(s, 'A2', '2')
    s = upsertCell(s, 'B1', '3')
    expect(evaluateFormula(s, '=SUM(A1:A2)+5')).toBe('8')
    expect(evaluateFormula(s, '=(A1+B1)*2')).toBe('8')
    expect(evaluateFormula(s, '=SUM(A1:A2)*SUM(A1:A2)')).toBe('9')
    expect(evaluateFormula(s, '=-A1+5')).toBe('4')
  })

  it('supports & concatenation including dynamic COUNTIF criteria', () => {
    let s = createBlankSheet('id', 'Test')
    s = upsertCell(s, 'G2', '10')
    s = upsertCell(s, 'G3', '20')
    s = upsertCell(s, 'G4', '30')
    expect(evaluateFormula(s, '="Rank "&G2')).toBe('Rank 10')
    // The dynamic-criteria ranking idiom
    expect(evaluateFormula(s, '=COUNTIF($G$2:$G$4,">"&G2)+1')).toBe('3')
    expect(evaluateFormula(s, '=COUNTIF($G$2:$G$4,">"&G4)+1')).toBe('1')
  })

  it('propagates errors through arithmetic', () => {
    let s = createBlankSheet('id', 'Test')
    s = upsertCell(s, 'A1', '10')
    s = upsertCell(s, 'B1', '0')
    s = upsertCell(s, 'C1', '=A1/B1')
    expect(evaluateFormula(s, '=C1+1')).toBe('#DIV/0!')
  })
})
