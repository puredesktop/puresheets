/**
 * Regression tests for the PureSheets data-integrity fixes (issue #175 sweep):
 *
 *  F1/F8 — blank cells must never satisfy a numeric criterion, and equality
 *          compares numbers numerically (COUNTIF/SUMIF/AVERAGEIF).
 *  F2    — the agent trust boundary derives a cell's kind from its value
 *          instead of trusting the model's self-declared label.
 *  F3    — circular references return #REF! while deep *acyclic* formula
 *          chains (>64 deep) still evaluate.
 *  F4    — a cell tagged "number" whose value is not a finite number is
 *          serialized as its literal string, never as null (data loss).
 */
import { describe, expect, it } from 'vitest'
import {
  applyAgentCellChangesDirect,
  cloneWithUpdatedSheet,
  createBlankWorkbook,
  deriveCellKind,
  evaluateFormula,
  upsertCell,
} from './workbookModel'
import { addConditionalFormatRule } from './sheetDataFeatures'
import { createUniverSnapshot } from './univerAdapter'
import type { PureSheetsDocument, WorkbookSheet } from '../types'

function sheet(cells: Record<string, string>): WorkbookSheet {
  let s = createBlankWorkbook('Integrity').workbook.sheets[0]
  for (const [key, value] of Object.entries(cells)) {
    s = upsertCell(s, key, value)
  }
  return s
}

// ---------------------------------------------------------------------------
// F1 / F8 — blank cells and numeric equality in *IF aggregates
// ---------------------------------------------------------------------------

describe('F1: blank cells never match numeric criteria', () => {
  // A1,A2,A5 hold numbers; A3,A4 are blank (absent) inside the range.
  const s = sheet({ A1: '5', A2: '10', A5: '0' })

  it('COUNTIF(">=0") counts only the real numbers, not the blanks', () => {
    expect(evaluateFormula(s, '=COUNTIF(A1:A5,">=0")')).toBe('3')
  })

  it('COUNTIF("<50") does not count blank cells as 0', () => {
    expect(evaluateFormula(s, '=COUNTIF(A1:A5,"<50")')).toBe('3')
  })

  it('SUMIF(">=0") sums only the real numbers', () => {
    expect(evaluateFormula(s, '=SUMIF(A1:A5,">=0")')).toBe('15')
  })

  it('AVERAGEIF(">=0") averages only the real numbers', () => {
    // (5 + 10 + 0) / 3 = 5, not diluted by two phantom zeros.
    expect(evaluateFormula(s, '=AVERAGEIF(A1:A5,">=0")')).toBe('5')
  })

  it('COUNTIF("") counts the blank cells', () => {
    expect(evaluateFormula(s, '=COUNTIF(A1:A5,"")')).toBe('2')
  })
})

describe('F8: equality compares numbers numerically', () => {
  it('a "5" criterion matches a cell holding 5.0', () => {
    const s = sheet({ A1: '5', A2: '5.0', A3: '6' })
    expect(evaluateFormula(s, '=COUNTIF(A1:A3,"5")')).toBe('2')
  })
})

// ---------------------------------------------------------------------------
// F2 — derive cell kind from value in the agent trust boundary
// ---------------------------------------------------------------------------

describe('F2: agent cell kind is derived from value, not trusted', () => {
  it('deriveCellKind classifies by value', () => {
    expect(deriveCellKind('=SUM(A1:A2)')).toBe('formula')
    expect(deriveCellKind('42')).toBe('number')
    expect(deriveCellKind('-3.5')).toBe('number')
    expect(deriveCellKind('hello')).toBe('text')
    expect(deriveCellKind('   ')).toBe('blank')
    expect(deriveCellKind('')).toBe('blank')
  })

  function proposedKindFor(value: string, claimedKind: 'text' | 'number' | 'formula' | 'blank'): string {
    const base = createBlankWorkbook('Agent trust')
    const sheetId = base.workbook.sheets[0].id
    const result = applyAgentCellChangesDirect(
      base,
      sheetId,
      'B2:B2',
      'apply model change',
      [{ cell: 'B2', value, kind: claimedKind }],
    )
    const logged = result.document.workbook.agentLog?.[0]?.cells?.find(
      c => c.cell === 'B2',
    )
    if (!logged) throw new Error('expected a logged agent change for B2')
    return logged.kind
  }

  it('a formula value mislabeled "text" is stored as a formula', () => {
    expect(proposedKindFor('=SUM(A1:A1)', 'text')).toBe('formula')
  })

  it('a literal number mislabeled "formula" is stored as a number', () => {
    expect(proposedKindFor('42', 'formula')).toBe('number')
  })
})

// ---------------------------------------------------------------------------
// F3 — circular refs vs deep acyclic chains
// ---------------------------------------------------------------------------

describe('F3: circular references are caught without breaking deep chains', () => {
  it('a direct self-reference returns #REF!', () => {
    const s = upsertCell(createBlankWorkbook('Cycle').workbook.sheets[0], 'A1', '=A1')
    expect(evaluateFormula(s, '=A1')).toBe('#REF!')
  })

  it('a two-cell cycle returns #REF!', () => {
    let s = createBlankWorkbook('Cycle2').workbook.sheets[0]
    s = upsertCell(s, 'A1', '=B1')
    s = upsertCell(s, 'B1', '=A1')
    expect(evaluateFormula(s, '=A1')).toBe('#REF!')
  })

  it('a deep acyclic chain (120 deep, past the old 64 cap) evaluates', () => {
    let s = createBlankWorkbook('Chain').workbook.sheets[0]
    const depth = 120
    // A120 = 1, A(n) = A(n+1) + 1  →  A1 = 120
    s = upsertCell(s, `A${depth}`, '1')
    for (let n = depth - 1; n >= 1; n -= 1) {
      s = upsertCell(s, `A${n}`, `=A${n + 1}+1`)
    }
    expect(evaluateFormula(s, '=A1')).toBe(String(depth))
  })
})

// ---------------------------------------------------------------------------
// F4 — a non-numeric "number" cell is never serialized to null
// ---------------------------------------------------------------------------

describe('F4: non-numeric number cells survive serialization', () => {
  it('preserves the literal value instead of writing NaN → null', () => {
    const base = createBlankWorkbook('NaN cell')
    const sheetId = base.workbook.sheets[0].id
    // Force the inconsistent state a stale/legacy file could carry: a cell
    // labelled "number" whose value is not a finite number.
    const doc: PureSheetsDocument = {
      ...base,
      workbook: {
        ...base.workbook,
        sheets: base.workbook.sheets.map(s =>
          s.id === sheetId
            ? { ...s, cells: { ...s.cells, A1: { value: 'N/A', kind: 'number' } } }
            : s,
        ),
      },
    }
    const snapshot = createUniverSnapshot(doc)
    const cellData = snapshot.sheets[sheetId].cellData as Record<
      number,
      Record<number, { v?: unknown }>
    >
    const a1 = cellData[0]?.[0]
    expect(a1?.v).toBe('N/A')
    expect(a1?.v).not.toBeNull()
  })
})

// ---------------------------------------------------------------------------

describe('cloneWithUpdatedSheet keeps engine.snapshot in sync with the model', () => {
  it('regenerates the snapshot from the UPDATED sheet, not the old one (#270)', () => {
    const base = createBlankWorkbook('CF sync')
    const sheetId = base.workbook.sheets[0].id
    const withRule = addConditionalFormatRule(base.workbook.sheets[0], {
      range: 'B1:B4',
      condition: { kind: 'greater', value: '50' },
      style: { fillColor: '#e5f3df' },
    })

    const next = cloneWithUpdatedSheet(base, withRule)

    // Model carries the rule …
    const sheetInModel = next.workbook.sheets.find(s => s.id === sheetId)
    expect(sheetInModel?.conditionalFormats).toHaveLength(1)

    // … and the freshly generated engine snapshot must carry it too, or the
    // canvas reload (which reads engine.snapshot) loses it silently.
    const snapshot = next.engine.snapshot as {
      sheets: Record<
        string,
        { custom?: { puresheets?: { conditionalFormats?: unknown[] | null } } }
      >
    }
    const cfInSnapshot =
      snapshot.sheets[sheetId]?.custom?.puresheets?.conditionalFormats
    expect(cfInSnapshot).toHaveLength(1)
  })
})
