import { describe, expect, it } from 'vitest'
import {
  activeCallContext,
  deleteNamedRange,
  explainFormulaError,
  extractFormulaPrecedents,
  FORMULA_CATALOG,
  isValidNamedRangeName,
  resolveNamedRangesInFormula,
  traceDependents,
  upsertNamedRange,
} from './formulaCatalog'
import {
  clearFormattingInRange,
  mergeRangeInSheet,
  moveSheetInWorkbook,
  profileColumn,
  removeDuplicatesInRange,
  setColumnFilter,
  setFreezePanes,
  setSheetHidden,
  sortRangeMulti,
  splitTextToColumns,
  toggleCheckboxInRange,
  transposeCopiedBlock,
  trimWhitespaceInRange,
  unmergeRangeInSheet,
  visibleRowsWithFilters,
} from './dataTransforms'
import {
  computePivotSheet,
  inferPivotFields,
  isPivotStale,
} from './pivotEngine'
import { extractChartData, recomputeSheetCharts } from './chartEngine'
import {
  addConditionalFormatRule,
  addValidationRule,
  conditionalStyleForCell,
  moveConditionalFormatRule,
  removeConditionalFormatRule,
  validateValue,
} from './sheetDataFeatures'
import {
  addSheet,
  copyRangeCells,
  createBlankWorkbook,
  evaluateFormula,
  setRangeStyle,
  upsertCell,
} from './workbookModel'

describe('Wave 2 — Formula Catalog, Auditing, and Named Ranges', () => {
  it('provides a rich catalog of 150+ categorized spreadsheet functions', () => {
    expect(FORMULA_CATALOG.length).toBeGreaterThanOrEqual(150)
    expect(FORMULA_CATALOG.some(f => f.name === 'XLOOKUP')).toBe(true)
    expect(FORMULA_CATALOG.some(f => f.name === 'SUMIFS')).toBe(true)
    expect(FORMULA_CATALOG.some(f => f.name === 'PMT')).toBe(true)
  })

  it('tracks active function call and parameter index inside nested formulas', () => {
    const formula = '=IF(SUM(A1:A5)>10, XLOOKUP(B2, D2:D10, E2:E10, "Missing"), 0)'
    const insideSum = activeCallContext(formula, formula.indexOf('A1:A5') + 2)
    expect(insideSum?.fn.name).toBe('SUM')
    expect(insideSum?.argIndex).toBe(0)

    const insideXlookupArg3 = activeCallContext(
      formula,
      formula.indexOf('"Missing"') + 1,
    )
    expect(insideXlookupArg3?.fn.name).toBe('XLOOKUP')
    expect(insideXlookupArg3?.argIndex).toBe(3)
  })

  it('explains formula errors and suggests IFERROR fixes', () => {
    const div0 = explainFormulaError('#DIV/0!', '=A1/B1')
    expect(div0.title).toContain('Division by zero')
    expect(div0.suggestedFix?.formula).toBe('=IFERROR(A1/B1, 0)')

    const alreadyWrapped = explainFormulaError('#DIV/0!', '=IFERROR(A1/B1, 0)')
    expect(alreadyWrapped.suggestedFix).toBeUndefined()
  })

  it('extracts precedents and traces cross-sheet dependents', () => {
    let doc = createBlankWorkbook('Audit')
    let s1 = upsertCell(doc.workbook.sheets[0], 'A1', '100')
    s1 = upsertCell(s1, 'B1', '=A1*2')
    doc = { ...doc, workbook: { ...doc.workbook, sheets: [s1] } }
    doc = addSheet(doc, 'Summary')
    const s2 = upsertCell(
      doc.workbook.sheets[1],
      'C3',
      "='Sheet 1'!A1 + 'Sheet 1'!B1",
    )
    doc = { ...doc, workbook: { ...doc.workbook, sheets: [s1, s2] } }

    const precs = extractFormulaPrecedents(s2.cells.C3.value, s2.name)
    expect(precs.map(p => `${p.sheetName}!${p.range}`)).toEqual([
      'Sheet 1!A1',
      'Sheet 1!B1',
    ])

    const deps = traceDependents(doc.workbook, s1.id, 'A1')
    expect(deps.map(d => `${d.sheetName}!${d.cell}`)).toEqual([
      'Sheet 1!B1',
      'Summary!C3',
    ])
  })

  it('creates, resolves, evaluates, and deletes named ranges', () => {
    expect(isValidNamedRangeName('Tax_Rate')).toBe(true)
    expect(isValidNamedRangeName('A1')).toBe(false)

    let doc = createBlankWorkbook('Names')
    let sheet = upsertCell(doc.workbook.sheets[0], 'B2', '0.15')
    sheet = upsertCell(sheet, 'C2', '200')
    doc = { ...doc, workbook: { ...doc.workbook, sheets: [sheet] } }
    doc = upsertNamedRange(doc, {
      name: 'TaxRate',
      sheetId: sheet.id,
      range: 'B2:B2',
    })
    expect(doc.workbook.namedRanges?.length).toBe(1)

    const resolved = resolveNamedRangesInFormula(
      '=C2*TaxRate',
      doc.workbook.namedRanges,
      doc.workbook.sheets,
      sheet.id,
    )
    expect(resolved).toBe('=C2*B2:B2')
    expect(evaluateFormula(sheet, '=C2*TaxRate', doc.workbook)).toBe('30')

    const removed = deleteNamedRange(doc, doc.workbook.namedRanges![0].id)
    expect(removed.workbook.namedRanges).toBeUndefined()
  })
})

describe('Wave 1 & 3 — Grid Polish, Multi-Sort, Smart Filters, Data Cleanup', () => {
  it('merges/unmerges ranges, clears formatting, freezes panes, and toggles checkboxes', () => {
    let sheet = createBlankWorkbook('Polish').workbook.sheets[0]
    sheet = upsertCell(sheet, 'A1', 'Title')
    sheet = upsertCell(sheet, 'B1', 'Extra')
    sheet = setRangeStyle(sheet, 'A1:B1', { bold: true, fillColor: '#eef1f5' })

    const merged = mergeRangeInSheet(sheet, 'A1:B1')
    expect(merged.merges).toEqual(['A1:B1'])
    expect(merged.cells.A1?.value).toBe('Title')
    expect(merged.cells.B1?.value).toBe('')

    const unmerged = unmergeRangeInSheet(merged, 'A1:B1')
    expect(unmerged.merges).toBeUndefined()

    const unstyled = clearFormattingInRange(unmerged, 'A1:B1')
    expect(unstyled.cells.A1?.style).toBeUndefined()
    expect(unstyled.cells.A1?.value).toBe('Title')

    const frozen = setFreezePanes(unstyled, 2, 3)
    expect(frozen.frozenRows).toBe(2)
    expect(frozen.frozenColumns).toBe(3)

    const checked = toggleCheckboxInRange(frozen, 'C1:C2')
    expect(checked.cells.C1?.checkbox).toBe(true)
    expect(checked.cells.C1?.value).toBe('FALSE')
  })

  it('sorts multi-level while keeping headers pinned', () => {
    let sheet = createBlankWorkbook('Sort').workbook.sheets[0]
    const data = [
      ['Region', 'Rep', 'Sales'],
      ['West', 'Zoe', '100'],
      ['East', 'Ana', '300'],
      ['West', 'Ben', '250'],
      ['East', 'Cal', '500'],
    ]
    data.forEach((row, r) =>
      row.forEach((v, c) => {
        const col = String.fromCharCode(65 + c)
        sheet = upsertCell(sheet, `${col}${r + 1}`, v)
      }),
    )

    const sorted = sortRangeMulti(
      sheet,
      'A1:C5',
      [
        { column: 0, direction: 'asc' },
        { column: 2, direction: 'desc' },
      ],
      true,
    )
    expect(sorted.cells.A1?.value).toBe('Region')
    expect(sorted.cells.B2?.value).toBe('Cal')
    expect(sorted.cells.B3?.value).toBe('Ana')
    expect(sorted.cells.B4?.value).toBe('Ben')
    expect(sorted.cells.B5?.value).toBe('Zoe')
  })

  it('filters across multiple columns and profiles column statistics', () => {
    let sheet = createBlankWorkbook('Filter').workbook.sheets[0]
    const data = [
      ['Region', 'Amount'],
      ['North', '100'],
      ['South', '250'],
      ['North', '400'],
      ['East', '50'],
    ]
    data.forEach((row, r) =>
      row.forEach((v, c) => {
        const col = String.fromCharCode(65 + c)
        sheet = upsertCell(sheet, `${col}${r + 1}`, v)
      }),
    )

    const profile = profileColumn(sheet, 1, true)
    expect(profile.numericStats?.min).toBe(50)
    expect(profile.numericStats?.max).toBe(400)
    expect(profile.numericStats?.median).toBe(175)

    let filtered = setColumnFilter(sheet, 0, { selectedValues: ['North'] })
    filtered = setColumnFilter(filtered, 1, { operator: 'gt', value: '200' })
    const visible = visibleRowsWithFilters(filtered)
    expect(visible).toEqual([0, 3])
  })

  it('removes duplicates, splits text to columns, trims whitespace, and transposes blocks', () => {
    let sheet = createBlankWorkbook('Clean').workbook.sheets[0]
    sheet = upsertCell(sheet, 'A1', 'Item')
    sheet = upsertCell(sheet, 'A2', '  Apple ,  Red ')
    sheet = upsertCell(sheet, 'A3', 'Banana, Yellow')
    sheet = upsertCell(sheet, 'A4', 'Banana, Yellow')

    const deduped = removeDuplicatesInRange(sheet, 'A1:A4', [0], true)
    expect(deduped.removedCount).toBe(1)

    const trimmed = trimWhitespaceInRange(deduped.sheet, 'A1:A3')
    expect(trimmed.sheet.cells.A2?.value).toBe('Apple , Red')

    const split = splitTextToColumns(trimmed.sheet, 'A2:A3', ',')
    expect(split.rowsSplit).toBe(2)
    expect(split.sheet.cells.A2?.value).toBe('Apple')
    expect(split.sheet.cells.B2?.value).toBe('Red')

    const block = copyRangeCells(split.sheet, 'A2:B3')!
    const transposed = transposeCopiedBlock(block)
    expect(transposed.rows).toBe(2)
    expect(transposed.columns).toBe(2)
    expect(transposed.cells[1][0]?.value).toBe('Red')
    expect(transposed.cells[0][1]?.value).toBe('Banana')
  })

  it('reorders and hides sheets safely without hiding the last visible sheet', () => {
    let doc = createBlankWorkbook('Tabs')
    doc = addSheet(doc, 'Second')
    const moved = moveSheetInWorkbook(doc, doc.workbook.sheets[1].id, 0)
    expect(moved.workbook.sheets[0].name).toBe('Second')

    const hidden = setSheetHidden(moved, moved.workbook.sheets[0].id, true)
    expect(hidden.workbook.sheets[0].hidden).toBe(true)
    // Cannot hide the sole remaining visible sheet
    const stillVisible = setSheetHidden(hidden, hidden.workbook.sheets[1].id, true)
    expect(stillVisible.workbook.sheets[1].hidden).toBeUndefined()
  })
})

describe('Wave 4 & 5 — Rules v2, Live Multi-Series Charts, and Pivot Tables', () => {
  it('supports range-backed list validation and advanced conditional formatting', () => {
    let sheet = createBlankWorkbook('Rules').workbook.sheets[0]
    sheet = upsertCell(sheet, 'F1', 'Open')
    sheet = upsertCell(sheet, 'F2', 'Closed')
    sheet = upsertCell(sheet, 'A1', 'Open')
    sheet = upsertCell(sheet, 'A2', 'Invalid')
    sheet = upsertCell(sheet, 'B1', '10')
    sheet = upsertCell(sheet, 'B2', '50')
    sheet = upsertCell(sheet, 'B3', '50')

    sheet = addValidationRule(sheet, {
      range: 'A1:A2',
      kind: 'list',
      sourceRange: 'F1:F2',
    })
    const rule = sheet.validations![0]
    expect(validateValue(rule, 'Open', sheet).valid).toBe(true)
    expect(validateValue(rule, 'Invalid', sheet).valid).toBe(false)

    sheet = addConditionalFormatRule(sheet, {
      range: 'B1:B3',
      condition: { kind: 'duplicate', value: '' },
      style: { fillColor: '#fdecea' },
    })
    sheet = addConditionalFormatRule(sheet, {
      range: 'B1:B3',
      condition: { kind: 'color-scale', value: '' },
      style: {},
      colorScale: { minColor: '#ffffff', maxColor: '#2f6ea2' },
    })
    expect(conditionalStyleForCell(sheet, 'B2')?.fillColor).toBe('#2f6ea2')
    const moved = moveConditionalFormatRule(sheet, sheet.conditionalFormats![1].id, -1)
    expect(moved.conditionalFormats![0].condition.kind).toBe('color-scale')
    const removed = removeConditionalFormatRule(moved, moved.conditionalFormats![0].id)
    expect(removed.conditionalFormats?.length).toBe(1)
  })

  it('extracts multi-series charts and refreshes reactively when source data changes', () => {
    let sheet = createBlankWorkbook('Charts').workbook.sheets[0]
    const rows = [
      ['Quarter', 'Revenue', 'Costs'],
      ['Q1', '120', '80'],
      ['Q2', '180', '95'],
    ]
    rows.forEach((r, ri) =>
      r.forEach((v, ci) => {
        sheet = upsertCell(sheet, `${String.fromCharCode(65 + ci)}${ri + 1}`, v)
      }),
    )

    const extracted = extractChartData(sheet, 'A1:C3', 'bar')!
    expect(extracted.labels).toEqual(['Q1', 'Q2'])
    expect(extracted.series.map(s => s.name)).toEqual(['Revenue', 'Costs'])
    expect(extracted.series[1].values).toEqual([80, 95])

    sheet = {
      ...sheet,
      charts: [
        {
          id: 'ch-1',
          title: 'Financials',
          type: 'bar',
          range: 'A1:C3',
          labels: extracted.labels,
          values: extracted.values,
          series: extracted.series,
        },
      ],
    }
    sheet = upsertCell(sheet, 'B2', '240')
    const refreshed = recomputeSheetCharts(sheet)
    expect(refreshed.charts![0].values).toEqual([240, 180])
    expect(refreshed.charts![0].series![0].values).toEqual([240, 180])
  })

  it('computes cross-tab pivot sheets with aggregates, grand totals, locks, and stale detection', () => {
    let source = createBlankWorkbook('Sales').workbook.sheets[0]
    const table = [
      ['Region', 'Quarter', 'Revenue'],
      ['North', 'Q1', '100'],
      ['North', 'Q2', '150'],
      ['South', 'Q1', '200'],
      ['South', 'Q1', '50'],
    ]
    table.forEach((r, ri) =>
      r.forEach((v, ci) => {
        source = upsertCell(source, `${String.fromCharCode(65 + ci)}${ri + 1}`, v)
      }),
    )

    const fields = inferPivotFields(source, 'A1:C5')
    expect(fields.map(f => `${f.label}:${f.isNumeric}`)).toEqual([
      'Region:false',
      'Quarter:false',
      'Revenue:true',
    ])

    const pivot = computePivotSheet(
      source,
      {
        id: 'piv-1',
        sourceSheetId: source.id,
        sourceRange: 'A1:C5',
        rows: [{ columnIndex: 0, label: 'Region' }],
        columns: [{ columnIndex: 1, label: 'Quarter' }],
        values: [{ columnIndex: 2, label: 'Revenue', aggregate: 'sum' }],
        refreshedAt: new Date().toISOString(),
      },
      { id: 'sheet-pivot-1', name: 'Pivot 1' },
    )

    expect(pivot.cells.A1?.value).toBe('Region')
    expect(pivot.cells.A2?.value).toBe('North')
    expect(pivot.cells.B2?.value).toBe('100')
    expect(pivot.cells.C2?.value).toBe('150')
    expect(pivot.cells.D2?.value).toBe('250')
    expect(pivot.cells.A3?.value).toBe('South')
    expect(pivot.cells.B3?.value).toBe('250')
    expect(pivot.cells.A4?.value).toBe('Grand Total')
    expect(pivot.cells.D4?.value).toBe('500')
    expect(pivot.protection?.lockedRanges?.length).toBe(1)

    expect(isPivotStale(pivot, source)).toBe(false)
    const editedSource = upsertCell(source, 'C2', '999')
    expect(isPivotStale(pivot, editedSource)).toBe(true)
  })
})
