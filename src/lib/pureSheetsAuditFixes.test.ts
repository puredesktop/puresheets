import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  addSheet,
  applyAgentCellChangesDirect,
  clearRangeStyle,
  copyRangeCells,
  createBlankSheet,
  createBlankWorkbook,
  createChartFromRange,
  deleteColumn,
  deleteRow,
  deleteSheet,
  displayCellValue,
  duplicateSheet,
  evaluateFormula,
  fillRange,
  findMatches,
  hasMergeInRange,
  insertRow,
  lockRange,
  mergeRange,
  moveSheet,
  pasteCellBlock,
  rangeToDelimitedText,
  renameSheet,
  renameSheetReferences,
  replaceFirstOccurrence,
  setCellComment,
  setCellStyle,
  setFreezePanes,
  setRangeStyle,
  sortRange,
  toggleFilterRow,
  unmergeRange,
  upsertCell,
  visibleRowsForSheet,
} from './workbookModel'
import { loadSheetsDocument, validateDocument } from './sheetsDocument'
import { normalizeComments, normalizeStyleObject } from './sheetMetadataNormalize'
import { csvFromSheet, jsonFromSheet, parseCsv, sheetFromCsv, toCsv } from './csv'
import {
  createUniverSnapshot,
  documentFromUniverSnapshot,
  withFreshUniverSnapshot,
} from './univerAdapter'
import { requestedSheetId } from './workbookView'
import { conditionMatches, validateValue } from './sheetDataFeatures'
import { sheetsSnapshotHtml } from './sheetsSnapshot'
import { exportXlsxWorkbook, importXlsxWorkbook } from './xlsxIO'
import { validateSafeFilePath } from './workbookFiles'
import {
  formatRangeHandler,
  setCellCommentHandler,
} from '../agents/handlers'
import type { SheetsAgentToolContext } from '../agents/catalog'
import type { PureSheetsDocument } from '../types'

describe('Security & validation hardening', () => {
  it('blocks CSS injection in textColor/fillColor and clamps fontSize across loaders', () => {
    const raw = JSON.stringify({
      app: 'PureSheets',
      version: 1,
      engine: { name: 'univer' },
      metadata: { title: 'CSS test', createdAt: '2026-01-01T00:00:00Z', updatedAt: '2026-01-01T00:00:00Z' },
      workbook: {
        activeSheetId: 's1',
        sheets: [
          {
            id: 's1',
            name: 'Sheet 1',
            rowCount: 20,
            columnCount: 10,
            cells: {
              A1: {
                value: 'x',
                kind: 'text',
                style: {
                  textColor: 'red; background-image: url(https://evil.example/x)',
                  fillColor: 'expression(alert(1))',
                  fontSize: 9999,
                  bold: true,
                },
              },
              A2: {
                value: 'y',
                kind: 'text',
                style: {
                  textColor: '#AbC',
                  fillColor: '#112233',
                  fontSize: 2,
                },
              },
            },
          },
        ],
      },
    })
    const doc = loadSheetsDocument(raw)
    const cells = doc.workbook.sheets[0].cells
    expect(cells.A1.style?.textColor).toBeUndefined()
    expect(cells.A1.style?.fillColor).toBeUndefined()
    expect(cells.A1.style?.fontSize).toBe(96)
    expect(cells.A1.style?.bold).toBe(true)
    expect(cells.A2.style?.textColor).toBe('#aabbcc')
    expect(cells.A2.style?.fillColor).toBe('#112233')
    expect(cells.A2.style?.fontSize).toBe(6)

    expect(
      normalizeStyleObject({
        textColor: 'javascript:alert(1)',
        fillColor: '#f0a',
        bold: true,
      }),
    ).toEqual({ fillColor: '#ff00aa', bold: true })
  })

  it('validates hex colors in xlsx export without throwing on malformed model colors', async () => {
    const doc = createBlankWorkbook('Color export')
    doc.workbook.sheets[0].cells.A1 = {
      value: 'Safe',
      kind: 'text',
      style: { textColor: 'not-a-hex' as string, fillColor: '#0f0' },
    }
    const { bytes } = await exportXlsxWorkbook(doc)
    expect(bytes.byteLength).toBeGreaterThan(0)
  })

  it('neutralizes CSV formula injection on export while preserving signed numbers', () => {
    const out = toCsv([
      ['=SUM(A1:A2)', '+1+1', '-1+cmd|"/C calc"!A0', '@SUM(1,2)', '\t=1', ' =1', '\uFEFF=1'],
      ['-42.5', '-100', '0.25', '3.14', 'plain', '100', '-1e5'],
    ])
    const [row1, row2] = out.split('\n')
    expect(row1).toBe(
      `"'=SUM(A1:A2)","'+1+1","'-1+cmd|""/C calc""!A0","'@SUM(1,2)","'\t=1","' =1","'\uFEFF=1"`,
    )
    expect(row2).toBe('-42.5,-100,0.25,3.14,plain,100,-1e5')
  })

  it('strips UTF-8 BOM on CSV import and stores = fields as literal text, never formulas', () => {
    const sheet = sheetFromCsv('\uFEFFHeader,Formula,Num\nAlice,"=SUM(1,2)",-42.5\n', 's1', 'Imp')
    expect(sheet.cells.A1).toMatchObject({ value: 'Header', kind: 'text' })
    expect(sheet.cells.B2).toMatchObject({ value: '=SUM(1,2)', kind: 'text' })
    expect(sheet.cells.C2).toMatchObject({ value: '-42.5', kind: 'number' })
  })

  it('blocks prototype pollution keys, bounds sheet dimensions, and validates agentLog kinds', () => {
    const parsed = JSON.parse(
      `{
        "app": "PureSheets",
        "version": 1,
        "engine": { "name": "univer" },
        "metadata": { "title": "Proto", "createdAt": "2026-01-01T00:00:00Z", "updatedAt": "2026-01-01T00:00:00Z" },
        "workbook": {
          "activeSheetId": "s1",
          "agentLog": [{
            "id": "log-1",
            "at": "2026-01-01T00:00:00Z",
            "agentName": "Bot",
            "tool": "applyCellChanges",
            "summary": "test",
            "sheetId": "s1",
            "sheetName": "Sheet 1",
            "cells": [
              { "cell": "A1", "after": "1", "kind": "number" },
              { "cell": "A2", "after": "2", "kind": "evil" },
              { "cell": "__proto__", "after": "3", "kind": "text" }
            ]
          }],
          "sheets": [{
            "id": "s1",
            "name": "Sheet 1",
            "rowCount": 9999999,
            "columnCount": 999999,
            "cells": {
              "__proto__": { "value": "pwn", "kind": "text" },
              "constructor": { "value": "pwn", "kind": "text" },
              "ZZZZ9999999": { "value": "oob", "kind": "text" },
              "A1": { "value": "ok", "kind": "text" }
            },
            "columnWidths": { "__proto__": 100, "-5": 100, "0": 50000, "1": 120 }
          }]
        }
      }`,
    ) as unknown
    const doc = validateDocument(parsed)
    const sheet = doc.workbook.sheets[0]
    expect(sheet.rowCount).toBeLessThanOrEqual(100_000)
    expect(sheet.columnCount).toBeLessThanOrEqual(16_384)
    expect(Object.keys(sheet.cells)).toEqual(['A1'])
    expect(sheet.columnWidths).toEqual({ '0': 2000, '1': 120 })
    expect(doc.workbook.agentLog?.[0].cells).toEqual([
      { cell: 'A1', after: '1', kind: 'number' },
    ])

    const snap = createUniverSnapshot(doc)
    expect(
      requestedSheetId({
        ...snap,
        custom: { puresheets: { activeSheetId: 'toString' } },
      }),
    ).toBeNull()
  })

  it('caps oversized formula range expansion and formatRangeHandler cleanly', async () => {
    const sheet = createBlankSheet('s1', 'Sheet 1')
    expect(evaluateFormula(sheet, '=SUM(A1:Z10000)')).toBe('#NUM!')

    let doc = createBlankWorkbook('Cap test')
    const ctx: SheetsAgentToolContext = {
      get document() { return doc },
      setDocument: u => { doc = u(doc) },
      filePath: '',
      selectedRange: 'A1:A1',
    }
    const oversized = await formatRangeHandler(ctx, { range: 'A1:Z500', bold: true })
    expect(oversized.isError).toBe(true)

    await formatRangeHandler(ctx, { range: 'A1:A1', fontSize: 500 })
    expect(doc.workbook.sheets[0].cells.A1?.style?.fontSize).toBe(96)
  })

  it('fixes setCellCommentHandler single-cell validation, protection, and length cap', async () => {
    let doc = createBlankWorkbook('Comments')
    const ctx: SheetsAgentToolContext = {
      get document() { return doc },
      setDocument: u => { doc = u(doc) },
      filePath: '',
      selectedRange: 'A1:A1',
    }
    const ok = await setCellCommentHandler(ctx, { cell: 'B4', comment: 'Check total' })
    expect(ok.isError).toBeUndefined()
    expect(doc.workbook.sheets[0].comments?.B4).toBe('Check total')

    const tooLong = await setCellCommentHandler(ctx, { cell: 'B4', comment: 'x'.repeat(9000) })
    expect(tooLong.isError).toBe(true)

    doc.workbook.sheets[0] = lockRange(doc.workbook.sheets[0], 'B4:B4')
    const locked = await setCellCommentHandler(ctx, { cell: 'B4', comment: 'Blocked' })
    expect(locked.isError).toBe(true)
  })

  it('rejects NUL bytes and parent traversal in validateSafeFilePath', () => {
    expect(validateSafeFilePath('/Users/test/budget.xlsx')).toBe('/Users/test/budget.xlsx')
    expect(() => validateSafeFilePath('/Users/test/../secret.xlsx')).toThrow()
    expect(() => validateSafeFilePath('/Users/test/a\0.xlsx')).toThrow()
  })
})

describe('Formula engine & spreadsheet core fixes', () => {
  it('rewrites cross-sheet formulas across all sheets on renameSheet and rejects duplicate names', () => {
    let doc = createBlankWorkbook('Cross')
    doc = addSheet(doc, 'Q1 Sales')
    doc.workbook.sheets[0] = upsertCell(
      doc.workbook.sheets[0],
      'A1',
      `='Q1 Sales'!B2 + SUM('Q1 Sales'!B2:B4) + IF("Q1 Sales!B2"="x",0,1)`,
    )
    const renamed = renameSheet(doc, doc.workbook.sheets[1].id, 'Annual')
    expect(renamed.workbook.sheets[1].name).toBe('Annual')
    expect(renamed.workbook.sheets[0].cells.A1?.value).toBe(
      `=Annual!B2 + SUM(Annual!B2:B4) + IF("Q1 Sales!B2"="x",0,1)`,
    )
    // Duplicate rename is rejected
    const dup = renameSheet(renamed, renamed.workbook.sheets[1].id, 'sheet 1')
    expect(dup).toBe(renamed)
  })

  it('shifts absolute ranges ($A$1:$A$10) and preserves single-quoted sheet names & scientific notation', () => {
    let sheet = createBlankSheet('s1', 'Sheet 1')
    sheet = upsertCell(sheet, 'B1', '=SUM($A$1:$A$10)')
    sheet = upsertCell(sheet, 'C1', `='Q1'!A1 + 1E2 + A1`)
    const inserted = insertRow(sheet, 0)
    expect(inserted.cells.B2?.value).toBe('=SUM($A$2:$A$11)')
    expect(inserted.cells.C2?.value).toBe(`='Q1'!A1 + 1E2 + A2`)

    // Paste & fill preserve 'Q1' and 1E2
    const block = copyRangeCells(sheet, 'C1:C1')!
    const pasted = pasteCellBlock(sheet, 'C2', block, 'all')
    expect(pasted.cells.C2?.value).toBe(`='Q1'!A2 + 1E2 + A2`)

    const filled = fillRange(sheet, 'C1:C1', 'C2:C2')
    expect(filled.cells.C2?.value).toBe(`='Q1'!A2 + 1E2 + A2`)
  })

  it('fixes freeze panes startRow/startColumn off-by-one and supports setFreezePanes', () => {
    let doc = createBlankWorkbook('Freeze')
    doc.workbook.sheets[0] = setFreezePanes(doc.workbook.sheets[0], 2, 3)
    const snap = createUniverSnapshot(doc)
    expect(snap.sheets['sheet-1'].freeze).toEqual({
      startRow: 2,
      startColumn: 3,
      ySplit: 2,
      xSplit: 3,
    })

    const unfrozen = createUniverSnapshot(createBlankWorkbook('Unfrozen'))
    expect(unfrozen.sheets['sheet-1'].freeze).toEqual({
      startRow: -1,
      startColumn: -1,
      ySplit: 0,
      xSplit: 0,
    })
  })

  it('preserves agentLog and generates unique copy names in duplicateSheet', () => {
    let doc = createBlankWorkbook('Dup')
    doc = {
      ...doc,
      workbook: {
        ...doc.workbook,
        agentLog: [
          {
            id: 'log-1',
            at: '2026-01-01T00:00:00Z',
            agentName: 'Assistant',
            tool: 'applyCellChanges',
            summary: 'Added data',
            sheetId: 'sheet-1',
            sheetName: 'Sheet 1',
          },
        ],
      },
    }
    const d1 = duplicateSheet(doc, 'sheet-1')
    expect(d1.workbook.agentLog).toHaveLength(1)
    expect(d1.workbook.sheets[1].name).toBe('Sheet 1 copy')
    const d2 = duplicateSheet(d1, 'sheet-1')
    expect(d2.workbook.sheets[2].name).toBe('Sheet 1 copy 2')
    expect(d2.workbook.agentLog).toHaveLength(1)
  })

  it('enforces protection on row/column deletion and shifts chart ranges on structural edits', () => {
    let sheet = createBlankSheet('s1', 'Sheet 1')
    sheet = upsertCell(sheet, 'A2', 'North')
    sheet = upsertCell(sheet, 'B2', '100')
    sheet = upsertCell(sheet, 'A3', 'South')
    sheet = upsertCell(sheet, 'B3', '200')
    sheet = createChartFromRange(sheet, 'A2:B3', 'bar')
    sheet = lockRange(sheet, 'A2:A2')

    // Deleting row containing locked populated cell is refused
    expect(deleteRow(sheet, 1)).toBe(sheet)
    expect(deleteColumn(sheet, 0)).toBe(sheet)

    // Inserting a row above shifts the chart range
    const inserted = insertRow(sheet, 0)
    expect(inserted.charts?.[0].range).toBe('A3:B4')
  })

  it('sorts formatted numbers numerically, remaps comments, preserves header row, and handles cross-sheet formulas', () => {
    let doc = createBlankWorkbook('Sort')
    doc = addSheet(doc, 'Rates')
    doc.workbook.sheets[1] = upsertCell(doc.workbook.sheets[1], 'A1', '10')

    let sheet = doc.workbook.sheets[0]
    sheet = upsertCell(sheet, 'A1', 'Amount')
    sheet = upsertCell(sheet, 'B1', 'Calc')
    sheet = upsertCell(sheet, 'A2', '1200')
    sheet = setCellStyle(sheet, 'A2', { numberFormat: 'currency' })
    sheet = upsertCell(sheet, 'B2', '=Rates!A1+A2+A1')
    sheet = setCellComment(sheet, 'A2', 'Big note')
    sheet = upsertCell(sheet, 'A3', '900')
    sheet = setCellStyle(sheet, 'A3', { numberFormat: 'currency' })
    sheet = upsertCell(sheet, 'B3', '=Rates!A1+A3+A1')
    sheet = setCellComment(sheet, 'A3', 'Small note')
    doc.workbook.sheets[0] = sheet

    const sorted = sortRange(sheet, 'A1:B3', 'asc', doc.workbook, { hasHeader: true })
    expect(sorted.cells.A1?.value).toBe('Amount')
    expect(sorted.cells.A2?.value).toBe('900')
    expect(sorted.cells.A3?.value).toBe('1200')
    expect(sorted.comments?.A2).toBe('Small note')
    expect(sorted.comments?.A3).toBe('Big note')
    // Cross-sheet ref Rates!A1 and outside-band header ref A1 stay anchored; in-band A3/A2 shifted
    expect(sorted.cells.B2?.value).toBe('=Rates!A1+A2+A1')
    expect(sorted.cells.B3?.value).toBe('=Rates!A1+A3+A1')

    // Cross-sheet formulas work in findMatches, rangeToDelimitedText, copyRangeCells, visibleRowsForSheet, charts
    doc.workbook.sheets[0] = upsertCell(sheet, 'C2', '=Rates!A1*5')
    expect(findMatches(doc.workbook.sheets[0], '50', {}, doc.workbook)).toContain('C2')
    expect(rangeToDelimitedText(doc.workbook.sheets[0], 'C2:C2', doc.workbook)).toBe('50')
    expect(copyRangeCells(doc.workbook.sheets[0], 'C2:C2', doc.workbook)?.computedValues).toEqual([['50']])
    const filtered = { ...toggleFilterRow(doc.workbook.sheets[0], 2), filterQuery: '50' }
    expect(visibleRowsForSheet(filtered, doc.workbook)).toContain(1)
    const charted = createChartFromRange(doc.workbook.sheets[0], 'C2:C2', 'bar', doc.workbook)
    expect(charted.charts?.[0].values).toEqual([50])
  })

  it('fixes ROUND negative places, IFERROR # text, empty formatCellValue, overflow, VLOOKUP, replaceFirstOccurrence, paste values, and data features', () => {
    let sheet = createBlankSheet('s1', 'Sheet 1')
    sheet = upsertCell(sheet, 'A1', 'South')
    sheet = upsertCell(sheet, 'B1', '200')
    sheet = setCellStyle(sheet, 'C1', { numberFormat: 'currency' })
    sheet = setCellStyle(sheet, 'C2', { numberFormat: 'percent' })

    expect(evaluateFormula(sheet, '=ROUND(1234.56, -2)')).toBe('1200')
    expect(evaluateFormula(sheet, '=ROUND(2.567, 2)')).toBe('2.57')
    expect(evaluateFormula(sheet, '=ROUND("abc", 2)')).toBe('#VALUE!')
    expect(evaluateFormula(sheet, '=IFERROR("#1 priority", "fallback")')).toBe('#1 priority')
    expect(evaluateFormula(sheet, '=IFERROR(1/0, "fallback")')).toBe('fallback')
    expect(displayCellValue(sheet, 'C1')).toBe('')
    expect(displayCellValue(sheet, 'C2')).toBe('')
    expect(evaluateFormula(sheet, '=1000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000*1000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000')).toBe('#NUM!')
    expect(evaluateFormula(sheet, '="   "+1')).toBe('#VALUE!')
    expect(evaluateFormula(sheet, '=VLOOKUP("south",A1:B1,2,FALSE)')).toBe('200')
    expect(evaluateFormula(sheet, '=VLOOKUP("south",A1:B1,5,FALSE)')).toBe('#REF!')

    // replaceFirstOccurrence literal $100 / $$
    let repSheet = upsertCell(sheet, 'D1', 'Price: OLD')
    repSheet = replaceFirstOccurrence(repSheet, 'D1', 'OLD', '$100 $$').sheet
    expect(repSheet.cells.D1?.value).toBe('Price: $100 $$')

    // pasteCellBlock values mode keeps = text as text
    const eqSheet = upsertCell(sheet, 'E1', `="=notAFormula"`)
    const eqBlock = copyRangeCells(eqSheet, 'E1:E1')!
    const pastedEq = pasteCellBlock(eqSheet, 'E2', eqBlock, 'values')
    expect(pastedEq.cells.E2).toMatchObject({ value: '=notAFormula', kind: 'text' })

    // sheetDataFeatures date validation & blank text-contains
    expect(validateValue({ id: 'v1', range: 'A1', kind: 'date' }, '1').valid).toBe(false)
    expect(validateValue({ id: 'v1', range: 'A1', kind: 'date' }, '2026-02-30').valid).toBe(false)
    expect(validateValue({ id: 'v1', range: 'A1', kind: 'date' }, '2026-02-28').valid).toBe(true)
    expect(conditionMatches({ kind: 'text-contains', value: '   ' }, 'anything')).toBe(false)
  })
})

describe('Persistence, snapshot round-trip & must-have features', () => {
  it('grows rowCount/columnCount in applyAgentCellChangesDirect and no-ops identical upsertCell', () => {
    const doc = createBlankWorkbook('Grow')
    const sheet = doc.workbook.sheets[0]
    const s1 = upsertCell(sheet, 'A1', 'hello')
    expect(upsertCell(s1, 'A1', 'hello')).toBe(s1)
    expect(upsertCell(sheet, 'Z9', '')).toBe(sheet)

    const result = applyAgentCellChangesDirect(
      doc,
      sheet.id,
      'A1:Z60',
      'Write beyond initial grid',
      [{ cell: 'Z60', value: '99', kind: 'number' }],
    )
    expect(result.applied).toBe(1)
    expect(result.document.workbook.sheets[0].rowCount).toBeGreaterThanOrEqual(60)
    expect(result.document.workbook.sheets[0].columnCount).toBeGreaterThanOrEqual(26)
  })

  it('preserves cleared filterRow/filterColumn, supports inline cell.s, and keeps stable unit id across title renames', () => {
    let doc = createBlankWorkbook('Original Title')
    const idBeforeRename = createUniverSnapshot(doc).id
    const renamedTitleDoc: PureSheetsDocument = {
      ...doc,
      metadata: { ...doc.metadata, title: 'Renamed Title' },
    }
    expect(createUniverSnapshot(renamedTitleDoc).id).toBe(idBeforeRename)

    // Filter cleared round-trip
    doc.workbook.sheets[0] = toggleFilterRow(doc.workbook.sheets[0], 1)
    const fallbackWithFilter = withFreshUniverSnapshot(doc)
    const clearedDoc: PureSheetsDocument = {
      ...fallbackWithFilter,
      workbook: {
        ...fallbackWithFilter.workbook,
        sheets: [toggleFilterRow(fallbackWithFilter.workbook.sheets[0], 1)],
      },
    }
    const clearedSnapshot = createUniverSnapshot(clearedDoc)
    const restored = documentFromUniverSnapshot(clearedSnapshot, fallbackWithFilter)
    expect(restored.workbook.sheets[0].filterRow).toBeUndefined()
    expect(restored.workbook.sheets[0].filterColumn).toBeUndefined()

    // Inline object cell.s support
    const snap = createUniverSnapshot(createBlankWorkbook('Inline style'))
    snap.sheets['sheet-1'].cellData = {
      0: { 0: { v: 'Styled', t: 1, s: { bl: 1, cl: { rgb: '#123456' } } } },
    }
    const inlineRestored = documentFromUniverSnapshot(snap)
    expect(inlineRestored.workbook.sheets[0].cells.A1?.style).toMatchObject({
      bold: true,
      textColor: '#123456',
    })
  })

  it('supports Merge/Unmerge, Clear Formatting, Reset Color, Sheet Reordering, Selection Export, and Computed Snapshots', () => {
    let doc = createBlankWorkbook('Features')
    doc = addSheet(doc, 'Second')
    doc = addSheet(doc, 'Third')
    expect(doc.workbook.sheets.map(s => s.name)).toEqual(['Sheet 1', 'Second', 'Third'])

    const movedLeft = moveSheet(doc, doc.workbook.sheets[2].id, 'left')
    expect(movedLeft.workbook.sheets.map(s => s.name)).toEqual(['Sheet 1', 'Third', 'Second'])
    const movedRight = moveSheet(movedLeft, movedLeft.workbook.sheets[0].id, 'right')
    expect(movedRight.workbook.sheets.map(s => s.name)).toEqual(['Third', 'Sheet 1', 'Second'])

    // Merge / Unmerge
    let sheet = createBlankSheet('s1', 'Sheet 1')
    sheet = upsertCell(sheet, 'A1', 'Title')
    sheet = upsertCell(sheet, 'B1', 'Extra')
    sheet = mergeRange(sheet, 'A1:B1')
    expect(sheet.merges).toEqual(['A1:B1'])
    expect(hasMergeInRange(sheet, 'B1:B1')).toBe(true)
    expect(sheet.cells.A1?.value).toBe('Title')
    expect(sheet.cells.B1).toBeUndefined()
    sheet = unmergeRange(sheet, 'A1:B1')
    expect(sheet.merges).toBeUndefined()
    expect(hasMergeInRange(sheet, 'A1:B1')).toBe(false)

    // Clear Formatting & Reset Color
    sheet = setRangeStyle(sheet, 'A1:A1', { bold: true, fillColor: '#ffeeaa', textColor: '#112233', wrap: 'wrap', verticalAlign: 'top' })
    expect(sheet.cells.A1?.style?.fillColor).toBe('#ffeeaa')
    sheet = setRangeStyle(sheet, 'A1:A1', { fillColor: undefined })
    expect(sheet.cells.A1?.style?.fillColor).toBeUndefined()
    expect(sheet.cells.A1?.style?.textColor).toBe('#112233')
    sheet = clearRangeStyle(sheet, 'A1:A1')
    expect(sheet.cells.A1?.style).toBeUndefined()
    expect(sheet.cells.A1?.value).toBe('Title')

    // Selection TSV / JSON export
    let exp = createBlankSheet('s1', 'Sheet 1')
    exp = upsertCell(exp, 'A1', 'Item')
    exp = upsertCell(exp, 'B1', 'Price')
    exp = upsertCell(exp, 'A2', 'Book')
    exp = upsertCell(exp, 'B2', '15')
    expect(csvFromSheet(exp, 'A1:B2', undefined, '\t')).toBe('Item\tPrice\nBook\t15\n')
    expect(JSON.parse(jsonFromSheet(exp, 'A1:B2'))).toEqual([{ Item: 'Book', Price: 15 }])

    // sheetsSnapshotHtml computes formula values
    const snapDoc = createBlankWorkbook('Snap')
    snapDoc.workbook.sheets[0] = upsertCell(
      upsertCell(snapDoc.workbook.sheets[0], 'A1', '21'),
      'B1',
      '=A1*2',
    )
    const html = sheetsSnapshotHtml(snapDoc)!
    expect(html).toContain('42')
    expect(html).not.toContain('=A1*2')
  })

  it('coerces blank cell references to 0 in arithmetic while rejecting blank string literals', () => {
    const base = createBlankSheet('s1', 'Sheet 1')
    const sheet = {
      ...base,
      cells: {
        B1: { value: '   ', kind: 'text' as const },
        C1: { value: 'hello', kind: 'text' as const },
      },
    }
    expect(evaluateFormula(sheet, '=A1+5')).toBe('5')
    expect(evaluateFormula(sheet, '=-A1')).toBe('0')
    expect(evaluateFormula(sheet, '=A1*2')).toBe('0')
    expect(evaluateFormula(sheet, '="   "+1')).toBe('#VALUE!')
    expect(evaluateFormula(sheet, '=""+1')).toBe('#VALUE!')
    expect(evaluateFormula(sheet, '=B1+1')).toBe('#VALUE!')
    expect(evaluateFormula(sheet, '=C1+1')).toBe('#VALUE!')
  })

  it('rounds half-away-from-zero cleanly across positive, negative, and negative-places ROUND calls', () => {
    const sheet = createBlankSheet('s1', 'Sheet 1')
    expect(evaluateFormula(sheet, '=ROUND(35.855, 2)')).toBe('35.86')
    expect(evaluateFormula(sheet, '=ROUND(-2.5, 0)')).toBe('-3')
    expect(evaluateFormula(sheet, '=ROUND(-15, -1)')).toBe('-20')
    expect(evaluateFormula(sheet, '=ROUND(1234.56, -2)')).toBe('1200')
  })

  it('replaces entire single-quoted cross-sheet ref with #REF! when shifted out of bounds', () => {
    let sheet = createBlankSheet('s1', 'Sheet 1')
    sheet = upsertCell(sheet, 'A2', "='Q1'!A1")
    // Fill upward by -1 row shifts A1 -> row -1 -> #REF! (never 'Q1'#REF!)
    const copied = copyRangeCells(sheet, 'A2:A2')
    expect(copied).not.toBeNull()
    const pasted = pasteCellBlock(sheet, 'A1', copied!)
    expect(pasted.cells.A1?.value).toBe('=#REF!')
  })

  it('gives each new workbook a unique stable unit ID preserved across title renames', () => {
    const wbA = createBlankWorkbook('Same Title')
    const wbB = createBlankWorkbook('Same Title')
    const snapA1 = createUniverSnapshot(wbA)
    const snapB = createUniverSnapshot(wbB)
    expect(snapA1.id).not.toBe(snapB.id)
    const renamedA = withFreshUniverSnapshot({
      ...wbA,
      metadata: { ...wbA.metadata, title: 'Renamed Title' },
    })
    expect((renamedA.engine.snapshot as { id: string }).id).toBe(snapA1.id)
  })

  it('hardens univerAdapter snapshot validation against forbidden keys, huge dimensions, and invalid row/col indices', () => {
    const snap = createUniverSnapshot(createBlankWorkbook('Hardened'))
    const s = snap.sheets['sheet-1']
    s.rowCount = 9_999_999
    s.columnCount = 999_999
    s.freeze = { startRow: 500_000, startColumn: 50_000, ySplit: 500_000, xSplit: 50_000 }
    s.rowData = { '0': { h: 99999 }, '-1': { h: 40 }, bad: { h: 40 } } as unknown as typeof s.rowData
    s.columnData = { '0': { w: 50000 }, constructor: { w: 100 } } as unknown as typeof s.columnData
    const restored = documentFromUniverSnapshot(snap)
    const out = restored.workbook.sheets[0]
    expect(out.rowCount).toBeLessThanOrEqual(100_000)
    expect(out.columnCount).toBeLessThanOrEqual(16_384)
    expect(out.frozenRows).toBeLessThanOrEqual(out.rowCount)
    expect(out.frozenColumns).toBeLessThanOrEqual(out.columnCount)
    expect(out.rowHeights).toEqual({ '0': 2000 })
    expect(out.columnWidths).toEqual({ '0': 2000 })
  })

  it('rejects oversized .xlsx zip archives pre-flight before exceljs loads them', async () => {
    const zip = new JSZip()
    for (let i = 0; i < 2005; i += 1) {
      zip.file(`entry-${i}.txt`, 'x')
    }
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    await expect(importXlsxWorkbook(bytes, 'Bomb')).rejects.toThrow(/too many archive entries/i)
  })

  it('hardens CSV control chars, parseCsv caps, non-decimal literals, and __proto__ JSON headers', () => {
    expect(toCsv([['\x00=1+1']])).toBe(`"'\x00=1+1"`)
    expect(() => parseCsv('a,b\n1,2\n3,4\n', { maxRows: 2 })).toThrow(/maximum of 2 rows/i)
    const sheet = sheetFromCsv('Hex,Bin,Dec\n0x10,0b101,-12.5\n', 's1', 'Nums')
    expect(sheet.cells.A2).toMatchObject({ value: '0x10', kind: 'text' })
    expect(sheet.cells.B2).toMatchObject({ value: '0b101', kind: 'text' })
    expect(sheet.cells.C2).toMatchObject({ value: '-12.5', kind: 'number' })

    let protoSheet = createBlankSheet('s1', 'Sheet 1')
    protoSheet = upsertCell(protoSheet, 'A1', '__proto__')
    protoSheet = upsertCell(protoSheet, 'A2', 'pwn')
    const json = jsonFromSheet(protoSheet, 'A1:A2')
    expect(json).toContain('"__proto__": "pwn"')
  })

  it('hardens validateSafeFilePath against encoded traversal and relative paths when required', () => {
    expect(() => validateSafeFilePath('/tmp/%2e%2e/secret.xlsx')).toThrow(/\.\./)
    expect(() => validateSafeFilePath('relative/out.xlsx', { requireAbsolute: true })).toThrow(/absolute/i)
    expect(validateSafeFilePath('/tmp/safe/out.xlsx', { requireAbsolute: true })).toBe('/tmp/safe/out.xlsx')
  })

  it('blocks deleting empty rows/cols inside multi-row/col locked ranges, normalizes comments, quotes cell-like sheet names, and guards draft-only deleteSheet', () => {
    let locked = createBlankSheet('s1', 'Sheet 1')
    locked = lockRange(locked, 'A1:A5')
    locked = lockRange(locked, 'C1:E1')
    // Row index 2 (row 3) is empty inside A1:A5 -> must be blocked
    expect(deleteRow(locked, 2)).toBe(locked)
    // Col index 3 (col D) is empty inside C1:E1 -> must be blocked
    expect(deleteColumn(locked, 3)).toBe(locked)

    expect(normalizeComments({ ' b2 ': 'Note', __proto__: 'bad' })).toEqual({ B2: 'Note' })

    expect(
      renameSheetReferences('=Data!A1+Data!B2', new Map([['Data', 'Q1']])),
    ).toBe("='Q1'!A1+'Q1'!B2")
    expect(
      renameSheetReferences('=Data!A1', new Map([['Data', 'FY2025']])),
    ).toBe("='FY2025'!A1")
    expect(
      renameSheetReferences('=Data!A1', new Map([['Data', 'TRUE']])),
    ).toBe("='TRUE'!A1")

    const doc = createBlankWorkbook('DraftGuard')
    const draftSheet = {
      ...createBlankSheet('draft-1', 'Agent Draft'),
      agentDraft: {
        draftId: 'd1',
        agentRunId: 'run-1',
        summary: 'Draft',
        agentName: 'Bot',
        createdAt: '2026-01-01T00:00:00Z',
        sourceSheetId: 'sheet-1',
        sourceSheetName: 'Sheet 1',
        selectedSourceRange: 'A1:A1',
        prompt: 'Draft prompt',
        qaStatus: 'pending' as const,
      },
    }
    const withDraft = {
      ...doc,
      workbook: { ...doc.workbook, sheets: [...doc.workbook.sheets, draftSheet] },
    }
    // Cannot delete the sole non-draft sheet even when an agentDraft sheet exists
    expect(deleteSheet(withDraft, 'sheet-1')).toBe(withDraft)
  })
})
