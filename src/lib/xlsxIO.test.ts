/**
 * Excel import and export (exceljs). Fixtures are real .xlsx bytes built with
 * exceljs (see xlsxFixtures.ts), so the tests exercise the file format end to
 * end, not mocks.
 */
import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import {
  base64ToBytes,
  bytesToBase64,
  excelSheetNames,
  exportXlsxWorkbook,
  formulaFromExcel,
  formulaToExcel,
  importXlsxWorkbook,
  listZipEntries,
  resolveColor,
  themePalette,
  unsupportedParts,
} from './xlsxIO'
import { buildLossyFixture, buildQuarterlyFixture, prefixedLikeDotNet } from './xlsxFixtures'
import { loadSheetsDocument, serializeSheetsDocument } from './sheetsDocument'
import { createUniverSnapshot, documentFromUniverSnapshot } from './univerAdapter'
import {
  computedCellValue,
  createBlankWorkbook,
  insertRow,
  lockRange,
  setCellStyle,
  upsertCell,
} from './workbookModel'
import type { PureSheetsDocument, WorkbookSheet } from '../types'

async function readWorkbook(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(bytes.slice().buffer as ArrayBuffer)
  return workbook
}

function sheetNamed(document: PureSheetsDocument, name: string): WorkbookSheet {
  const sheet = document.workbook.sheets.find(candidate => candidate.name === name)
  if (!sheet) throw new Error(`No sheet ${name}`)
  return sheet
}

/** A cell as Excel holds it, comparable across files: formulas as text, dates as instants. */
function excelCellShape(cell: ExcelJS.Cell, worksheet: ExcelJS.Worksheet): unknown {
  if (cell.isMerged && cell.master !== cell) return { merged: cell.master.address }
  const value = cell.value
  if (value instanceof Date) return { date: value.getTime() }
  if (value && typeof value === 'object') {
    const formula = value as { formula?: string; sharedFormula?: string }
    if (formula.formula) return { formula: formula.formula }
    if (formula.sharedFormula) {
      // The file Excel wrote shares one formula down a column; ours writes each in full.
      return { formula: worksheet.getCell(cell.address).formula }
    }
    if ('error' in value) return { error: (value as { error: string }).error }
  }
  return { value }
}

describe('Excel import', () => {
  it('keeps every sheet and its name, in order', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    expect(document.metadata.title).toBe('Quarterly')
    expect(document.workbook.sheets.map(sheet => sheet.name)).toEqual(['Sales', 'Costs', 'Q1 plan'])
  })

  it('keeps values and formulas, including formulas across sheets', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const sales = sheetNamed(document, 'Sales')
    expect(sales.cells.A2).toMatchObject({ value: 'North', kind: 'text' })
    expect(sales.cells.B2).toMatchObject({ value: '1200.5', kind: 'number' })
    expect(sales.cells.B3).toMatchObject({ value: '-870.25', kind: 'number' })
    expect(sales.cells.D2?.value).toBe('2026-01-15')
    // Text that looks like a number stays text: an ID's leading zeros survive.
    expect(sales.cells.E2).toMatchObject({ value: '00123', kind: 'text' })
    expect(sales.cells.D9?.value).toBe('TRUE')
    expect(sales.cells.E9?.value).toBe('#N/A')
    expect(sales.cells.B5).toMatchObject({ value: '=SUM(B2:B4)', kind: 'formula' })
    expect(sales.cells.C3).toMatchObject({ value: '=B3/$B$5', kind: 'formula' })
    expect(sales.cells.C4?.value).toBe('=B4/$B$5')
    expect(sales.cells.B7?.value).toBe('=B5-SUM(Costs!B2:B3)')
    expect(sales.cells.B8?.value).toBe("='Q1 plan'!B2*2")
    // Excel's _xlfn. prefix comes off; the formula reads as typed.
    expect(sales.cells.B9?.value).toBe('=XLOOKUP("South",A2:A4,B2:B4)')
    expect(sales.cells.C9?.value).toBe('=IF(A2="North","A1 stays text",0)')
    expect(sheetNamed(document, 'Q1 plan').cells.B3?.value).toBe('=B2-Sales!B5')

    // Cross-sheet formulas calculate against the imported sheets.
    expect(Number(computedCellValue(sales, 'B5', document.workbook))).toBeCloseTo(785.25)
    expect(Number(computedCellValue(sales, 'B7', document.workbook))).toBeCloseTo(535.25)
    expect(computedCellValue(sales, 'B8', document.workbook)).toBe('1800')
    const plan = sheetNamed(document, 'Q1 plan')
    expect(Number(computedCellValue(plan, 'B3', document.workbook))).toBeCloseTo(114.75)
  })

  it('keeps number formats: the family for the toolbar, the exact code for display', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const sales = sheetNamed(document, 'Sales')
    expect(sales.cells.B2?.style).toMatchObject({
      numberFormat: 'number',
      decimals: 1,
      numberPattern: '#,##0.0;[Red]-#,##0.0',
    })
    expect(sales.cells.B5?.style).toMatchObject({ numberFormat: 'currency' })
    expect(sales.cells.B5?.style?.numberPattern).toBeUndefined()
    expect(sales.cells.C2?.style).toMatchObject({ numberFormat: 'percent', decimals: 1 })
    expect(sales.cells.D2?.style).toMatchObject({ numberFormat: 'date', numberPattern: 'd mmm yyyy' })
    expect(sales.cells.E2?.style?.numberFormat).toBe('text')
    expect(sheetNamed(document, 'Q1 plan').cells.B3?.style).toMatchObject({
      numberFormat: 'number',
      numberPattern: '#,##0.00_);(#,##0.00)',
    })
  })

  it('keeps merged cells without copying the value into the hidden parts', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const sales = sheetNamed(document, 'Sales')
    expect(sales.merges).toEqual(expect.arrayContaining(['A11:E11', 'F2:F4']))
    expect(sales.cells.A11?.value).toBe('Quarter one summary')
    for (const key of ['B11', 'C11', 'D11', 'E11', 'F3', 'F4']) expect(sales.cells[key]?.value ?? '').toBe('')
  })

  it('keeps column widths and row heights in pixels', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const sales = sheetNamed(document, 'Sales')
    expect(sales.columnWidths?.['0']).toBe(131) // 18 characters
    expect(sales.columnWidths?.['1']).toBe(93) // 12.5 characters
    expect(sales.columnWidths?.['5']).toBe(215)
    expect(sales.rowHeights?.['0']).toBe(40) // 30pt
    expect(sales.rowHeights?.['10']).toBe(32)
    // Rows at the sheet's default height are left to PureSheets.
    expect(sales.rowHeights?.['1']).toBeUndefined()
  })

  it('keeps bold, italic, fill, font colour, borders and alignment', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const sales = sheetNamed(document, 'Sales')
    expect(sales.cells.A1?.style).toMatchObject({
      bold: true,
      textColor: '#1f3864',
      fillColor: '#dce6f2', // the file's accent 1 (exceljs writes the Office 2007 theme) at 80% lighter
      borderSides: ['bottom'],
      align: 'center',
      verticalAlign: 'middle',
      wrap: 'wrap',
    })
    expect(sales.cells.A1?.style?.border).toBeUndefined()
    // Theme dark 1 is black, Excel's automatic ink: the PureSheets default stands.
    expect(sales.cells.A2?.style).toEqual({ italic: true })
    expect(sales.cells.B5?.style?.border).toBe(true)
    expect(sales.cells.E2?.style).toMatchObject({ align: 'right', verticalAlign: 'top' })
    expect(sales.cells.A11?.style).toMatchObject({ bold: true, fontSize: 14, textColor: '#c00000' })
  })

  it('keeps frozen panes and cell notes', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const sales = sheetNamed(document, 'Sales')
    expect(sales.frozenRows).toBe(1)
    expect(sales.frozenColumns).toBe(1)
    expect(sales.comments?.B2).toBe('Includes the January catch-up order.')
  })

  it('has nothing to report for a workbook it carries whole', async () => {
    const { warnings } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    expect(warnings).toEqual([])
  })

  it('names everything it could not carry, sheet by sheet', async () => {
    const { document, warnings } = await importXlsxWorkbook(await buildLossyFixture(), 'Odds')
    const odds = warnings.find(line => line.startsWith('Odds:')) ?? ''
    for (const phrase of [
      'mixed text formatting',
      'links (kept as their text)',
      'struck-through',
      'cells in Georgia',
      'thick, dashed or double borders',
      'rotated text',
      'hidden rows (shown)',
      'data-validation rules',
      'conditional-formatting rules',
      'pictures',
    ])
      expect(odds).toContain(phrase)
    expect(warnings).toContain('Lookups: hidden sheet (shown as an ordinary sheet).')
    expect(warnings.find(line => line.startsWith('Workbook:'))).toContain('named ranges')
    // What it could not format, it still brought in.
    const sheet = sheetNamed(document, 'Odds')
    expect(sheet.cells.A1?.value).toBe('mixed text')
    expect(sheet.cells.A2?.value).toBe('Pure Science')
    expect(sheet.cells.A7?.value).toBe('Rich link')
    expect(sheet.cells.A8?.value).toBe('hidden row')
    expect(sheetNamed(document, 'Lookups').cells.A1?.value).toBe('kept')
  })

  it('names charts, pivot tables and macros it finds in the file', async () => {
    expect(listZipEntries(await buildQuarterlyFixture())).toEqual(
      expect.arrayContaining(['xl/workbook.xml', 'xl/worksheets/sheet3.xml']),
    )
    expect(
      unsupportedParts([
        'xl/charts/chart1.xml',
        'xl/charts/chart2.xml',
        'xl/drawings/drawing1.xml',
        'xl/pivotTables/pivotTable1.xml',
        'xl/vbaProject.bin',
        'xl/externalLinks/externalLink1.xml',
      ]),
    ).toEqual([
      '2 charts (left out; make them again with Chart)',
      '1 pivot table (its last values came in as plain cells)',
      'macros (left out)',
      'links to 1 other workbook (formulas that use them will show #REF!)',
    ])
  })

  it('reads files written with prefixed XML and absolute part paths, as .NET exporters write them', async () => {
    const plain = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const prefixed = await importXlsxWorkbook(await prefixedLikeDotNet(await buildQuarterlyFixture()), 'Quarterly')
    expect(prefixed.document.workbook.sheets.map(sheet => sheet.name)).toEqual(['Sales', 'Costs', 'Q1 plan'])
    expect(sheetNamed(prefixed.document, 'Sales').cells).toEqual(sheetNamed(plain.document, 'Sales').cells)
  })

  it('says plainly when a file is not a workbook at all', async () => {
    await expect(importXlsxWorkbook(new TextEncoder().encode('not a spreadsheet'), 'Junk')).rejects.toThrow(
      /could not be read as an Excel workbook/,
    )
  })

  it('translates shared formulas instead of replacing them with cached constants', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Shared')
    sheet.getCell('A1').value = 2
    sheet.getCell('A2').value = 4
    sheet.getCell('B1').value = { formula: "LOG10(A1)+$C$1+'Q1 data'!A1", result: 6 }
    sheet.getCell('B2').value = { sharedFormula: 'B1', result: 12 }
    const { document } = await importXlsxWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()), 'Shared')
    // The sheet name 'Q1 data' is quoted text, not a reference to cell Q1.
    expect(document.workbook.sheets[0].cells.B2).toMatchObject({
      value: "=LOG10(A2)+$C$1+'Q1 data'!A2",
      kind: 'formula',
    })
  })

  it('preserves the time portion of imported Excel datetimes', async () => {
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('Dates')
    sheet.getCell('A1').value = new Date('2026-10-25T14:35:42.000Z')
    sheet.getCell('A1').numFmt = 'yyyy-mm-dd hh:mm:ss'
    const { document } = await importXlsxWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()), 'Dates')
    expect(document.workbook.sheets[0].cells.A1?.value).toBe('2026-10-25T14:35:42.000Z')
  })

  it('imports only the selected byte slice rather than unrelated bytes in its backing buffer', async () => {
    const bytes = await buildQuarterlyFixture()
    const padded = new Uint8Array(bytes.length + 20)
    padded.fill(255)
    padded.set(bytes, 10)
    const result = await importXlsxWorkbook(padded.subarray(10, 10 + bytes.length), 'Slice')
    expect(result.document.workbook.sheets[0].name).toBe('Sales')
  })
})

describe('Excel round trip (xlsx → PureSheets → xlsx)', () => {
  it('keeps every value and formula identical, cell by cell', async () => {
    const original = await buildQuarterlyFixture()
    const { document } = await importXlsxWorkbook(original, 'Quarterly')
    const { bytes, warnings } = await exportXlsxWorkbook(document)
    expect(warnings).toEqual([])
    const before = await readWorkbook(original)
    const after = await readWorkbook(bytes)

    expect(after.worksheets.map(sheet => sheet.name)).toEqual(before.worksheets.map(sheet => sheet.name))
    let compared = 0
    for (const sheetBefore of before.worksheets) {
      const sheetAfter = after.getWorksheet(sheetBefore.name)!
      const addresses = new Set<string>()
      for (const sheet of [sheetBefore, sheetAfter])
        sheet.eachRow({ includeEmpty: false }, row =>
          row.eachCell({ includeEmpty: false }, cell => {
            if (cell.value !== null && cell.value !== undefined) addresses.add(cell.address)
          }),
        )
      for (const address of addresses) {
        const cellBefore = sheetBefore.getCell(address)
        const cellAfter = sheetAfter.getCell(address)
        expect(
          excelCellShape(cellAfter, sheetAfter),
          `${sheetBefore.name}!${address}`,
        ).toEqual(excelCellShape(cellBefore, sheetBefore))
        compared += 1
      }
    }
    expect(compared).toBeGreaterThan(45)
  })

  it('writes back number formats, merges, sizes, styles, frozen panes and notes', async () => {
    const original = await buildQuarterlyFixture()
    const { document } = await importXlsxWorkbook(original, 'Quarterly')
    const before = await readWorkbook(original)
    const after = await readWorkbook((await exportXlsxWorkbook(document)).bytes)
    const salesBefore = before.getWorksheet('Sales')!
    const sales = after.getWorksheet('Sales')!

    for (const key of ['B2', 'B5', 'C3', 'D2', 'E2']) expect(sales.getCell(key).numFmt, key).toBe(salesBefore.getCell(key).numFmt)
    expect(after.getWorksheet('Q1 plan')!.getCell('B3').numFmt).toBe('#,##0.00_);(#,##0.00)')
    expect((sales.model as { merges: string[] }).merges.sort()).toEqual(
      (salesBefore.model as { merges: string[] }).merges.sort(),
    )
    for (const column of [1, 2, 6])
      expect(sales.getColumn(column).width).toBeCloseTo(salesBefore.getColumn(column).width!, 0)
    expect(sales.getRow(1).height).toBe(30)
    expect(sales.getRow(11).height).toBe(24)
    expect(sales.views[0]).toMatchObject({ state: 'frozen', xSplit: 1, ySplit: 1 })
    expect(sales.getCell('B2').note).toBe('Includes the January catch-up order.')

    const header = sales.getCell('A1')
    expect(header.font).toMatchObject({ bold: true, color: { argb: 'FF1F3864' } })
    expect(header.fill).toMatchObject({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F2' } })
    expect(Object.keys(header.border ?? {}).filter(side => header.border?.[side as 'top']?.style)).toEqual(['bottom'])
    expect(header.alignment).toMatchObject({ horizontal: 'center', vertical: 'middle', wrapText: true })
    expect(sales.getCell('A2').font).toMatchObject({ italic: true })
    expect(sales.getCell('B5').border).toMatchObject({
      top: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
    })
    expect(sales.getCell('E2').alignment).toMatchObject({ horizontal: 'right', vertical: 'top' })
    expect(sales.getCell('A11').font).toMatchObject({ bold: true, size: 14 })
  })

  it('puts the _xlfn. prefix back so Excel recognises newer functions', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const after = await readWorkbook((await exportXlsxWorkbook(document)).bytes)
    expect((after.getWorksheet('Sales')!.getCell('B9').value as { formula: string }).formula).toBe(
      '_xlfn.XLOOKUP("South",A2:A4,B2:B4)',
    )
  })

  it('is stable: a second round trip gives the same workbook', async () => {
    const first = (await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')).document
    const second = (await importXlsxWorkbook((await exportXlsxWorkbook(first)).bytes, 'Quarterly')).document
    for (const sheet of first.workbook.sheets) {
      const again = sheetNamed(second, sheet.name)
      expect(again.cells).toEqual(sheet.cells)
      expect(again.merges).toEqual(sheet.merges)
      expect(again.columnWidths).toEqual(sheet.columnWidths)
      expect(again.rowHeights).toEqual(sheet.rowHeights)
      expect(again.comments).toEqual(sheet.comments)
      expect([again.frozenRows, again.frozenColumns]).toEqual([sheet.frozenRows, sheet.frozenColumns])
    }
  })

  it('survives the native .sheets save and the grid editor between import and export', async () => {
    const { document } = await importXlsxWorkbook(await buildQuarterlyFixture(), 'Quarterly')
    const saved = loadSheetsDocument(serializeSheetsDocument(document))
    const edited = documentFromUniverSnapshot(createUniverSnapshot(saved), saved)
    for (const sheet of document.workbook.sheets) {
      const kept = sheetNamed(edited, sheet.name)
      expect(kept.merges, sheet.name).toEqual(sheet.merges)
      expect(kept.comments, sheet.name).toEqual(sheet.comments)
      for (const [key, cell] of Object.entries(sheet.cells)) {
        expect(kept.cells[key]?.value, `${sheet.name}!${key}`).toBe(cell.value)
        expect(kept.cells[key]?.style, `${sheet.name}!${key}`).toEqual(cell.style)
      }
    }
  })
})

describe('Excel export', () => {
  function buildDocument(): PureSheetsDocument {
    const base = createBlankWorkbook('Export me')
    let sheet = base.workbook.sheets[0]
    sheet = upsertCell(sheet, 'A1', 'Region')
    sheet = setCellStyle(sheet, 'A1', { bold: true, fillColor: '#eef1eb' })
    sheet = upsertCell(sheet, 'A2', 'North')
    sheet = upsertCell(sheet, 'B2', '1200')
    sheet = setCellStyle(sheet, 'B2', { numberFormat: 'currency', decimals: 0 })
    sheet = upsertCell(sheet, 'B3', '=SUM(B2:B2)')
    sheet = { ...sheet, frozenRows: 1, columnWidths: { '0': 150 } }
    return { ...base, workbook: { ...base.workbook, sheets: [sheet] } }
  }

  it('exports values, formulas, styles, and freeze state that re-import cleanly', async () => {
    const { bytes, warnings } = await exportXlsxWorkbook(buildDocument())
    expect(bytes.byteLength).toBeGreaterThan(1000)
    expect(warnings).toEqual([])
    const { document } = await importXlsxWorkbook(bytes, 'Reimported')
    const sheet = document.workbook.sheets[0]
    expect(sheet.name).toBe('Sheet 1')
    expect(sheet.cells.A1).toMatchObject({ value: 'Region', style: { bold: true, fillColor: '#eef1eb' } })
    expect(sheet.cells.B2).toMatchObject({ value: '1200', kind: 'number' })
    expect(sheet.cells.B2?.style).toMatchObject({ numberFormat: 'currency', decimals: 0 })
    expect(sheet.cells.B3?.value).toBe('=SUM(B2:B2)')
    expect(sheet.frozenRows).toBe(1)
    expect(sheet.columnWidths?.['0']).toBe(150)
  })

  it('writes the cached result of each formula', async () => {
    const { bytes } = await exportXlsxWorkbook(buildDocument())
    const workbook = await readWorkbook(bytes)
    expect(workbook.getWorksheet('Sheet 1')!.getCell('B3').value).toMatchObject({ formula: 'SUM(B2:B2)', result: 1200 })
  })

  it('names what Excel will not get', async () => {
    const base = buildDocument()
    const sheet: WorkbookSheet = {
      ...base.workbook.sheets[0],
      charts: [{ id: 'c1', title: 'Revenue', type: 'bar', range: 'A1:B2', labels: ['North'], values: [1200] }],
      validations: [{ id: 'v1', range: 'C1:C3', kind: 'list', options: ['Yes', 'No'] }],
    }
    const { warnings } = await exportXlsxWorkbook({ ...base, workbook: { ...base.workbook, sheets: [sheet] } })
    expect(warnings).toEqual(['Sheet 1: charts (Excel will not get them); data-validation rules.'])
  })

  it('renames sheets Excel would refuse, and the formulas that point at them', async () => {
    const base = createBlankWorkbook('Names')
    const first = { ...upsertCell(base.workbook.sheets[0], 'A1', '7'), name: 'Q1/Q2' }
    const second = { ...upsertCell(base.workbook.sheets[0], 'A1', "='Q1/Q2'!A1*2"), id: 'sheet-2', name: 'Summary' }
    const { bytes, warnings } = await exportXlsxWorkbook({ ...base, workbook: { ...base.workbook, sheets: [first, second] } })
    const workbook = await readWorkbook(bytes)
    expect(workbook.worksheets.map(sheet => sheet.name)).toEqual(['Q1-Q2', 'Summary'])
    expect(workbook.getWorksheet('Summary')!.getCell('A1').value).toMatchObject({ formula: "'Q1-Q2'!A1*2", result: 14 })
    expect(warnings).toEqual([
      'Workbook: sheet "Q1/Q2" is named "Q1-Q2" (Excel does not allow that name; formulas follow it).',
    ])
    expect([...excelSheetNames(['History', 'A'.repeat(40), 'a'.repeat(31), "'quoted'"]).values()]).toEqual([
      'History sheet',
      'A'.repeat(31),
      `${'a'.repeat(29)} 2`,
      'quoted',
    ])
  })

  it('excludes agent draft sheets from exports', async () => {
    const base = buildDocument()
    const draft = {
      ...base.workbook.sheets[0],
      id: 'draft-1',
      name: 'Sheet 1 - Agent draft',
      agentDraft: {
        sourceSheetId: base.workbook.sheets[0].id,
        sourceSheetName: 'Sheet 1',
        agentRunId: 'run',
        agentName: 'Assistant',
        createdAt: new Date().toISOString(),
        selectedSourceRange: 'A1:A1',
        prompt: 'x',
        qaStatus: 'pending' as const,
      },
    }
    const withDraft: PureSheetsDocument = {
      ...base,
      workbook: { ...base.workbook, sheets: [base.workbook.sheets[0], draft] },
    }
    const { document } = await importXlsxWorkbook((await exportXlsxWorkbook(withDraft)).bytes, 'Check')
    expect(document.workbook.sheets.map(sheet => sheet.name)).toEqual(['Sheet 1'])
  })

  it('locked source cells still export their values (locks are app metadata)', async () => {
    const base = buildDocument()
    const locked = {
      ...base,
      workbook: { ...base.workbook, sheets: [lockRange(base.workbook.sheets[0], 'A1:B3')] },
    }
    const { bytes, warnings } = await exportXlsxWorkbook(locked)
    expect(warnings[0]).toContain('locked ranges')
    const { document } = await importXlsxWorkbook(bytes, 'Locked check')
    expect(document.workbook.sheets[0].cells.A1?.value).toBe('Region')
  })

  it('exports unsafe-integer identifiers as text instead of rounding them', async () => {
    const base = createBlankWorkbook('Identifiers')
    const sheet = upsertCell(base.workbook.sheets[0], 'A1', '9007199254740993')
    const { bytes } = await exportXlsxWorkbook({ ...base, workbook: { ...base.workbook, sheets: [sheet] } })
    const { document } = await importXlsxWorkbook(bytes, 'Identifiers')
    expect(document.workbook.sheets[0].cells.A1?.value).toBe('9007199254740993')
  })
})

describe('model support for Excel features', () => {
  it('moves merges with inserted rows', () => {
    const base = createBlankWorkbook('Merges').workbook.sheets[0]
    const sheet = insertRow({ ...base, merges: ['A2:C2', 'D1:D4'] }, 1)
    expect(sheet.merges).toEqual(['A3:C3', 'D1:D5'])
  })

  it('drops an imported format code when the user picks a format', () => {
    const base = createBlankWorkbook('Formats').workbook.sheets[0]
    const imported = setCellStyle(base, 'A1', {
      numberFormat: 'number',
      decimals: 1,
      numberPattern: '#,##0.0;[Red]-#,##0.0',
      borderSides: ['bottom'],
    })
    const chosen = setCellStyle(setCellStyle(imported, 'A1', { numberFormat: 'currency' }), 'A1', { border: true })
    expect(chosen.cells.A1?.style?.numberPattern).toBeUndefined()
    expect(chosen.cells.A1?.style?.borderSides).toBeUndefined()
    expect(setCellStyle(imported, 'A1', { bold: true }).cells.A1?.style?.numberPattern).toBe('#,##0.0;[Red]-#,##0.0')
  })
})

describe('formula prefixes', () => {
  it('strips and restores _xlfn. only on known newer functions, outside text', () => {
    expect(formulaFromExcel('_xlfn._xlws.FILTER(A1:A9,B1:B9>0)')).toBe('FILTER(A1:A9,B1:B9>0)')
    expect(formulaToExcel('FILTER(A1:A9,B1:B9>0)')).toBe('_xlfn._xlws.FILTER(A1:A9,B1:B9>0)')
    expect(formulaToExcel('IFS(A1>1,"TEXTJOIN(",TRUE,SUM(A1:A3))')).toBe('_xlfn.IFS(A1>1,"TEXTJOIN(",TRUE,SUM(A1:A3))')
    expect(formulaToExcel("'XLOOKUP(x'!A1+DAYS360(A1,A2)")).toBe("'XLOOKUP(x'!A1+DAYS360(A1,A2)")
    // An unknown prefixed name stays as Excel wrote it.
    expect(formulaFromExcel('_xlfn.NEWTHING(A1)')).toBe('_xlfn.NEWTHING(A1)')
  })
})

describe('colours', () => {
  it('reads theme slots from the file and applies tints', () => {
    const palette = themePalette(
      '<a:theme><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:srgbClr val="FEFEFE"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme></a:themeElements></a:theme>',
    )
    expect(palette.slice(0, 5)).toEqual(['#fefefe', '#111111', '#eeeeee', '#222222', '#4472c4'])
    expect(resolveColor({ theme: 4, tint: 0.7999816888943144 }, palette)).toBe('#dae3f3')
    expect(resolveColor({ theme: 4, tint: -0.249977111117893 }, palette)).toBe('#2f5597')
    expect(resolveColor({ argb: 'FF00B050' }, palette)).toBe('#00b050')
    expect(resolveColor({ indexed: 10 } as never, palette)).toBe('#ff0000')
  })
})

describe('base64 helpers', () => {
  it('round-trips bytes', () => {
    const bytes = new Uint8Array([0, 1, 2, 250, 255, 128, 64])
    expect(Array.from(base64ToBytes(bytesToBase64(bytes)))).toEqual(Array.from(bytes))
  })
})
