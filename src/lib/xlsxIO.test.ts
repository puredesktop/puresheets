/**
 * Phase S4 — XLSX import mapping and export round-trip (exceljs).
 * Import fixtures are built with exceljs itself, so the tests exercise real
 * xlsx bytes end to end, not mocks.
 */
import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import {
  base64ToBytes,
  bytesToBase64,
  exportXlsxWorkbook,
  importXlsxWorkbook,
} from './xlsxIO'
import { loadSheetsDocument, serializeSheetsDocument } from './sheetsDocument'
import { createBlankWorkbook, upsertCell, setCellStyle, lockRange } from './workbookModel'
import type { PureSheetsDocument } from '../types'

async function buildFixtureXlsx(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Budget')
  sheet.getCell('A1').value = 'Region'
  sheet.getCell('A1').font = { bold: true, color: { argb: 'FF982D25' } }
  sheet.getCell('B1').value = 'Revenue'
  sheet.getCell('B1').alignment = { horizontal: 'right' }
  sheet.getCell('A2').value = 'North'
  sheet.getCell('B2').value = 1200
  sheet.getCell('B2').numFmt = '$#,##0.00'
  sheet.getCell('A3').value = 'South'
  sheet.getCell('B3').value = 870
  sheet.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 2070 }
  sheet.getCell('C1').value = {
    richText: [{ text: 'rich ' }, { text: 'text' }],
  }
  sheet.getCell('C2').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE5F3DF' },
  }
  sheet.getCell('C2').value = 'filled'
  sheet.mergeCells('A5:B5')
  sheet.getCell('A5').value = 'merged banner'
  sheet.getColumn(1).width = 20
  sheet.views = [{ state: 'frozen', ySplit: 1, xSplit: 0 }]
  const second = workbook.addWorksheet('Notes')
  second.getCell('A1').value = 'hello'
  const buffer = await workbook.xlsx.writeBuffer()
  return new Uint8Array(buffer as ArrayBuffer)
}

describe('XLSX import', () => {
  it('maps sheets, values, formulas, styles, sizing, and freeze state', async () => {
    const bytes = await buildFixtureXlsx()
    const { document, warnings } = await importXlsxWorkbook(bytes, 'Budget import')

    expect(document.metadata.title).toBe('Budget import')
    expect(document.workbook.sheets.map(sheet => sheet.name)).toEqual([
      'Budget',
      'Notes',
    ])
    const budget = document.workbook.sheets[0]
    expect(budget.cells.A1).toMatchObject({
      value: 'Region',
      kind: 'text',
      style: { bold: true, textColor: '#982d25' },
    })
    expect(budget.cells.B1?.style?.align).toBe('right')
    expect(budget.cells.B2).toMatchObject({ value: '1200', kind: 'number' })
    expect(budget.cells.B2?.style?.numberFormat).toBe('currency')
    expect(budget.cells.B4).toMatchObject({
      value: '=SUM(B2:B3)',
      kind: 'formula',
    })
    expect(budget.cells.C1?.value).toBe('rich text')
    expect(budget.cells.C2?.style?.fillColor).toBe('#e5f3df')
    expect(budget.cells.A5?.value).toBe('merged banner')
    expect(budget.columnWidths?.['0']).toBeGreaterThan(100)
    expect(budget.frozenRows).toBe(1)
    expect(document.workbook.sheets[1].cells.A1?.value).toBe('hello')

    // Warnings are explicit and user-readable.
    expect(warnings.some(warning => warning.includes('merged cell'))).toBe(true)
    expect(
      warnings.some(warning => warning.includes('rich-text')),
    ).toBe(true)
    expect(
      warnings.some(warning =>
        warning.includes('charts, pivot tables, macros'),
      ),
    ).toBe(true)
  })

  it('imported documents serialize through the native .sheets format', async () => {
    const bytes = await buildFixtureXlsx()
    const { document } = await importXlsxWorkbook(bytes, 'Native check')
    const reloaded = loadSheetsDocument(serializeSheetsDocument(document))
    expect(reloaded.workbook.sheets[0].cells.B4?.value).toBe('=SUM(B2:B3)')
    expect(reloaded.workbook.sheets[0].cells.A1?.style?.bold).toBe(true)
    expect(reloaded.workbook.sheets[0].frozenRows).toBe(1)
  })
})

describe('XLSX export', () => {
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
    return {
      ...base,
      workbook: { ...base.workbook, sheets: [sheet] },
    }
  }

  it('exports values, formulas, styles, and freeze state that re-import cleanly', async () => {
    const bytes = await exportXlsxWorkbook(buildDocument())
    expect(bytes.byteLength).toBeGreaterThan(1000)
    // Round-trip through our own importer.
    const { document } = await importXlsxWorkbook(bytes, 'Reimported')
    const sheet = document.workbook.sheets[0]
    expect(sheet.name).toBe('Sheet 1')
    expect(sheet.cells.A1).toMatchObject({
      value: 'Region',
      style: { bold: true, fillColor: '#eef1eb' },
    })
    expect(sheet.cells.B2).toMatchObject({ value: '1200', kind: 'number' })
    expect(sheet.cells.B2?.style?.numberFormat).toBe('currency')
    expect(sheet.cells.B2?.style?.decimals).toBe(0)
    expect(sheet.cells.B3?.value).toBe('=SUM(B2:B2)')
    expect(sheet.frozenRows).toBe(1)
    expect(sheet.columnWidths?.['0']).toBeGreaterThan(100)
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
    const bytes = await exportXlsxWorkbook(withDraft)
    const { document } = await importXlsxWorkbook(bytes, 'Check')
    expect(document.workbook.sheets.map(sheet => sheet.name)).toEqual([
      'Sheet 1',
    ])
  })

  it('locked source cells still export their values (locks are app metadata)', async () => {
    const base = buildDocument()
    const locked = {
      ...base,
      workbook: {
        ...base.workbook,
        sheets: [lockRange(base.workbook.sheets[0], 'A1:B3')],
      },
    }
    const bytes = await exportXlsxWorkbook(locked)
    const { document } = await importXlsxWorkbook(bytes, 'Locked check')
    expect(document.workbook.sheets[0].cells.A1?.value).toBe('Region')
  })
})

describe('base64 helpers', () => {
  it('round-trips bytes', () => {
    const bytes = new Uint8Array([0, 1, 2, 250, 255, 128, 64])
    expect(Array.from(base64ToBytes(bytesToBase64(bytes)))).toEqual(
      Array.from(bytes),
    )
  })
})
