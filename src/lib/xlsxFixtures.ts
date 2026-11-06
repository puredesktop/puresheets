import ExcelJS from 'exceljs'
import JSZip from 'jszip'

/**
 * Fixture workbooks for the Excel import/export tests, written with exceljs so
 * the tests read real .xlsx bytes. They carry what Excel itself writes: theme
 * colours with tints, `_xlfn.` prefixes, shared formulas, sheet names that
 * need quoting.
 */

/** Three sheets with cross-sheet formulas, merges, number formats and styles: everything import keeps. */
export async function buildQuarterlyFixture(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  const sales = workbook.addWorksheet('Sales')
  const costs = workbook.addWorksheet('Costs')
  const plan = workbook.addWorksheet('Q1 plan')

  sales.getCell('A1').value = 'Region'
  sales.getCell('B1').value = 'Revenue'
  sales.getCell('C1').value = 'Share'
  sales.getCell('D1').value = 'Closed'
  sales.getCell('E1').value = 'Code'
  for (const key of ['A1', 'B1', 'C1', 'D1', 'E1']) {
    const cell = sales.getCell(key)
    cell.font = { bold: true, color: { argb: 'FF1F3864' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 4, tint: 0.7999816888943144 } as Partial<ExcelJS.Color> }
    cell.border = { bottom: { style: 'thin' } }
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
  }
  const rows: Array<[string, number, Date, string]> = [
    ['North', 1200.5, new Date('2026-01-15T00:00:00.000Z'), '00123'],
    ['South', -870.25, new Date('2026-02-03T00:00:00.000Z'), '00456'],
    ['East', 455, new Date('2026-03-31T00:00:00.000Z'), '00789'],
  ]
  rows.forEach(([region, revenue, closed, code], index) => {
    const row = index + 2
    sales.getCell(`A${row}`).value = region
    sales.getCell(`A${row}`).font = { italic: true, color: { theme: 1 } }
    sales.getCell(`B${row}`).value = revenue
    sales.getCell(`B${row}`).numFmt = '#,##0.0;[Red]-#,##0.0'
    sales.getCell(`D${row}`).value = closed
    sales.getCell(`D${row}`).numFmt = 'd mmm yyyy'
    sales.getCell(`E${row}`).value = code
    sales.getCell(`E${row}`).numFmt = '@'
    sales.getCell(`E${row}`).alignment = { horizontal: 'right', vertical: 'top' }
  })
  // Share of total: a shared formula, as Excel writes a filled-down column.
  sales.getCell('C2').value = { formula: 'B2/$B$5', result: 1 }
  sales.getCell('C3').value = { sharedFormula: 'C2', result: 1 }
  sales.getCell('C4').value = { sharedFormula: 'C2', result: 1 }
  for (const key of ['C2', 'C3', 'C4']) sales.getCell(key).numFmt = '0.0%'
  sales.getCell('A5').value = 'Total'
  sales.getCell('A5').font = { bold: true }
  sales.getCell('B5').value = { formula: 'SUM(B2:B4)', result: 785.25 }
  sales.getCell('B5').numFmt = '$#,##0.00'
  sales.getCell('B5').border = {
    top: { style: 'thin' },
    bottom: { style: 'thin' },
    left: { style: 'thin' },
    right: { style: 'thin' },
  }
  sales.getCell('A7').value = 'Net after costs'
  sales.getCell('B7').value = { formula: 'B5-SUM(Costs!B2:B3)', result: 535.25 }
  sales.getCell('A8').value = 'Plan target'
  sales.getCell('B8').value = { formula: "'Q1 plan'!B2*2", result: 1800 }
  sales.getCell('A9').value = 'Lookup'
  sales.getCell('B9').value = { formula: '_xlfn.XLOOKUP("South",A2:A4,B2:B4)', result: -870.25 }
  sales.getCell('C9').value = { formula: 'IF(A2="North","A1 stays text",0)', result: 'A1 stays text' }
  sales.getCell('D9').value = true
  sales.getCell('E9').value = { error: '#N/A' }
  sales.getCell('A11').value = 'Quarter one summary'
  sales.getCell('A11').font = { bold: true, size: 14, color: { argb: 'FFC00000' } }
  sales.getCell('A11').alignment = { horizontal: 'center', vertical: 'middle' }
  sales.mergeCells('A11:E11')
  sales.getCell('F2').value = 'Notes'
  sales.getCell('F2').alignment = { vertical: 'top', wrapText: true }
  sales.mergeCells('F2:F4')
  sales.getCell('B2').note = 'Includes the January catch-up order.'
  sales.getColumn(1).width = 18
  sales.getColumn(2).width = 12.5
  sales.getColumn(6).width = 30
  sales.getRow(1).height = 30
  sales.getRow(11).height = 24
  sales.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }]

  costs.getCell('A1').value = 'Item'
  costs.getCell('B1').value = 'Cost'
  costs.getCell('A2').value = 'Rent'
  costs.getCell('B2').value = 200
  costs.getCell('A3').value = 'Travel'
  costs.getCell('B3').value = 50
  costs.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 250 }
  costs.getCell('B4').font = { bold: true }
  costs.getCell('C4').value = { formula: 'B4/Sales!B5', result: 0.318 }
  costs.getCell('C4').numFmt = '0.00%'

  plan.getCell('A1').value = 'Metric'
  plan.getCell('B1').value = 'Target'
  plan.getCell('A2').value = 'Revenue'
  plan.getCell('B2').value = 900
  plan.getCell('B3').value = { formula: 'B2-Sales!B5', result: 114.75 }
  plan.getCell('B3').numFmt = '#,##0.00_);(#,##0.00)'

  return new Uint8Array((await workbook.xlsx.writeBuffer()) as ArrayBuffer)
}

const ONE_PIXEL_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/** One sheet of things PureSheets cannot hold, each of which the notice must name. */
export async function buildLossyFixture(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Odds')
  sheet.getCell('A1').value = { richText: [{ text: 'mixed ', font: { bold: true } }, { text: 'text' }] }
  sheet.getCell('A2').value = { text: 'Pure Science', hyperlink: 'https://pure.science' }
  sheet.getCell('A3').value = 'struck'
  sheet.getCell('A3').font = { strike: true }
  sheet.getCell('A4').value = 'serif'
  sheet.getCell('A4').font = { name: 'Georgia' }
  sheet.getCell('A5').value = 'thick'
  sheet.getCell('A5').border = { bottom: { style: 'thick' } }
  sheet.getCell('A6').value = 'turned'
  sheet.getCell('A6').alignment = { textRotation: 45 }
  sheet.getCell('A7').value = {
    text: { richText: [{ text: 'Rich ' }, { text: 'link', font: { bold: true } }] },
    hyperlink: 'https://pure.science',
  } as unknown as ExcelJS.CellHyperlinkValue
  sheet.getRow(8).hidden = true
  sheet.getCell('A8').value = 'hidden row'
  sheet.getCell('B1').dataValidation = { type: 'list', allowBlank: true, formulae: ['"Yes,No"'] }
  sheet.addConditionalFormatting({
    ref: 'C1:C5',
    rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [5], priority: 1, style: { font: { bold: true } } }],
  })
  workbook.definedNames.add("'Odds'!$A$1:$A$3", 'Labels')
  const image = workbook.addImage({ base64: ONE_PIXEL_PNG, extension: 'png' })
  sheet.addImage(image, 'D2:E4')
  const hidden = workbook.addWorksheet('Lookups', { state: 'hidden' })
  hidden.getCell('A1').value = 'kept'
  return new Uint8Array((await workbook.xlsx.writeBuffer()) as ArrayBuffer)
}

/**
 * The same workbook as .NET exporters write it: every element prefixed
 * (`<x:workbook>`) and relationships pointing at parts by absolute path.
 */
export async function prefixedLikeDotNet(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  for (const name of Object.keys(zip.files)) {
    if (name === 'xl/workbook.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(name)) {
      const xml = await zip.file(name)!.async('string')
      zip.file(
        name,
        xml
          .replace(/<(\/?)([A-Za-z]+)(?=[\s/>])/g, '<$1x:$2')
          .replace(`xmlns="${main}"`, `xmlns:x="${main}"`),
      )
    }
    if (name === 'xl/_rels/workbook.xml.rels') {
      const xml = await zip.file(name)!.async('string')
      zip.file(name, xml.replace(/Target="(?!\/)([^"]+)"/g, 'Target="/xl/$1"'))
    }
  }
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }))
}
