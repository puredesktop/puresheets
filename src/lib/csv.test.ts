/**
 * Phase S4 — CSV import/export: RFC 4180 parsing (quotes, escaped quotes,
 * embedded delimiters/newlines), delimiter auto-detection, serialization
 * quoting, and the sheet round-trip through the document model.
 */
import { describe, expect, it } from 'vitest'
import {
  csvFromSheet,
  detectCsvDelimiter,
  isDelimitedTextPath,
  parseCsv,
  sheetFromCsv,
  toCsv,
  workbookFromCsv,
} from './csv'
import { createBlankWorkbook, upsertCell } from './workbookModel'

describe('parseCsv', () => {
  it('parses plain rows with CRLF and LF endings', () => {
    expect(parseCsv('a,b\r\nc,d\ne,f')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
      ['e', 'f'],
    ])
  })

  it('handles quoted fields with embedded delimiters, quotes, and newlines', () => {
    const text = 'name,note\n"Smith, Jane","said ""hi""\nthen left"\nplain,ok'
    expect(parseCsv(text)).toEqual([
      ['name', 'note'],
      ['Smith, Jane', 'said "hi"\nthen left'],
      ['plain', 'ok'],
    ])
  })

  it('keeps empty fields and drops one trailing newline row', () => {
    expect(parseCsv('a,,c\n,,\n')).toEqual([
      ['a', '', 'c'],
      ['', '', ''],
    ])
  })

  it('auto-detects semicolon and tab delimiters', () => {
    expect(detectCsvDelimiter('a;b;c\nd;e;f')).toBe(';')
    expect(detectCsvDelimiter('a\tb\tc')).toBe('\t')
    expect(detectCsvDelimiter('a,b,c')).toBe(',')
    expect(parseCsv('x;y\n1;2')).toEqual([
      ['x', 'y'],
      ['1', '2'],
    ])
  })

  it('does not split on delimiters inside quotes when detecting', () => {
    const text = '"a;long;text",b\nc,d'
    expect(parseCsv(text, { delimiter: ',' })).toEqual([
      ['a;long;text', 'b'],
      ['c', 'd'],
    ])
  })
})

describe('toCsv', () => {
  it('quotes only when needed and escapes quotes', () => {
    expect(
      toCsv([
        ['plain', 'with,comma', 'with"quote', 'with\nnewline'],
        ['1', '2', '3', '4'],
      ]),
    ).toBe('plain,"with,comma","with""quote","with\nnewline"\n1,2,3,4')
  })

  it('round-trips through parseCsv', () => {
    const rows = [
      ['a', 'b,c', 'd"e'],
      ['multi\nline', '', 'last'],
    ]
    expect(parseCsv(toCsv(rows), { delimiter: ',' })).toEqual(rows)
  })
})

describe('sheetFromCsv / csvFromSheet', () => {
  it('imports CSV into a sheet through the normal model', () => {
    const sheet = sheetFromCsv('Region,Revenue\nNorth,1200\nSouth,870', 's1', 'Imported')
    expect(sheet.name).toBe('Imported')
    expect(sheet.cells.A1).toEqual({ value: 'Region', kind: 'text', style: undefined })
    expect(sheet.cells.B2).toEqual({ value: '1200', kind: 'number', style: undefined })
    expect(sheet.cells.A3?.value).toBe('South')
  })

  it('grows the sheet beyond the default grid for large imports', () => {
    const rows = Array.from({ length: 60 }, (_, index) => `r${index}`).join('\n')
    const sheet = sheetFromCsv(rows, 's1', 'Tall')
    expect(sheet.rowCount).toBe(60)
    expect(sheet.cells.A60?.value).toBe('r59')
  })

  it('exports computed formula values, quoting where needed', () => {
    let sheet = createBlankWorkbook('Export').workbook.sheets[0]
    sheet = upsertCell(sheet, 'A1', '2')
    sheet = upsertCell(sheet, 'A2', '3')
    sheet = upsertCell(sheet, 'B1', '=SUM(A1:A2)')
    sheet = upsertCell(sheet, 'B2', 'hello, world')
    expect(csvFromSheet(sheet)).toBe('2,5\n3,"hello, world"\n')
  })

  it('exports a specific range when given', () => {
    let sheet = createBlankWorkbook('Export range').workbook.sheets[0]
    sheet = upsertCell(sheet, 'A1', 'skip')
    sheet = upsertCell(sheet, 'B2', 'keep')
    sheet = upsertCell(sheet, 'C2', 'also')
    expect(csvFromSheet(sheet, 'B2:C2')).toBe('keep,also\n')
  })

  it('CSV import → export round-trips values', () => {
    const original = 'a,b\n"1,5",2\nx,"y""z"'
    const sheet = sheetFromCsv(original, 's1', 'RT')
    expect(csvFromSheet(sheet, 'A1:B3')).toBe(`${original}\n`)
  })
})

describe('workbookFromCsv', () => {
  it('opens a CSV as a workbook of one sheet named after the file', () => {
    const doc = workbookFromCsv('Item,Cost\nTea,3\n', 'Budget.csv')
    expect(doc.app).toBe('PureSheets')
    expect(doc.metadata.title).toBe('Budget')
    expect(doc.workbook.sheets).toHaveLength(1)
    expect(doc.workbook.sheets[0].name).toBe('Budget')
    expect(doc.workbook.activeSheetId).toBe(doc.workbook.sheets[0].id)
    expect(doc.workbook.sheets[0].cells.A1?.value).toBe('Item')
    expect(doc.workbook.sheets[0].cells.B2).toEqual({ value: '3', kind: 'number', style: undefined })
  })

  it('tells delimited text paths from workbook files', () => {
    expect(isDelimitedTextPath('/docs/Budget.csv')).toBe(true)
    expect(isDelimitedTextPath('/docs/Budget.TSV')).toBe(true)
    expect(isDelimitedTextPath('/docs/Budget.sheets')).toBe(false)
    expect(isDelimitedTextPath('/docs/Budget.sheets.html')).toBe(false)
  })

  it('always uses tabs for .tsv even when fields contain more commas', () => {
    const doc = workbookFromCsv('name\tnote\nalpha\t1,2,3\n', 'demo.tsv')
    expect(doc.workbook.sheets[0].cells.A2?.value).toBe('alpha')
    expect(doc.workbook.sheets[0].cells.B2?.value).toBe('1,2,3')
    expect(doc.workbook.sheets[0].cells.C2).toBeUndefined()
  })
})

describe('CSV formula export limits', () => {
  it('refuses an unsupported formula result instead of exporting #NAME? as data', () => {
    let sheet = createBlankWorkbook('Export limits').workbook.sheets[0]
    sheet = upsertCell(sheet, 'A1', '1')
    sheet = upsertCell(sheet, 'A2', '2')
    sheet = upsertCell(sheet, 'B1', '=SUMPRODUCT(A1:A2,A1:A2)')
    expect(() => csvFromSheet(sheet)).toThrow('B1')
  })
})
