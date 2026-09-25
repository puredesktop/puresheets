import { expect, it, vi } from 'vitest'
import { readWorkbookDatasets } from './data'
import { createBlankWorkbook, createBlankSheet } from './lib/workbookModel'

it('exports separate sheets, calculated formulas and quoted text without changing the workbook', async () => {
  const doc = createBlankWorkbook('Sales')
  doc.workbook.sheets[0].cells = {
    A1: { kind: 'text', value: 'Name' }, B1: { kind: 'text', value: 'Total' },
    A2: { kind: 'text', value: 'A, B' }, B2: { kind: 'formula', value: '=2+3' },
  }
  doc.workbook.sheets.push(createBlankSheet('empty', 'Empty'))
  const raw = JSON.stringify(doc)
  const read = vi.fn().mockResolvedValue(raw)
  const result = await readWorkbookDatasets('/Sales.sheets', read)
  expect(read).toHaveBeenCalledWith('/Sales.sheets/workbook.json')
  expect(result.sheets[0].csv).toBe('Name,Total\n"A, B",5\n')
  expect(result.sheets[1].csv).toBe('')
  expect(JSON.stringify(doc)).toBe(raw)
})

it('reads portable HTML workbooks and rejects an invalid format', async () => {
  const doc = createBlankWorkbook('Portable')
  const result = await readWorkbookDatasets('/Portable.sheets.html', async () => `<script id="puresheets-workbook-data">${JSON.stringify(doc)}</script>`)
  expect(result.title).toBe('Portable')
  await expect(readWorkbookDatasets('/bad.sheets', async () => '{}')).rejects.toThrow('not a PureSheets')
})

it('does not silently turn a formula error into report data', async () => {
  const doc = createBlankWorkbook('Broken')
  doc.workbook.sheets[0].cells = { A1: { kind: 'text', value: 'Total' }, A2: { kind: 'formula', value: '=1/0' } }
  const result = await readWorkbookDatasets('/Broken.sheets', async () => JSON.stringify(doc))
  expect(result.sheets[0].error).toContain('Formula at A2')
  expect(result.sheets[0].csv).toBe('')
})

it('exports two populated sheets with independent schemas and formula results', async () => {
  const doc = createBlankWorkbook('Two sheets')
  const revenue = doc.workbook.sheets[0]
  revenue.name = 'Revenue'
  revenue.cells = {
    A1: { kind: 'text', value: 'month' }, B1: { kind: 'text', value: 'revenue_usd' },
    A2: { kind: 'text', value: 'Jan' }, B2: { kind: 'formula', value: '=100*10' },
  }
  const costs = createBlankSheet('costs', 'Costs')
  costs.cells = {
    A1: { kind: 'text', value: 'month' }, B1: { kind: 'text', value: 'note' },
    A2: { kind: 'text', value: 'Jan' }, B2: { kind: 'text', value: 'Steady, controlled' },
  }
  doc.workbook.sheets.push(costs)
  const result = await readWorkbookDatasets('/two.sheets', async () => JSON.stringify(doc))
  expect(result.sheets).toEqual([
    { id: revenue.id, name: 'Revenue', csv: 'month,revenue_usd\nJan,1000\n' },
    { id: costs.id, name: 'Costs', csv: 'month,note\nJan,"Steady, controlled"\n' },
  ])
})
