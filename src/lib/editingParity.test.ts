/**
 * Phase S1 — core editing parity at the model layer:
 * copy blocks, paste / paste-values / paste-formatting semantics, and
 * fill-handle (drag-fill) semantics. These are the shared semantics behind
 * both the Univer surface (persisted through the snapshot bridge) and the
 * fallback grid.
 */
import { describe, expect, it } from 'vitest'
import {
  applyTableStylePreset,
  copyRangeCells,
  createBlankWorkbook,
  createChartFromRange,
  displayCellValue,
  fillRange,
  pasteCellBlock,
  setCellStyle,
  updateChart,
  upsertCell,
} from './workbookModel'
import type { WorkbookSheet } from '../types'

function sheetWith(cells: Record<string, string>): WorkbookSheet {
  let sheet = createBlankWorkbook('Editing parity').workbook.sheets[0]
  for (const [key, value] of Object.entries(cells)) {
    sheet = upsertCell(sheet, key, value)
  }
  return sheet
}

describe('copy + paste (mode: all)', () => {
  it('pastes values, formulas (relatively shifted), and styles', () => {
    let sheet = sheetWith({ A1: '10', A2: '=A1*2' })
    sheet = setCellStyle(sheet, 'A1', { bold: true, fillColor: '#eef1eb' })
    const block = copyRangeCells(sheet, 'A1:A2')
    if (!block) throw new Error('expected a copied block')

    const pasted = pasteCellBlock(sheet, 'C5', block, 'all')
    expect(pasted.cells.C5).toEqual({
      value: '10',
      kind: 'number',
      style: { bold: true, fillColor: '#eef1eb' },
    })
    expect(pasted.cells.C6.value).toBe('=C5*2')
    expect(pasted.cells.C6.kind).toBe('formula')
  })

  it('keeps absolute references pinned when pasting', () => {
    const sheet = sheetWith({ A1: '5', B1: '=$A$1+A1' })
    const block = copyRangeCells(sheet, 'B1:B1')
    if (!block) throw new Error('expected a copied block')
    const pasted = pasteCellBlock(sheet, 'B3', block, 'all')
    expect(pasted.cells.B3.value).toBe('=$A$1+A3')
  })

  it('blank source cells clear the target, like a spreadsheet paste', () => {
    const sheet = sheetWith({ A1: 'keep-src', C5: 'stale', C6: 'stale too' })
    const block = copyRangeCells(sheet, 'A1:A2') // A2 is blank
    if (!block) throw new Error('expected a copied block')
    const pasted = pasteCellBlock(sheet, 'C5', block, 'all')
    expect(pasted.cells.C5.value).toBe('keep-src')
    expect(pasted.cells.C6).toBeUndefined()
  })

  it('grows the sheet when pasting past its bounds', () => {
    const sheet = sheetWith({ A1: 'x', B2: 'y' })
    const block = copyRangeCells(sheet, 'A1:B2')
    if (!block) throw new Error('expected a copied block')
    const bottomRight = `${'R'}${sheet.rowCount}` // R40 with default 18x40
    const pasted = pasteCellBlock(sheet, bottomRight, block, 'all')
    expect(pasted.rowCount).toBe(sheet.rowCount + 1)
    expect(pasted.columnCount).toBe(sheet.columnCount + 1)
  })
})

describe('paste values only', () => {
  it('pastes formula results as literal values and keeps target formatting', () => {
    let sheet = sheetWith({ A1: '1', A2: '2', B1: '=SUM(A1:A2)' })
    sheet = setCellStyle(sheet, 'B1', { bold: true })
    sheet = upsertCell(sheet, 'D1', 'old')
    sheet = setCellStyle(sheet, 'D1', { italic: true })
    const block = copyRangeCells(sheet, 'B1:B1')
    if (!block) throw new Error('expected a copied block')

    const pasted = pasteCellBlock(sheet, 'D1', block, 'values')
    expect(pasted.cells.D1).toEqual({
      value: '3',
      kind: 'number',
      style: { italic: true }, // target style kept; source bold NOT applied
    })
  })

  it('clears target values for blank source cells but keeps target styles', () => {
    let sheet = sheetWith({ C1: 'stale' })
    sheet = setCellStyle(sheet, 'C1', { fillColor: '#ffeeee' })
    const block = copyRangeCells(sheet, 'A1:A1') // blank source
    if (!block) throw new Error('expected a copied block')
    const pasted = pasteCellBlock(sheet, 'C1', block, 'values')
    expect(pasted.cells.C1).toEqual({
      value: '',
      kind: 'blank',
      style: { fillColor: '#ffeeee' },
    })
  })
})

describe('paste formatting only', () => {
  it('applies source styles without touching target values', () => {
    let sheet = sheetWith({ A1: 'styled', D1: 'hello' })
    sheet = setCellStyle(sheet, 'A1', { bold: true, textColor: '#333333' })
    const block = copyRangeCells(sheet, 'A1:A1')
    if (!block) throw new Error('expected a copied block')
    const pasted = pasteCellBlock(sheet, 'D1', block, 'formats')
    expect(pasted.cells.D1).toEqual({
      value: 'hello',
      kind: 'text',
      style: { bold: true, textColor: '#333333' },
    })
  })

  it('a style-less source clears the target formatting but keeps its value', () => {
    let sheet = sheetWith({ A1: 'plain', D1: 'hello' })
    sheet = setCellStyle(sheet, 'D1', { bold: true })
    const block = copyRangeCells(sheet, 'A1:A1')
    if (!block) throw new Error('expected a copied block')
    const pasted = pasteCellBlock(sheet, 'D1', block, 'formats')
    expect(pasted.cells.D1).toEqual({ value: 'hello', kind: 'text' })
  })
})

describe('fillRange (drag-fill semantics)', () => {
  it('continues a numeric series downward', () => {
    const sheet = sheetWith({ A1: '1', A2: '2' })
    const filled = fillRange(sheet, 'A1:A2', 'A3:A5')
    expect(filled.cells.A3.value).toBe('3')
    expect(filled.cells.A4.value).toBe('4')
    expect(filled.cells.A5.value).toBe('5')
    expect(filled.cells.A3.kind).toBe('number')
  })

  it('copies a single number instead of inventing a series', () => {
    const sheet = sheetWith({ A1: '7' })
    const filled = fillRange(sheet, 'A1:A1', 'A2:A3')
    expect(filled.cells.A2.value).toBe('7')
    expect(filled.cells.A3.value).toBe('7')
  })

  it('shifts relative formula references with the fill', () => {
    const sheet = sheetWith({ A1: '2', A2: '3', A3: '4', B1: '=A1*2' })
    const filled = fillRange(sheet, 'B1:B1', 'B2:B3')
    expect(filled.cells.B2.value).toBe('=A2*2')
    expect(filled.cells.B3.value).toBe('=A3*2')
  })

  it('fills horizontally, tiling text and styles', () => {
    let sheet = sheetWith({ A1: 'x' })
    sheet = setCellStyle(sheet, 'A1', { italic: true })
    const filled = fillRange(sheet, 'A1:A1', 'B1:C1')
    expect(filled.cells.B1).toEqual({
      value: 'x',
      kind: 'text',
      style: { italic: true },
    })
    expect(filled.cells.C1.value).toBe('x')
  })

  it('ignores targets that do not extend the source along one axis', () => {
    const sheet = sheetWith({ A1: '1' })
    expect(fillRange(sheet, 'A1:A1', 'B2:C3')).toBe(sheet)
  })
})

describe('decimal places in number formats (Phase S2)', () => {
  it('formats number, currency, and percent with explicit decimals', () => {
    let sheet = sheetWith({ A1: '3.14159', A2: '1200', A3: '0.5' })
    sheet = setCellStyle(sheet, 'A1', { numberFormat: 'number', decimals: 2 })
    sheet = setCellStyle(sheet, 'A2', { numberFormat: 'currency', decimals: 0 })
    sheet = setCellStyle(sheet, 'A3', { numberFormat: 'percent', decimals: 1 })
    expect(displayCellValue(sheet, 'A1')).toBe('3.14')
    expect(displayCellValue(sheet, 'A2')).toBe('$1,200')
    expect(displayCellValue(sheet, 'A3')).toBe('50.0%')
  })

  it('keeps the pre-decimals defaults when decimals are unset', () => {
    let sheet = sheetWith({ A1: '2', A2: '0.5' })
    sheet = setCellStyle(sheet, 'A1', { numberFormat: 'currency' })
    sheet = setCellStyle(sheet, 'A2', { numberFormat: 'percent' })
    expect(displayCellValue(sheet, 'A1')).toBe('$2.00')
    expect(displayCellValue(sheet, 'A2')).toBe('50%')
  })
})

describe('table style presets (Phase S2)', () => {
  const base = sheetWith({
    A1: 'Region',
    B1: 'Revenue',
    A2: 'North',
    B2: '10',
    A3: 'South',
    B3: '20',
    A4: 'Total',
    B4: '=SUM(B2:B3)',
  })

  it('header row: bold + fill + border on the first row only', () => {
    const styled = applyTableStylePreset(base, 'A1:B4', 'header-row')
    expect(styled.cells.A1.style).toEqual({
      bold: true,
      fillColor: '#eef1eb',
      border: true,
    })
    expect(styled.cells.B1.style?.bold).toBe(true)
    expect(styled.cells.A2.style).toBeUndefined()
  })

  it('banded rows: alternating fill on data rows, header untouched', () => {
    const styled = applyTableStylePreset(base, 'A1:B4', 'banded-rows')
    expect(styled.cells.A1.style).toBeUndefined()
    expect(styled.cells.A2.style?.fillColor).toBe('#f5f7f2')
    expect(styled.cells.A3.style?.fillColor).toBeUndefined()
    expect(styled.cells.A4.style?.fillColor).toBe('#f5f7f2')
  })

  it('totals row: bold + fill on the last row, formula preserved', () => {
    const styled = applyTableStylePreset(base, 'A1:B4', 'totals-row')
    expect(styled.cells.B4.style?.bold).toBe(true)
    expect(styled.cells.B4.value).toBe('=SUM(B2:B3)')
    expect(styled.cells.B4.kind).toBe('formula')
    expect(styled.cells.A1.style).toBeUndefined()
  })

  it('grid borders: every cell in the range gets a border', () => {
    const styled = applyTableStylePreset(base, 'A1:B4', 'grid-borders')
    for (const key of ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4']) {
      expect(styled.cells[key]?.style?.border, key).toBe(true)
    }
  })

  it('presets compose: header + bands + totals + grid', () => {
    let styled = applyTableStylePreset(base, 'A1:B4', 'header-row')
    styled = applyTableStylePreset(styled, 'A1:B4', 'banded-rows')
    styled = applyTableStylePreset(styled, 'A1:B4', 'totals-row')
    styled = applyTableStylePreset(styled, 'A1:B4', 'grid-borders')
    expect(styled.cells.A1.style?.bold).toBe(true)
    expect(styled.cells.A2.style?.fillColor).toBe('#f5f7f2')
    expect(styled.cells.A4.style?.bold).toBe(true)
    expect(styled.cells.B3.style?.border).toBe(true)
  })
})

describe('chart editor (Phase S4)', () => {
  const base = sheetWith({
    A1: 'North',
    B1: '10',
    A2: 'South',
    B2: '20',
    A3: 'West',
    B3: '30',
    A5: 'East',
    B5: '99',
  })

  function chartId(sheet: WorkbookSheet): string {
    const chart = sheet.charts?.[0]
    if (!chart) throw new Error('expected a chart')
    return chart.id
  }

  it('creates charts of every S4 type', () => {
    for (const type of ['area', 'pie', 'donut', 'scatter'] as const) {
      const sheet = createChartFromRange(base, 'A1:B3', type)
      expect(sheet.charts?.[0].type).toBe(type)
      expect(sheet.charts?.[0].values).toEqual([10, 20, 30])
    }
  })

  it('patches config fields in place without touching the data', () => {
    const sheet = createChartFromRange(base, 'A1:B3', 'pie')
    const updated = updateChart(sheet, chartId(sheet), {
      title: 'Revenue share',
      legend: 'bottom',
      showAxes: false,
      xLabel: 'Region',
      yLabel: 'Revenue',
    })
    expect(updated.charts?.[0]).toMatchObject({
      title: 'Revenue share',
      legend: 'bottom',
      showAxes: false,
      xLabel: 'Region',
      yLabel: 'Revenue',
      values: [10, 20, 30],
      range: 'A1:B3',
    })
  })

  it('re-derives the data when the range changes', () => {
    const sheet = createChartFromRange(base, 'A1:B3', 'bar')
    const updated = updateChart(sheet, chartId(sheet), { range: 'A5:B5' })
    expect(updated.charts?.[0].values).toEqual([99])
    expect(updated.charts?.[0].labels).toEqual(['East'])
    expect(updated.charts?.[0].range).toBe('A5:B5')
  })

  it('re-derives the data when the type changes', () => {
    const sheet = createChartFromRange(base, 'A1:B3', 'bar')
    const updated = updateChart(sheet, chartId(sheet), { type: 'donut' })
    expect(updated.charts?.[0].type).toBe('donut')
    expect(updated.charts?.[0].values).toEqual([10, 20, 30])
  })

  it('refuses ranges without plottable data and unknown charts', () => {
    const sheet = createChartFromRange(base, 'A1:B3', 'bar')
    expect(updateChart(sheet, chartId(sheet), { range: 'D1:D3' })).toBe(sheet)
    expect(updateChart(sheet, 'missing-chart', { title: 'x' })).toBe(sheet)
  })
})
