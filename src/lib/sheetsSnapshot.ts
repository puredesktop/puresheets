import type { PureSheetsDocument, SheetCell } from '../types'

/**
 * A distilled snapshot of the workbook for the document switcher: the
 * active sheet's top-left cells drawn as one self-contained SVG grid.
 */

const COLS = 8
const ROWS = 12
const CELL_W = 110
const CELL_H = 46
const PAD = 8

function cellAt(
  cells: Record<string, SheetCell>,
  column: number,
  row: number,
): SheetCell | undefined {
  return cells[`${String.fromCharCode(65 + column)}${row + 1}`]
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

export function sheetsSnapshotHtml(
  document: PureSheetsDocument,
): string | null {
  const active =
    document.workbook.sheets.find(
      candidate => candidate.id === document.workbook.activeSheetId,
    ) ?? document.workbook.sheets[0]
  const sheet = active && Object.keys(active.cells ?? {}).length
    ? active
    : document.workbook.sheets.find(candidate => Object.keys(candidate.cells ?? {}).length)
  if (!sheet) return null
  const cells = sheet.cells ?? {}
  if (!Object.keys(cells).length) return null

  const width = COLS * CELL_W + PAD * 2
  const height = ROWS * CELL_H + PAD * 2
  const parts: string[] = []
  for (let row = 0; row <= ROWS; row += 1) {
    const y = PAD + row * CELL_H
    parts.push(
      `<line x1="${PAD}" y1="${y}" x2="${width - PAD}" y2="${y}" stroke="#e6e3ec" stroke-width="2"/>`,
    )
  }
  for (let column = 0; column <= COLS; column += 1) {
    const x = PAD + column * CELL_W
    parts.push(
      `<line x1="${x}" y1="${PAD}" x2="${x}" y2="${height - PAD}" stroke="#e6e3ec" stroke-width="2"/>`,
    )
  }
  for (let row = 0; row < ROWS; row += 1) {
    for (let column = 0; column < COLS; column += 1) {
      const cell = cellAt(cells, column, row)
      const value = cell?.value?.trim()
      if (!value) continue
      const x = PAD + column * CELL_W + 8
      const y = PAD + row * CELL_H + CELL_H / 2 + 5
      const numeric = cell?.kind === 'number' || /^-?[\d,.%$€£]+$/.test(value)
      parts.push(
        `<text x="${numeric ? PAD + (column + 1) * CELL_W - 8 : x}" y="${y}" text-anchor="${numeric ? 'end' : 'start'}" font-family="Archivo, system-ui, sans-serif" font-size="15" font-weight="${row === 0 ? 700 : 400}" fill="${row === 0 ? '#2b2933' : '#57545f'}">${escapeXml(value.slice(0, 12))}</text>`,
      )
    }
  }
  return `<!doctype html>
<html><head><meta charset="utf-8"><style>html,body{margin:0;background:#ffffff}</style></head>
<body><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">${parts.join('')}</svg></body></html>`
}
