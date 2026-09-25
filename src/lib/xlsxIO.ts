import ExcelJS from 'exceljs'
import {
  cellKey,
  computedCellValue,
  createBlankSheet,
  createBlankWorkbook,
} from './workbookModel'
import { withFreshUniverSnapshot } from './univerAdapter'
import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT } from '../constants'
import type {
  CellStyle,
  PureSheetsDocument,
  SheetCell,
  WorkbookSheet,
} from '../types'

/**
 * XLSX import/export (Phase S4), built on exceljs.
 *
 * Decision: no XLSX library existed in the workspace; `exceljs` was added as
 * a puresheets dependency (well-known, actively maintained, reads AND writes
 * values, formulas, and basic styles — the npm `xlsx` package is stale on the
 * public registry and its community edition cannot write styles).
 *
 * Import maps common workbooks into the PureSheets model and returns
 * explicit, user-visible warnings for anything that cannot be preserved.
 * `.sheets` remains the native format — an import produces an ordinary
 * PureSheetsDocument through the normal model.
 */

export interface XlsxImportResult {
  document: PureSheetsDocument
  warnings: string[]
}

export async function importXlsxWorkbook(
  data: ArrayBuffer | Uint8Array,
  title: string,
): Promise<XlsxImportResult> {
  const workbook = new ExcelJS.Workbook()
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
  await workbook.xlsx.load(bytes.buffer as ArrayBuffer)
  const warnings: string[] = []
  const sheets: WorkbookSheet[] = []
  let flattenedRichText = 0
  let flattenedHyperlinks = 0

  workbook.eachSheet(worksheet => {
    const id = `sheet-xlsx-${sheets.length + 1}`
    let sheet = createBlankSheet(id, worksheet.name || `Sheet ${sheets.length + 1}`)
    const cells: Record<string, SheetCell> = {}
    let maxRow = 0
    let maxColumn = 0

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
        const mapped = mapXlsxCell(cell)
        if (!mapped) return
        if (mapped.richText) flattenedRichText += 1
        if (mapped.hyperlink) flattenedHyperlinks += 1
        const key = cellKey(rowNumber - 1, columnNumber - 1)
        cells[key] = mapped.cell
        maxRow = Math.max(maxRow, rowNumber)
        maxColumn = Math.max(maxColumn, columnNumber)
      })
    })

    // Column widths (Excel character units → approximate pixels).
    const columnWidths: Record<string, number> = {}
    worksheet.columns?.forEach((column, index) => {
      if (typeof column?.width === 'number') {
        columnWidths[String(index)] = clamp(Math.round(column.width * 7.5), 56, 360)
      }
    })
    // Row heights (points → approximate pixels).
    const rowHeights: Record<string, number> = {}
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (typeof row.height === 'number') {
        rowHeights[String(rowNumber - 1)] = clamp(
          Math.round(row.height * 1.34),
          24,
          160,
        )
      }
    })

    const view = worksheet.views?.[0]
    const frozen =
      view && view.state === 'frozen'
        ? { rows: view.ySplit ?? 0, columns: view.xSplit ?? 0 }
        : { rows: 0, columns: 0 }

    const mergeCount = Object.keys(
      (worksheet.model as { merges?: unknown[] }).merges ?? {},
    ).length
    if (mergeCount > 0) {
      warnings.push(
        `${worksheet.name}: ${mergeCount} merged cell range${
          mergeCount === 1 ? '' : 's'
        } could not be preserved (values kept in the top-left cell).`,
      )
    }
    if (worksheet.getImages().length > 0) {
      warnings.push(
        `${worksheet.name}: embedded images were not imported.`,
      )
    }
    const validationCount = Object.keys(
      (
        (worksheet as unknown as {
          dataValidations?: { model?: Record<string, unknown> }
        }).dataValidations ?? {}
      ).model ?? {},
    ).length
    if (validationCount > 0) {
      warnings.push(
        `${worksheet.name}: ${validationCount} Excel data-validation rule${
          validationCount === 1 ? '' : 's'
        } were not imported — recreate them with the Data validation tool.`,
      )
    }

    sheet = {
      ...sheet,
      rowCount: Math.max(DEFAULT_ROW_COUNT, maxRow),
      columnCount: Math.max(DEFAULT_COLUMN_COUNT, maxColumn),
      cells,
      columnWidths: Object.keys(columnWidths).length ? columnWidths : undefined,
      rowHeights: Object.keys(rowHeights).length ? rowHeights : undefined,
      frozenRows: frozen.rows || undefined,
      frozenColumns: frozen.columns || undefined,
    }
    sheets.push(sheet)
  })

  if (flattenedRichText > 0) {
    warnings.push(
      `${flattenedRichText} rich-text cell${
        flattenedRichText === 1 ? ' was' : 's were'
      } flattened to plain text.`,
    )
  }
  if (flattenedHyperlinks > 0) {
    warnings.push(
      `${flattenedHyperlinks} hyperlink${
        flattenedHyperlinks === 1 ? ' was' : 's were'
      } kept as text only.`,
    )
  }
  warnings.push(
    'Excel charts, pivot tables, macros, and comments are not imported.',
  )

  if (!sheets.length) {
    return { document: createBlankWorkbook(title), warnings }
  }
  const now = new Date().toISOString()
  const document = withFreshUniverSnapshot({
    app: 'PureSheets',
    version: 1,
    engine: { name: 'univer' },
    metadata: { title, createdAt: now, updatedAt: now },
    workbook: { activeSheetId: sheets[0].id, sheets },
  })
  return { document, warnings }
}

function mapXlsxCell(
  cell: ExcelJS.Cell,
): { cell: SheetCell; richText?: boolean; hyperlink?: boolean } | null {
  let value = ''
  let kind: SheetCell['kind'] = 'text'
  let richText = false
  let hyperlink = false
  const raw = cell.value
  if (raw === null || raw === undefined) {
    value = ''
  } else if (typeof raw === 'number') {
    value = String(raw)
    kind = 'number'
  } else if (typeof raw === 'boolean') {
    value = raw ? 'TRUE' : 'FALSE'
  } else if (raw instanceof Date) {
    value = raw.toISOString().slice(0, 10)
  } else if (typeof raw === 'object') {
    const objectValue = raw as {
      formula?: string
      sharedFormula?: string
      result?: unknown
      richText?: Array<{ text: string }>
      text?: string
      hyperlink?: string
      error?: string
    }
    if (objectValue.formula) {
      value = `=${objectValue.formula}`
      kind = 'formula'
    } else if (objectValue.sharedFormula) {
      // Shared formulas resolve to their cached result — the master formula
      // text is not addressable per-cell here.
      value = objectValue.result === undefined ? '' : String(objectValue.result)
    } else if (objectValue.richText) {
      value = objectValue.richText.map(part => part.text).join('')
      richText = true
    } else if (objectValue.hyperlink) {
      value = objectValue.text ?? objectValue.hyperlink
      hyperlink = true
    } else if (objectValue.error) {
      value = objectValue.error
    } else {
      value = String(raw)
    }
  } else {
    value = String(raw)
  }
  const style = mapXlsxStyle(cell)
  if (!value && !style) return null
  if (kind === 'text' && value !== '' && Number.isFinite(Number(value))) {
    kind = 'number'
  }
  if (!value) kind = 'blank'
  return {
    cell: { value, kind, ...(style ? { style } : {}) },
    richText,
    hyperlink,
  }
}

function mapXlsxStyle(cell: ExcelJS.Cell): CellStyle | undefined {
  const style: CellStyle = {}
  const font = cell.font
  if (font?.bold) style.bold = true
  if (font?.italic) style.italic = true
  if (font?.underline) style.underline = true
  if (typeof font?.size === 'number' && font.size !== 11) {
    style.fontSize = clamp(Math.round(font.size), 8, 48)
  }
  const textColor = argbToHex(font?.color?.argb)
  if (textColor) style.textColor = textColor
  const fill = cell.fill as
    | { type?: string; fgColor?: { argb?: string } }
    | undefined
  if (fill?.type === 'pattern') {
    const fillColor = argbToHex(fill.fgColor?.argb)
    if (fillColor && fillColor.toLowerCase() !== '#ffffff') {
      style.fillColor = fillColor
    }
  }
  const border = cell.border
  if (border && (border.top || border.bottom || border.left || border.right)) {
    style.border = true
  }
  const alignment = cell.alignment?.horizontal
  if (alignment === 'left' || alignment === 'center' || alignment === 'right') {
    style.align = alignment
  }
  const numberFormat = mapNumFmt(cell.numFmt)
  if (numberFormat) {
    style.numberFormat = numberFormat.format
    if (numberFormat.decimals !== undefined) {
      style.decimals = numberFormat.decimals
    }
  }
  return Object.keys(style).length ? style : undefined
}

function mapNumFmt(
  numFmt: string | undefined,
): { format: NonNullable<CellStyle['numberFormat']>; decimals?: number } | null {
  if (!numFmt || numFmt === 'General') return null
  if (numFmt === '@') return { format: 'text' }
  const zeros = /\.(0+)/.exec(numFmt)?.[1]?.length
  if (numFmt.includes('$') || /\[\$/.test(numFmt)) {
    return { format: 'currency', decimals: zeros === 2 ? undefined : zeros ?? 0 }
  }
  if (numFmt.includes('%')) {
    return { format: 'percent', decimals: zeros === 2 ? undefined : zeros ?? 0 }
  }
  if (/[ymd]/i.test(numFmt) && !numFmt.includes('#')) {
    return { format: 'date' }
  }
  if (/[#0]/.test(numFmt)) {
    return { format: 'number', decimals: zeros }
  }
  return null
}

function argbToHex(argb: string | undefined): string | undefined {
  if (!argb || argb.length < 6) return undefined
  const rgb = argb.length === 8 ? argb.slice(2) : argb
  if (!/^[0-9a-f]{6}$/i.test(rgb)) return undefined
  return `#${rgb.toLowerCase()}`
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------

/** Export the whole workbook to XLSX bytes (values, formulas, basic styles). */
export async function exportXlsxWorkbook(
  document: PureSheetsDocument,
): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  workbook.created = new Date()
  for (const sheet of document.workbook.sheets) {
    // Agent draft sheets are internal review state, not export content.
    if (sheet.agentDraft) continue
    const worksheet = workbook.addWorksheet(sheet.name || 'Sheet')
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const excelCell = worksheet.getCell(key)
      if (cell.kind === 'formula') {
        excelCell.value = {
          formula: cell.value.replace(/^=/, ''),
          result: numericOrText(computedCellValue(sheet, key)),
        } as ExcelJS.CellFormulaValue
      } else if (cell.kind === 'number' && Number.isFinite(Number(cell.value))) {
        excelCell.value = Number(cell.value)
      } else if (cell.value) {
        excelCell.value = cell.value
      }
      applyStyleToExcelCell(excelCell, cell.style)
    }
    for (const [column, width] of Object.entries(sheet.columnWidths ?? {})) {
      worksheet.getColumn(Number(column) + 1).width = Math.max(
        4,
        Math.round(width / 7.5),
      )
    }
    for (const [row, height] of Object.entries(sheet.rowHeights ?? {})) {
      worksheet.getRow(Number(row) + 1).height = Math.round(height / 1.34)
    }
    if (sheet.frozenRows || sheet.frozenColumns) {
      worksheet.views = [
        {
          state: 'frozen',
          ySplit: sheet.frozenRows ?? 0,
          xSplit: sheet.frozenColumns ?? 0,
        },
      ]
    }
  }
  const buffer = await workbook.xlsx.writeBuffer()
  return new Uint8Array(buffer as ArrayBuffer)
}

function numericOrText(value: string): number | string {
  const numeric = Number(value)
  return value.trim() !== '' && Number.isFinite(numeric) ? numeric : value
}

function applyStyleToExcelCell(
  excelCell: ExcelJS.Cell,
  style: CellStyle | undefined,
): void {
  if (!style) return
  const font: Partial<ExcelJS.Font> = {}
  if (style.bold) font.bold = true
  if (style.italic) font.italic = true
  if (style.underline) font.underline = true
  if (style.fontSize) font.size = style.fontSize
  if (style.textColor) font.color = { argb: hexToArgb(style.textColor) }
  if (Object.keys(font).length) excelCell.font = font as ExcelJS.Font
  if (style.fillColor) {
    excelCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: hexToArgb(style.fillColor) },
    }
  }
  if (style.border) {
    const line: Partial<ExcelJS.Border> = { style: 'thin' }
    excelCell.border = { top: line, bottom: line, left: line, right: line }
  }
  if (style.align) {
    excelCell.alignment = { horizontal: style.align }
  }
  const pattern = excelNumFmt(style)
  if (pattern) excelCell.numFmt = pattern
}

function excelNumFmt(style: CellStyle): string | undefined {
  const decimals = style.decimals
  const zeros =
    decimals !== undefined && decimals > 0
      ? `.${'0'.repeat(Math.min(decimals, 8))}`
      : ''
  if (style.numberFormat === 'currency') {
    return decimals === undefined ? '$#,##0.00' : `$#,##0${zeros}`
  }
  if (style.numberFormat === 'percent') {
    return decimals === undefined ? '0.00%' : `0${zeros}%`
  }
  if (style.numberFormat === 'date') return 'yyyy-mm-dd'
  if (style.numberFormat === 'number') {
    return decimals === undefined ? undefined : `#,##0${zeros}`
  }
  if (style.numberFormat === 'text') return '@'
  return undefined
}

function hexToArgb(hex: string): string {
  const rgb = hex.replace('#', '')
  return `FF${rgb.toUpperCase()}`
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

// ---------------------------------------------------------------------------
// Base64 helpers for the binary bridge (renderer-safe, chunked).
// ---------------------------------------------------------------------------

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunkSize = 0x8000
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize))
  }
  return btoa(binary)
}
