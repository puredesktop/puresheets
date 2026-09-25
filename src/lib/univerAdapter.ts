import {
  BorderStyleTypes,
  BooleanNumber,
  CellValueType,
  HorizontalAlign,
  LocaleType,
  TextDecoration,
  Univer,
  UniverInstanceType,
  VerticalAlign,
  WrapStrategy,
  type ICellData,
  type IStyleData,
  type IWorkbookData,
  type IWorksheetData,
  type Workbook,
} from '@univerjs/core'
import {
  CELL_BORDER_COLOR,
  DEFAULT_COLUMN_COUNT,
  DEFAULT_ROW_COUNT,
  UNIVER_SURFACE_COLUMN_COUNT,
  UNIVER_SURFACE_ROW_COUNT,
} from '../constants'

import {
  normalizeComments,
  normalizeConditionalFormats,
  normalizeFilterQuery,
  normalizeProtection,
  normalizeValidations,
} from './sheetMetadataNormalize'
import type {
  CellStyle,
  PureSheetsDocument,
  SheetCell,
  WorkbookSheet,
} from '../types'

const UNIVER_APP_VERSION = '0.25.0'
const DEFAULT_COLUMN_WIDTH = 96
const DEFAULT_ROW_HEIGHT = 28
const ROW_HEADER_WIDTH = 46
const COLUMN_HEADER_HEIGHT = 24

export type PureSheetsUniverSnapshot = IWorkbookData

export function withFreshUniverSnapshot(
  document: PureSheetsDocument,
): PureSheetsDocument {
  return {
    ...document,
    engine: {
      name: 'univer',
      version: UNIVER_APP_VERSION,
      snapshot: createUniverSnapshot(document),
    },
  }
}

export function createUniverSnapshot(
  document: PureSheetsDocument,
): PureSheetsUniverSnapshot {
  const sheets = Object.fromEntries(
    document.workbook.sheets.map(sheet => [
      sheet.id,
      createUniverSheetSnapshot(sheet),
    ]),
  )

  return {
    id: `puresheets-${stableId(document.metadata.title)}`,
    name: document.metadata.title,
    appVersion: UNIVER_APP_VERSION,
    locale: LocaleType.EN_US,
    styles: collectStyles(document),
    sheetOrder: document.workbook.sheets.map(sheet => sheet.id),
    sheets,
    custom: {
      puresheets: {
        activeSheetId: document.workbook.activeSheetId,
        updatedAt: document.metadata.updatedAt,
      },
    },
  }
}

export function validateUniverSnapshot(
  snapshot: PureSheetsUniverSnapshot,
): PureSheetsUniverSnapshot {
  const univer = new Univer({ locale: LocaleType.EN_US })
  try {
    const workbook = univer.createUnit<IWorkbookData, Workbook>(
      UniverInstanceType.UNIVER_SHEET,
      snapshot,
    )
    return workbook.save()
  } finally {
    univer.dispose()
  }
}

export function documentFromUniverSnapshot(
  snapshot: PureSheetsUniverSnapshot,
  fallback?: PureSheetsDocument,
): PureSheetsDocument {
  const saved = validateUniverSnapshot(snapshot)
  const sheetOrder = saved.sheetOrder?.length
    ? saved.sheetOrder
    : Object.keys(saved.sheets ?? {})
  const fallbackSheetsById = new Map(
    fallback?.workbook.sheets.map(sheet => [sheet.id, sheet]) ?? [],
  )
  const sheets = sheetOrder.flatMap(sheetId => {
    const sheet = saved.sheets?.[sheetId]
    if (!sheet) return []
    return [
      workbookSheetFromUniverSheet(
        sheetId,
        sheet,
        saved.styles,
        fallbackSheetsById.get(sheetId),
      ),
    ]
  })

  if (!sheets.length) {
    throw new Error('Univer snapshot does not contain any worksheets.')
  }

  const custom =
    isRecord(saved.custom) && isRecord(saved.custom.puresheets)
      ? saved.custom.puresheets
      : {}
  const activeSheetId =
    typeof custom.activeSheetId === 'string' &&
    sheets.some(sheet => sheet.id === custom.activeSheetId)
      ? custom.activeSheetId
      : sheets[0].id
  const now = new Date().toISOString()

  return withFreshUniverSnapshot({
    app: 'PureSheets',
    version: 1,
    engine: {
      name: 'univer',
      version: UNIVER_APP_VERSION,
    },
    metadata: {
      title:
        typeof saved.name === 'string' && saved.name.trim()
          ? saved.name
          : fallback?.metadata.title ?? 'Untitled',
      createdAt: fallback?.metadata.createdAt ?? now,
      updatedAt:
        typeof custom.updatedAt === 'string'
          ? custom.updatedAt
          : fallback?.metadata.updatedAt ?? now,
    },
    workbook: {
      activeSheetId,
      sheets,
      agentLog: fallback?.workbook.agentLog,
    },
  })
}

/** App chrome owns the title; the mounted editor may still hold its old name. */
export function documentFromEditorSnapshot(
  snapshot: PureSheetsUniverSnapshot,
  current: PureSheetsDocument,
): PureSheetsDocument {
  return documentFromUniverSnapshot({ ...snapshot, name: current.metadata.title }, current)
}

function createUniverSheetSnapshot(
  sheet: WorkbookSheet,
): Partial<IWorksheetData> {
  return {
    id: sheet.id,
    name: sheet.name,
    tabColor: '',
    hidden: BooleanNumber.FALSE,
    freeze: {
      startRow: sheet.frozenRows ? sheet.frozenRows - 1 : -1,
      startColumn: sheet.frozenColumns ? sheet.frozenColumns - 1 : -1,
      ySplit: sheet.frozenRows ?? 0,
      xSplit: sheet.frozenColumns ?? 0,
    },
    // Render a full spreadsheet-sized canvas, growing past it only when the
    // sheet already holds more data — so scrolling never runs off into the
    // grey void the old 40×18 floor produced (#272).
    rowCount: Math.max(sheet.rowCount, UNIVER_SURFACE_ROW_COUNT),
    columnCount: Math.max(sheet.columnCount, UNIVER_SURFACE_COLUMN_COUNT),
    zoomRatio: 1,
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth: DEFAULT_COLUMN_WIDTH,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
    mergeData: [],
    cellData: collectCellData(sheet),
    rowData: Object.fromEntries(
      Object.entries(sheet.rowHeights ?? {}).map(([row, height]) => [
        row,
        { h: height },
      ]),
    ),
    columnData: Object.fromEntries(
      Object.entries(sheet.columnWidths ?? {}).map(([column, width]) => [
        column,
        { w: width },
      ]),
    ),
    rowHeader: { width: ROW_HEADER_WIDTH },
    columnHeader: { height: COLUMN_HEADER_HEIGHT },
    showGridlines: BooleanNumber.TRUE,
    rightToLeft: BooleanNumber.FALSE,
    custom: {
      puresheets: {
        filterRow: sheet.filterRow,
        filterColumn: sheet.filterColumn,
        filterHasHeader: sheet.filterHasHeader ?? null,
        // Phase S3 metadata. Keys are ALWAYS written (null when unset) so the
        // restore path can distinguish "cleared by the user" from "legacy
        // snapshot that never carried the field" (which falls back).
        filterQuery: sheet.filterQuery ?? null,
        comments: sheet.comments ?? null,
        validations: sheet.validations ?? null,
        protection: sheet.protection ?? null,
        conditionalFormats: sheet.conditionalFormats ?? null,
      },
    },
  }
}

function workbookSheetFromUniverSheet(
  id: string,
  sheet: Partial<IWorksheetData>,
  styles: IWorkbookData['styles'] | undefined,
  fallback?: WorkbookSheet,
): WorkbookSheet {
  const custom =
    isRecord(sheet.custom) && isRecord(sheet.custom.puresheets)
      ? sheet.custom.puresheets
      : {}
  return {
    id,
    name:
      typeof sheet.name === 'string' && sheet.name.trim()
        ? sheet.name
        : fallback?.name ?? id,
    rowCount:
      positiveInteger(sheet.rowCount) ??
      fallback?.rowCount ??
      DEFAULT_ROW_COUNT,
    columnCount:
      positiveInteger(sheet.columnCount) ??
      fallback?.columnCount ??
      DEFAULT_COLUMN_COUNT,
    cells: cellsFromUniverSheet(sheet.cellData, styles),
    columnWidths:
      dimensionsFromUniverData(sheet.columnData, 'w') ?? fallback?.columnWidths,
    rowHeights:
      dimensionsFromUniverData(sheet.rowData, 'h') ?? fallback?.rowHeights,
    frozenRows:
      typeof sheet.freeze?.ySplit === 'number'
        ? sheet.freeze.ySplit
        : fallback?.frozenRows,
    frozenColumns:
      typeof sheet.freeze?.xSplit === 'number'
        ? sheet.freeze.xSplit
        : fallback?.frozenColumns,
    filterRow: optionalNumber(custom.filterRow) ?? fallback?.filterRow,
    filterColumn: optionalNumber(custom.filterColumn) ?? fallback?.filterColumn,
    filterHasHeader:
      typeof custom.filterHasHeader === 'boolean'
        ? custom.filterHasHeader
        : fallback?.filterHasHeader,
    // Phase S3 metadata: when the snapshot carries the key (even as null),
    // the snapshot wins; a missing key means a legacy snapshot → fallback.
    filterQuery:
      'filterQuery' in custom
        ? normalizeFilterQuery(custom.filterQuery)
        : fallback?.filterQuery,
    comments:
      'comments' in custom
        ? normalizeComments(custom.comments)
        : fallback?.comments,
    validations:
      'validations' in custom
        ? normalizeValidations(custom.validations)
        : fallback?.validations,
    protection:
      'protection' in custom
        ? normalizeProtection(custom.protection)
        : fallback?.protection,
    conditionalFormats:
      'conditionalFormats' in custom
        ? normalizeConditionalFormats(custom.conditionalFormats)
        : fallback?.conditionalFormats,
    charts: fallback?.charts,
    agentDraft: fallback?.agentDraft,
  }
}

function cellsFromUniverSheet(
  cellData: IWorksheetData['cellData'] | undefined,
  styles: IWorkbookData['styles'] | undefined,
): Record<string, SheetCell> {
  const cells: Record<string, SheetCell> = {}
  for (const [rowIndex, row] of Object.entries(cellData ?? {})) {
    if (!row) continue
    for (const [columnIndex, rawCell] of Object.entries(row)) {
      const cell = rawCell as ICellData
      const key = toCellKey(Number(rowIndex), Number(columnIndex))
      if (!key || !cell) continue
      const value = univerCellValue(cell)
      if (!value && !cell.f && !cell.s) continue
      const style =
        typeof cell.s === 'string'
          ? fromUniverStyle(styles?.[cell.s] ?? undefined)
          : undefined
      cells[key] = {
        value,
        kind: cell.f
          ? 'formula'
          : cell.t === CellValueType.NUMBER
          ? 'number'
          : value
          ? 'text'
          : 'blank',
        style,
      }
    }
  }
  return cells
}

function univerCellValue(cell: ICellData): string {
  if (cell.f) return cell.f
  if (cell.v === null || cell.v === undefined) return ''
  return String(cell.v)
}

function fromUniverStyle(style: IStyleData | undefined): CellStyle | undefined {
  if (!style) return undefined
  const cellStyle: CellStyle = {}
  if (style.tb === WrapStrategy.WRAP) cellStyle.wrap = 'wrap'
  if (style.tb === WrapStrategy.CLIP) cellStyle.wrap = 'clip'
  if (style.tb === WrapStrategy.OVERFLOW) cellStyle.wrap = 'overflow'
  if (style.bl === BooleanNumber.TRUE) cellStyle.bold = true
  if (style.it === BooleanNumber.TRUE) cellStyle.italic = true
  // Univer encodes underline as `{ s: BooleanNumber, t: TextDecoration }` —
  // check the `s` flag, not object truthiness, so an explicitly disabled
  // decoration (`{ s: 0 }`) doesn't round-trip back as underlined (#270).
  if (typeof style.ul === 'object' ? style.ul?.s : style.ul) {
    cellStyle.underline = true
  }
  if (typeof style.fs === 'number') cellStyle.fontSize = style.fs
  if (typeof style.cl?.rgb === 'string') cellStyle.textColor = style.cl.rgb
  if (typeof style.bg?.rgb === 'string') cellStyle.fillColor = style.bg.rgb
  if (
    style.bd &&
    Object.values(style.bd).some(border => border?.s === BorderStyleTypes.THIN)
  ) {
    cellStyle.border = true
  }
  if (style.ht === HorizontalAlign.CENTER) cellStyle.align = 'center'
  if (style.ht === HorizontalAlign.RIGHT) cellStyle.align = 'right'
  if (style.ht === HorizontalAlign.LEFT) cellStyle.align = 'left'
  const pattern = style.n?.pattern
  if (typeof pattern === 'string') {
    cellStyle.numberFormat = numberFormatFromPattern(pattern)
    const decimals = decimalsFromPattern(pattern, cellStyle.numberFormat)
    if (decimals !== undefined) cellStyle.decimals = decimals
  }
  return Object.keys(cellStyle).length ? cellStyle : undefined
}

function numberFormatFromPattern(
  pattern: string,
): NonNullable<CellStyle['numberFormat']> {
  if (pattern.includes('$')) return 'currency'
  if (pattern.includes('%')) return 'percent'
  if (/y|m|d/i.test(pattern)) return 'date'
  if (pattern === '@') return 'text'
  return 'number'
}

/**
 * Recover the explicit decimal count from a number pattern. Patterns that
 * match a format's default (currency/percent: two places; plain number:
 * optional `.########` digits) normalize to "unset" so styles round-trip
 * stably — the display is identical either way.
 */
function decimalsFromPattern(
  pattern: string,
  format: NonNullable<CellStyle['numberFormat']>,
): number | undefined {
  if (format === 'date' || format === 'text') return undefined
  if (/\.#/.test(pattern)) return undefined // optional digits: as entered
  const zeros = /\.(0+)/.exec(pattern)?.[1]?.length ?? 0
  if (format === 'number') {
    if (zeros > 0) return zeros
    // Only a strict integer digit pattern (e.g. '#,##0') means "0 decimals";
    // anything else ('General', free-form patterns) stays as-entered.
    return /^[#,0]+$/.test(pattern) ? 0 : undefined
  }
  // currency / percent default to two places
  return zeros === 2 ? undefined : zeros
}

function dimensionsFromUniverData(
  data: IWorksheetData['rowData'] | IWorksheetData['columnData'] | undefined,
  key: 'h' | 'w',
): Record<string, number> | undefined {
  const dimensions: Record<string, number> = {}
  for (const [index, value] of Object.entries(data ?? {})) {
    // Native auto-fit stores its measured height in ah, separately from the
    // manual minimum h. Carry the displayed height into model-driven reloads.
    const row = value as { h?: number; ah?: number; ia?: BooleanNumber } | undefined
    const dimension = key === 'h' && row?.ia !== BooleanNumber.FALSE &&
      typeof row?.ah === 'number' && row.ah > 0 ? row.ah : value?.[key]
    if (typeof dimension === 'number' && Number.isFinite(dimension)) {
      dimensions[index] = dimension
    }
  }
  return Object.keys(dimensions).length ? dimensions : undefined
}

function collectCellData(sheet: WorkbookSheet): IWorksheetData['cellData'] {
  const rows: IWorksheetData['cellData'] = {}
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position) continue
    rows[position.row] = {
      ...rows[position.row],
      [position.column]: createUniverCell(cell),
    }
  }
  return rows
}

function createUniverCell(cell: SheetCell): ICellData {
  const value = cell.value
  const data: ICellData = {}
  if (cell.kind === 'formula') {
    data.f = value
  } else if (cell.kind === 'number' && Number.isFinite(Number(value))) {
    data.v = Number(value)
    data.t = CellValueType.NUMBER
  } else if (cell.kind === 'number') {
    // A cell tagged "number" whose raw value is not a finite number (e.g. a
    // stale kind on legacy data, or "N/A"). Number(value) would be NaN, which
    // JSON.stringify writes as null — silently destroying the text on save.
    // Preserve the literal value as a string instead.
    data.v = value
    data.t = CellValueType.STRING
  } else {
    data.v = value
    data.t = CellValueType.STRING
  }
  const styleId = styleKey(cell.style)
  if (styleId) data.s = styleId
  return data
}

function collectStyles(document: PureSheetsDocument): IWorkbookData['styles'] {
  const styles: IWorkbookData['styles'] = {}
  for (const sheet of document.workbook.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      const id = styleKey(cell.style)
      if (!id || styles[id]) continue
      styles[id] = toUniverStyle(cell.style)
    }
  }
  return styles
}

function styleKey(style: CellStyle | undefined): string | null {
  if (!style || Object.keys(style).length === 0) return null
  return `style-${stableId(JSON.stringify(Object.entries(style).sort()))}`
}

function toUniverStyle(style: CellStyle | undefined): IStyleData {
  const data: IStyleData = {}
  if (!style) return data
  if (style.wrap) data.tb = style.wrap === 'wrap' ? WrapStrategy.WRAP : style.wrap === 'clip' ? WrapStrategy.CLIP : WrapStrategy.OVERFLOW
  if (style.bold) data.bl = BooleanNumber.TRUE
  if (style.italic) data.it = BooleanNumber.TRUE
  if (style.underline)
    data.ul = { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE }
  if (style.fontSize) data.fs = style.fontSize
  if (style.textColor) data.cl = { rgb: style.textColor }
  if (style.fillColor) data.bg = { rgb: style.fillColor }
  if (style.border) {
    const border = { s: BorderStyleTypes.THIN, cl: { rgb: CELL_BORDER_COLOR } }
    data.bd = { t: border, r: border, b: border, l: border }
  }
  if (style.align) data.ht = horizontalAlign(style.align)
  if (style.numberFormat)
    data.n = {
      pattern: numberFormatPattern(style.numberFormat, style.decimals),
    }
  data.vt = VerticalAlign.MIDDLE
  return data
}

function horizontalAlign(
  align: NonNullable<CellStyle['align']>,
): HorizontalAlign {
  if (align === 'center') return HorizontalAlign.CENTER
  if (align === 'right') return HorizontalAlign.RIGHT
  return HorizontalAlign.LEFT
}

function numberFormatPattern(
  format: NonNullable<CellStyle['numberFormat']>,
  decimals?: number,
): string {
  const zeros =
    decimals !== undefined && decimals > 0
      ? `.${'0'.repeat(Math.min(decimals, 8))}`
      : ''
  if (format === 'currency')
    return decimals === undefined ? '$#,##0.00' : `$#,##0${zeros}`
  if (format === 'percent')
    return decimals === undefined ? '0.00%' : `0${zeros}%`
  if (format === 'date') return 'yyyy-mm-dd'
  if (format === 'number')
    return decimals === undefined ? '#,##0.########' : `#,##0${zeros}`
  return '@'
}

function parseCellKey(key: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)([1-9][0-9]*)$/i.exec(key.trim())
  if (!match) return null
  return {
    row: Number(match[2]) - 1,
    column:
      match[1]
        .toUpperCase()
        .split('')
        .reduce((total, letter) => {
          return total * 26 + letter.charCodeAt(0) - 64
        }, 0) - 1,
  }
}

function toCellKey(row: number, column: number): string | null {
  if (
    !Number.isInteger(row) ||
    !Number.isInteger(column) ||
    row < 0 ||
    column < 0
  ) {
    return null
  }
  let value = column + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return `${label}${row + 1}`
}

function positiveInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isInteger(value) && value > 0
    ? value
    : undefined
}

function optionalNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined
}

function isRecord(input: unknown): input is Record<string, unknown> {
  return typeof input === 'object' && input !== null && !Array.isArray(input)
}

function stableId(input: string): string {
  let hash = 0
  for (let index = 0; index < input.length; index += 1) {
    hash = (hash * 31 + input.charCodeAt(index)) >>> 0
  }
  return hash.toString(36)
}
