import { normalizeNamedRanges, parseQualifiedRange, qualifiedRange } from './rangeNames'
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
  clampFontSize,
  isForbiddenKey,
  MAX_PERSISTED_ROWS,
  MAX_SHEET_COLUMNS,
  MAX_SHEET_ROWS,
  normalizeColumnFilters,
  normalizeComments,
  normalizeConditionalFormats,
  normalizeFilterQuery,
  normalizeHexColor,
  normalizeIndexList,
  normalizePivotDefinition,
  normalizeProtection,
  normalizeValidations,
} from './sheetMetadataNormalize'
import { patternOfStyle, styleFromPattern } from './numberFormats'
import type {
  CellStyle,
  PureSheetsDocument,
  SheetCell,
  WorkbookSheet,
} from '../types'

export const DEFINED_NAMES_RESOURCE = 'SHEET_DEFINED_NAME_PLUGIN'

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

let fallbackUnitCounter = 1

export function createWorkbookUnitId(): string {
  if (typeof globalThis.crypto?.randomUUID === 'function') {
    return `puresheets-${globalThis.crypto.randomUUID()}`
  }
  return `puresheets-${Date.now().toString(36)}-${fallbackUnitCounter++}`
}

function workbookUnitId(document: PureSheetsDocument): string {
  const existing = document.engine.snapshot
  if (
    isRecord(existing) &&
    typeof existing.id === 'string' &&
    !isForbiddenKey(existing.id) &&
    existing.id.trim()
  ) {
    return existing.id
  }
  return createWorkbookUnitId()
}

interface StyleRegistry {
  styles: NonNullable<IWorkbookData['styles']>
  idByCanonical: Map<string, string>
  nextId: number
}

function createStyleRegistry(): StyleRegistry {
  return { styles: {}, idByCanonical: new Map(), nextId: 1 }
}

function registerStyle(
  registry: StyleRegistry,
  style: CellStyle | undefined,
): string | null {
  if (!style) return null
  const entries = Object.entries(style)
    .filter(([, value]) => value !== undefined)
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
  if (entries.length === 0) return null
  const canonical = JSON.stringify(entries)
  let id = registry.idByCanonical.get(canonical)
  if (!id) {
    id = `style-${registry.nextId++}`
    registry.idByCanonical.set(canonical, id)
    registry.styles[id] = toUniverStyle(style)
  }
  return id
}

export function createUniverSnapshot(
  document: PureSheetsDocument,
): PureSheetsUniverSnapshot {
  const registry = createStyleRegistry()
  const sheets = Object.fromEntries(
    document.workbook.sheets.map(sheet => [
      sheet.id,
      createUniverSheetSnapshot(sheet, registry),
    ]),
  )

  return {
    id: workbookUnitId(document),
    name: document.metadata.title,
    appVersion: UNIVER_APP_VERSION,
    locale: LocaleType.EN_US,
    defaultStyle: CELL_DEFAULT_STYLE,
    styles: registry.styles,
    sheetOrder: document.workbook.sheets.map(sheet => sheet.id),
    sheets,
    resources: [{
      name: DEFINED_NAMES_RESOURCE,
      data: JSON.stringify(Object.fromEntries((document.workbook.namedRanges ?? []).flatMap((entry, id) => {
        const sheet = document.workbook.sheets.find(sheet => sheet.id === entry.sheetId)
        return sheet ? [[String(id), {
          id: String(id),
          name: entry.name,
          formulaOrRefString: qualifiedRange(sheet.name, entry.range),
          localSheetId: 'AllDefaultWorkbook',
        }]] : []
      }))),
    }],
    custom: {
      puresheets: {
        activeSheetId: document.workbook.activeSheetId,
        updatedAt: document.metadata.updatedAt,
        namedRanges: document.workbook.namedRanges ?? null,
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
  const sheetOrder = (
    saved.sheetOrder?.length
      ? saved.sheetOrder
      : Object.keys(saved.sheets ?? {})
  ).filter(sheetId => typeof sheetId === 'string' && !isForbiddenKey(sheetId))
  const fallbackSheetsById = new Map(
    fallback?.workbook.sheets.map(sheet => [sheet.id, sheet]) ?? [],
  )
  const sheets = sheetOrder.flatMap(sheetId => {
    const sheet =
      saved.sheets &&
      !isForbiddenKey(sheetId) &&
      Object.hasOwn(saved.sheets, sheetId)
        ? saved.sheets[sheetId]
        : undefined
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
  const baseNamedRanges =
    'namedRanges' in custom
      ? normalizeNamedRanges(custom.namedRanges, sheets)
      : normalizeNamedRanges(fallback?.workbook.namedRanges, sheets)
  let namedRanges = baseNamedRanges
  const baseByName = new Map(
    (baseNamedRanges ?? []).map(entry => [entry.name.toLowerCase(), entry]),
  )
  const resource = snapshot.resources?.find(entry => entry.name === DEFINED_NAMES_RESOURCE)
  if (resource?.data) {
    try {
      const entries: unknown = JSON.parse(resource.data)
      if (isRecord(entries)) {
        namedRanges = normalizeNamedRanges(Object.values(entries).flatMap(entry => {
          if (!isRecord(entry) || typeof entry.name !== 'string' || typeof entry.formulaOrRefString !== 'string' ||
            (entry.localSheetId && entry.localSheetId !== 'AllDefaultWorkbook')) return []
          const ref = parseQualifiedRange(entry.formulaOrRefString)
          const sheet = ref && sheets.find(sheet => sheet.name.toLowerCase() === ref.sheet.toLowerCase())
          if (!ref || !sheet) return []
          const prev = baseByName.get(entry.name.trim().toLowerCase())
          return [{
            id: prev?.id ?? `nr-${entry.name.trim().toLowerCase()}`,
            name: entry.name,
            sheetId: sheet.id,
            range: ref.range,
            ...(prev?.comment ? { comment: prev.comment } : {}),
          }]
        }), sheets)
      }
    } catch {
      // Keep the document's definitions if the editor did not provide a usable name resource.
    }
  }
  const now = new Date().toISOString()
  const preservedUnitId =
    typeof saved.id === 'string' &&
    !isForbiddenKey(saved.id) &&
    saved.id.trim()
      ? saved.id
      : undefined

  return withFreshUniverSnapshot({
    app: 'PureSheets',
    version: 1,
    engine: {
      name: 'univer',
      version: UNIVER_APP_VERSION,
      ...(preservedUnitId ? { snapshot: { id: preservedUnitId } } : {}),
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
      namedRanges,
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

/**
 * Univer's default cell text (11pt black Arial) reads heavy: 10pt in a softened
 * ink instead, so a wrapped cell's two lines sit comfortably in its row. Arial
 * stays: it is hinted for 1x screens, where a canvas renders other faces soft.
 * Set on the workbook and each sheet (the renderer reads the sheet's); never
 * written into cell styles, so it is not saved into documents, and a cell that
 * sets its own font, size or colour keeps it.
 */
const CELL_DEFAULT_STYLE = { ff: 'Arial', fs: 10, cl: { rgb: '#3a404a' } }

function createUniverSheetSnapshot(
  sheet: WorkbookSheet,
  registry: StyleRegistry,
): Partial<IWorksheetData> {
  return {
    id: sheet.id,
    name: sheet.name,
    defaultStyle: CELL_DEFAULT_STYLE,
    tabColor: sheet.tabColor ?? '',
    hidden: sheet.hidden ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    freeze: {
      startRow: sheet.frozenRows ? sheet.frozenRows : -1,
      startColumn: sheet.frozenColumns ? sheet.frozenColumns : -1,
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
    mergeData: (sheet.merges ?? []).flatMap(merge => {
      const [start, end] = merge.split(':').map(parseCellKey)
      return start && end
        ? [{ startRow: start.row, startColumn: start.column, endRow: end.row, endColumn: end.column }]
        : []
    }),
    cellData: collectCellData(sheet, registry),
    rowData: buildRowData(sheet),
    columnData: buildColumnData(sheet),
    rowHeader: { width: ROW_HEADER_WIDTH },
    columnHeader: { height: COLUMN_HEADER_HEIGHT },
    showGridlines:
      sheet.showGridlines === false ? BooleanNumber.FALSE : BooleanNumber.TRUE,
    rightToLeft: BooleanNumber.FALSE,
    custom: {
      puresheets: {
        // Keys are ALWAYS written (null when unset) so the restore path can
        // distinguish "cleared by the user" from "legacy snapshot that never
        // carried the field" (which falls back).
        filterRow: sheet.filterRow ?? null,
        filterColumn: sheet.filterColumn ?? null,
        filterHasHeader: sheet.filterHasHeader ?? null,
        filterQuery: sheet.filterQuery ?? null,
        columnFilters: sheet.columnFilters ?? null,
        pivot: sheet.pivot ?? null,
        hiddenRows: sheet.hiddenRows ?? null,
        hiddenColumns: sheet.hiddenColumns ?? null,
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
  const rawRowCount =
    positiveInteger(sheet.rowCount) ??
    fallback?.rowCount ??
    DEFAULT_ROW_COUNT
  const rowCount = Math.min(MAX_PERSISTED_ROWS, Math.max(1, rawRowCount))
  const rawColumnCount =
    positiveInteger(sheet.columnCount) ??
    fallback?.columnCount ??
    DEFAULT_COLUMN_COUNT
  const columnCount = Math.min(MAX_SHEET_COLUMNS, Math.max(1, rawColumnCount))
  const rawFrozenRows =
    typeof sheet.freeze?.ySplit === 'number' &&
    Number.isFinite(sheet.freeze.ySplit)
      ? Math.floor(sheet.freeze.ySplit)
      : fallback?.frozenRows
  const frozenRows =
    rawFrozenRows !== undefined
      ? Math.min(rowCount, Math.max(0, rawFrozenRows))
      : undefined
  const rawFrozenColumns =
    typeof sheet.freeze?.xSplit === 'number' &&
    Number.isFinite(sheet.freeze.xSplit)
      ? Math.floor(sheet.freeze.xSplit)
      : fallback?.frozenColumns
  const frozenColumns =
    rawFrozenColumns !== undefined
      ? Math.min(columnCount, Math.max(0, rawFrozenColumns))
      : undefined
  return {
    id,
    name:
      typeof sheet.name === 'string' && sheet.name.trim()
        ? sheet.name
        : fallback?.name ?? id,
    rowCount,
    columnCount,
    cells: cellsFromUniverSheet(sheet.cellData, styles, fallback?.cells),
    columnWidths:
      dimensionsFromUniverData(sheet.columnData, 'w', MAX_SHEET_COLUMNS - 1) ??
      fallback?.columnWidths,
    rowHeights:
      dimensionsFromUniverData(sheet.rowData, 'h', MAX_SHEET_ROWS - 1) ??
      fallback?.rowHeights,
    frozenRows,
    frozenColumns,
    merges: Array.isArray(sheet.mergeData)
      ? mergesFromUniverData(sheet.mergeData)
      : fallback?.merges,
    tabColor:
      typeof sheet.tabColor === 'string' && sheet.tabColor.trim()
        ? normalizeHexColor(sheet.tabColor) ?? fallback?.tabColor
        : fallback?.tabColor,
    hidden:
      sheet.hidden === BooleanNumber.TRUE
        ? true
        : sheet.hidden === BooleanNumber.FALSE
        ? undefined
        : fallback?.hidden,
    showGridlines:
      sheet.showGridlines === BooleanNumber.FALSE
        ? false
        : sheet.showGridlines === BooleanNumber.TRUE
        ? undefined
        : fallback?.showGridlines,
    hiddenRows:
      'hiddenRows' in custom
        ? normalizeIndexList(custom.hiddenRows)
        : fallback?.hiddenRows,
    hiddenColumns:
      'hiddenColumns' in custom
        ? normalizeIndexList(custom.hiddenColumns)
        : fallback?.hiddenColumns,
    filterRow:
      'filterRow' in custom
        ? optionalNumber(custom.filterRow)
        : fallback?.filterRow,
    filterColumn:
      'filterColumn' in custom
        ? optionalNumber(custom.filterColumn)
        : fallback?.filterColumn,
    filterHasHeader:
      'filterHasHeader' in custom
        ? typeof custom.filterHasHeader === 'boolean'
          ? custom.filterHasHeader
          : undefined
        : fallback?.filterHasHeader,
    // Phase S3 metadata: when the snapshot carries the key (even as null),
    // the snapshot wins; a missing key means a legacy snapshot → fallback.
    filterQuery:
      'filterQuery' in custom
        ? normalizeFilterQuery(custom.filterQuery)
        : fallback?.filterQuery,
    columnFilters:
      'columnFilters' in custom
        ? normalizeColumnFilters(custom.columnFilters)
        : fallback?.columnFilters,
    pivot:
      'pivot' in custom
        ? normalizePivotDefinition(custom.pivot)
        : fallback?.pivot,
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

function mergesFromUniverData(
  mergeData: NonNullable<IWorksheetData['mergeData']>,
): string[] | undefined {
  const merges = mergeData.flatMap(range => {
    const start = toCellKey(range.startRow, range.startColumn)
    const end = toCellKey(range.endRow, range.endColumn)
    return start && end && start !== end ? [`${start}:${end}`] : []
  })
  return merges.length ? merges : undefined
}

function cellsFromUniverSheet(
  cellData: IWorksheetData['cellData'] | undefined,
  styles: IWorkbookData['styles'] | undefined,
  fallbackCells?: Record<string, SheetCell>,
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
      const univerStyle =
        typeof cell.s === 'string'
          ? styles && Object.hasOwn(styles, cell.s)
            ? styles[cell.s] ?? undefined
            : undefined
          : isRecord(cell.s)
          ? (cell.s as IStyleData)
          : undefined
      let style = fromUniverStyle(univerStyle)
      // Every styled cell renders middle, so a chosen middle (an Excel import's) is kept from the model.
      const fallbackCell =
        fallbackCells && Object.hasOwn(fallbackCells, key)
          ? fallbackCells[key]
          : undefined
      if (univerStyle?.vt === VerticalAlign.MIDDLE && fallbackCell?.style?.verticalAlign === 'middle')
        style = { ...style, verticalAlign: 'middle' }
      const custom = cell.custom as {
        puresheetsKind?: unknown
        puresheetsNumberText?: unknown
        puresheetsHyperlink?: unknown
        puresheetsCheckbox?: unknown
      } | undefined
      const preservedNumber = custom?.puresheetsKind === 'number' && (
        custom.puresheetsNumberText === value ||
        // Older snapshots only marked the kind; retain literals that required text storage.
        (custom.puresheetsNumberText === undefined && value.trim() !== '' && !Number.isNaN(Number(value)) &&
          (!Number.isFinite(Number(value)) || (Number.isInteger(Number(value)) && !Number.isSafeInteger(Number(value)))))
      )
      const hyperlink =
        typeof custom?.puresheetsHyperlink === 'string' && custom.puresheetsHyperlink.trim()
          ? custom.puresheetsHyperlink
          : fallbackCells?.[key]?.hyperlink
      const checkbox =
        typeof custom?.puresheetsCheckbox === 'boolean'
          ? custom.puresheetsCheckbox
          : fallbackCells?.[key]?.checkbox
      cells[key] = {
        value,
        kind: cell.f
          ? 'formula'
          : cell.t === CellValueType.NUMBER || preservedNumber
          ? 'number'
          : value
          ? 'text'
          : 'blank',
        style,
        ...(hyperlink ? { hyperlink } : {}),
        ...(checkbox ? { checkbox } : {}),
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
  if (typeof style.st === 'object' ? style.st?.s : style.st) {
    cellStyle.strikethrough = true
  }
  if (typeof style.ff === 'string' && style.ff.trim() && style.ff !== 'Arial') {
    cellStyle.fontFamily = style.ff
  }
  const fontSize = clampFontSize(style.fs)
  if (fontSize !== undefined) cellStyle.fontSize = fontSize
  const textColor = normalizeHexColor(style.cl?.rgb)
  if (textColor) cellStyle.textColor = textColor
  const fillColor = normalizeHexColor(style.bg?.rgb)
  if (fillColor) cellStyle.fillColor = fillColor
  const activeBorders = BORDER_SIDES.filter(
    ([key]) => style.bd?.[key] && style.bd[key]?.s !== BorderStyleTypes.NONE,
  )
  const sides = activeBorders.map(([, side]) => side)
  if (sides.length === 4) cellStyle.border = true
  else if (sides.length) cellStyle.borderSides = sides
  const firstBorder = activeBorders[0] ? style.bd?.[activeBorders[0][0]] : undefined
  if (firstBorder?.s === BorderStyleTypes.DOUBLE) cellStyle.borderStyle = 'double'
  else if (firstBorder?.s === BorderStyleTypes.MEDIUM) cellStyle.borderStyle = 'medium'
  if (
    typeof firstBorder?.cl?.rgb === 'string' &&
    firstBorder.cl.rgb.toLowerCase() !== CELL_BORDER_COLOR.toLowerCase()
  ) {
    cellStyle.borderColor = firstBorder.cl.rgb
  }
  if (style.ht === HorizontalAlign.CENTER) cellStyle.align = 'center'
  if (style.ht === HorizontalAlign.RIGHT) cellStyle.align = 'right'
  if (style.ht === HorizontalAlign.LEFT) cellStyle.align = 'left'
  if (style.vt === VerticalAlign.TOP) cellStyle.verticalAlign = 'top'
  if (style.vt === VerticalAlign.BOTTOM) cellStyle.verticalAlign = 'bottom'
  const pattern = style.n?.pattern
  if (typeof pattern === 'string') {
    // A code PureSheets cannot name ('General') still keeps its family 'number', as before.
    Object.assign(cellStyle, styleFromPattern(pattern) ?? { numberFormat: 'number' })
  }
  return Object.keys(cellStyle).length ? cellStyle : undefined
}

function dimensionsFromUniverData(
  data: IWorksheetData['rowData'] | IWorksheetData['columnData'] | undefined,
  key: 'h' | 'w',
  maxIndex: number,
): Record<string, number> | undefined {
  const dimensions: Record<string, number> = {}
  for (const [index, value] of Object.entries(data ?? {})) {
    if (isForbiddenKey(index) || !/^(0|[1-9][0-9]*)$/.test(index)) continue
    const numericIndex = Number(index)
    if (
      !Number.isSafeInteger(numericIndex) ||
      numericIndex < 0 ||
      numericIndex > maxIndex
    ) {
      continue
    }
    // Native auto-fit stores its measured height in ah, separately from the
    // manual minimum h. Carry the displayed height into model-driven reloads.
    const row = value as { h?: number; ah?: number; ia?: BooleanNumber } | undefined
    const dimension = key === 'h' && row?.ia !== BooleanNumber.FALSE &&
      typeof row?.ah === 'number' && row.ah > 0 ? row.ah : value?.[key]
    if (typeof dimension === 'number' && Number.isFinite(dimension) && dimension > 0) {
      dimensions[index] = Math.min(2000, Math.max(1, dimension))
    }
  }
  return Object.keys(dimensions).length ? dimensions : undefined
}

function collectCellData(
  sheet: WorkbookSheet,
  registry: StyleRegistry,
): IWorksheetData['cellData'] {
  const rows: IWorksheetData['cellData'] = {}
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position) continue
    const row = rows[position.row] ?? (rows[position.row] = {})
    row[position.column] = createUniverCell(cell, registry)
  }
  return rows
}

function createUniverCell(
  cell: SheetCell,
  registry: StyleRegistry,
): ICellData {
  const value = cell.value
  const data: ICellData = {}
  if (cell.kind === 'formula') {
    data.f = value
  } else if (
    cell.kind === 'number' &&
    Number.isFinite(Number(value)) &&
    (!Number.isInteger(Number(value)) || Number.isSafeInteger(Number(value)))
  ) {
    data.v = Number(value)
    data.t = CellValueType.NUMBER
  } else if (cell.kind === 'number') {
    // Non-finite legacy numbers would serialize as null, and unsafe integers
    // would round through Number. Preserve the exact literal as a string.
    data.v = value
    data.t = CellValueType.STRING
    data.custom = { puresheetsKind: 'number', puresheetsNumberText: value }
  } else {
    data.v = value
    data.t = CellValueType.STRING
  }
  const styleId = registerStyle(registry, cell.style)
  if (styleId) data.s = styleId
  if (cell.hyperlink || cell.checkbox) {
    data.custom = {
      ...(data.custom as Record<string, unknown> | undefined),
      ...(cell.hyperlink ? { puresheetsHyperlink: cell.hyperlink } : {}),
      ...(cell.checkbox ? { puresheetsCheckbox: true } : {}),
    }
  }
  return data
}

function buildRowData(sheet: WorkbookSheet): IWorksheetData['rowData'] {
  const rows: Record<string, { h?: number; hd?: BooleanNumber }> = {}
  for (const [row, height] of Object.entries(sheet.rowHeights ?? {})) {
    rows[row] = { ...rows[row], h: height }
  }
  for (const row of sheet.hiddenRows ?? []) {
    rows[String(row)] = { ...rows[String(row)], hd: BooleanNumber.TRUE }
  }
  return rows
}

function buildColumnData(sheet: WorkbookSheet): IWorksheetData['columnData'] {
  const cols: Record<string, { w?: number; hd?: BooleanNumber }> = {}
  for (const [col, width] of Object.entries(sheet.columnWidths ?? {})) {
    cols[col] = { ...cols[col], w: width }
  }
  for (const col of sheet.hiddenColumns ?? []) {
    cols[String(col)] = { ...cols[String(col)], hd: BooleanNumber.TRUE }
  }
  return cols
}

function toUniverStyle(style: CellStyle | undefined): IStyleData {
  const data: IStyleData = {}
  if (!style) return data
  if (style.wrap) data.tb = style.wrap === 'wrap' ? WrapStrategy.WRAP : style.wrap === 'clip' ? WrapStrategy.CLIP : WrapStrategy.OVERFLOW
  if (style.bold) data.bl = BooleanNumber.TRUE
  if (style.italic) data.it = BooleanNumber.TRUE
  if (style.underline)
    data.ul = { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE }
  if (style.strikethrough)
    data.st = { s: BooleanNumber.TRUE }
  if (style.fontFamily) data.ff = style.fontFamily
  const fontSize = clampFontSize(style.fontSize)
  if (fontSize !== undefined) data.fs = fontSize
  const textColor = normalizeHexColor(style.textColor)
  if (textColor) data.cl = { rgb: textColor }
  const fillColor = normalizeHexColor(style.fillColor)
  if (fillColor) data.bg = { rgb: fillColor }
  const sides = style.border
    ? BORDER_SIDES
    : BORDER_SIDES.filter(([, side]) => style.borderSides?.includes(side))
  if (sides.length) {
    const borderType =
      style.borderStyle === 'double'
        ? BorderStyleTypes.DOUBLE
        : style.borderStyle === 'medium'
        ? BorderStyleTypes.MEDIUM
        : BorderStyleTypes.THIN
    const border = {
      s: borderType,
      cl: { rgb: style.borderColor ?? CELL_BORDER_COLOR },
    }
    data.bd = Object.fromEntries(sides.map(([key]) => [key, border]))
  }
  if (style.align) data.ht = horizontalAlign(style.align)
  const pattern = patternOfStyle(style)
  if (pattern) data.n = { pattern }
  data.vt =
    style.verticalAlign === 'top'
      ? VerticalAlign.TOP
      : style.verticalAlign === 'bottom'
      ? VerticalAlign.BOTTOM
      : VerticalAlign.MIDDLE
  return data
}

const BORDER_SIDES: Array<['t' | 'r' | 'b' | 'l', NonNullable<CellStyle['borderSides']>[number]]> = [
  ['t', 'top'],
  ['r', 'right'],
  ['b', 'bottom'],
  ['l', 'left'],
]

function horizontalAlign(
  align: NonNullable<CellStyle['align']>,
): HorizontalAlign {
  if (align === 'center') return HorizontalAlign.CENTER
  if (align === 'right') return HorizontalAlign.RIGHT
  return HorizontalAlign.LEFT
}

function parseCellKey(key: string): { row: number; column: number } | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/i.exec(key.trim())
  if (!match) return null
  const row = Number(match[2]) - 1
  const column =
    match[1]
      .toUpperCase()
      .split('')
      .reduce((total, letter) => {
        return total * 26 + letter.charCodeAt(0) - 64
      }, 0) - 1
  if (
    !Number.isSafeInteger(row) ||
    row < 0 ||
    row >= MAX_SHEET_ROWS ||
    column < 0 ||
    column >= MAX_SHEET_COLUMNS
  ) {
    return null
  }
  return { row, column }
}

function toCellKey(row: number, column: number): string | null {
  if (
    !Number.isInteger(row) ||
    !Number.isInteger(column) ||
    row < 0 ||
    column < 0 ||
    row >= MAX_SHEET_ROWS ||
    column >= MAX_SHEET_COLUMNS
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
