import { normalizeNamedRanges } from './rangeNames'
import { workbookLinkedTables } from './linkedTables'
import type {
  AgentLogCellChange,
  AgentLogEntry,
  AgentDraftMetadata,
  CellStyle,
  CellValueKind,
  PureSheetsDocument,
  SheetCell,
  WorkbookChart,
  WorkbookSheet,
} from '../types'
import {
  documentFromUniverSnapshot,
  withFreshUniverSnapshot,
  type PureSheetsUniverSnapshot,
} from './univerAdapter'
import {
  clampFontSize,
  isForbiddenKey,
  isSafeCellKey,
  isSafeRangeToken,
  MAX_PERSISTED_ROWS,
  MAX_SHEET_COLUMNS,
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

const HTML_DATA_SCRIPT_ID = 'puresheets-workbook-data'

export function loadSheetsDocument(raw: string): PureSheetsDocument {
  const parsed = JSON.parse(raw) as unknown
  return validateDocument(parsed)
}

export function serializeSheetsDocument(document: PureSheetsDocument): string {
  return `${JSON.stringify({ ...touchDocument(document), linkedTables: workbookLinkedTables(document) }, null, 2)}\n`
}

export function loadSheetsHtmlDocument(raw: string): PureSheetsDocument {
  const scriptMatch = new RegExp(
    `<script[^>]*id=["']${HTML_DATA_SCRIPT_ID}["'][^>]*>([\\s\\S]*?)<\\/script>`,
    'i',
  ).exec(raw)
  if (!scriptMatch?.[1]) {
    throw new Error('PureSheets HTML file is missing embedded workbook data.')
  }
  return loadSheetsDocument(scriptMatch[1])
}

export function serializeSheetsHtmlDocument(
  document: PureSheetsDocument,
): string {
  const safeJson = JSON.stringify({ ...touchDocument(document), linkedTables: workbookLinkedTables(document) }, null, 2)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
  return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <title>${escapeHtml(document.metadata.title)}</title>
  </head>
  <body>
    <main>
      <h1>${escapeHtml(document.metadata.title)}</h1>
      <p>Portable PureSheets workbook. Open this file in PureDesktop to edit it.</p>
    </main>
    <script type="application/json" data-linked-table-source id="${HTML_DATA_SCRIPT_ID}">${safeJson}</script>
  </body>
</html>
`
}

export function isSheetsHtmlPath(path: string): boolean {
  return path.toLowerCase().endsWith('.sheets.html')
}

export function validateDocument(input: unknown): PureSheetsDocument {
  if (!isRecord(input))
    throw new Error('PureSheets file must be a JSON object.')
  if (input.app !== 'PureSheets')
    throw new Error('This is not a PureSheets document.')
  if (input.version !== 1)
    throw new Error('Unsupported PureSheets document version.')
  if (!isRecord(input.workbook))
    throw new Error('PureSheets workbook data is missing.')
  const workbook = input.workbook
  if (!Array.isArray(workbook.sheets) || workbook.sheets.length === 0) {
    throw new Error('PureSheets workbook must contain at least one sheet.')
  }
  const sheets = workbook.sheets.map(normalizeSheet)
  const activeSheetId =
    typeof workbook.activeSheetId === 'string' &&
    sheets.some(sheet => sheet.id === workbook.activeSheetId)
      ? workbook.activeSheetId
      : sheets[0].id
  const metadata = isRecord(input.metadata) ? input.metadata : {}
  const engine = isRecord(input.engine) ? input.engine : {}
  const preservedSnapshotId =
    isRecord(engine.snapshot) &&
    typeof engine.snapshot.id === 'string' &&
    engine.snapshot.id.trim() &&
    !isForbiddenKey(engine.snapshot.id)
      ? engine.snapshot.id
      : undefined
  const fallbackDocument: PureSheetsDocument = {
    app: 'PureSheets',
    version: 1,
    engine: {
      name: 'univer',
      version: typeof engine.version === 'string' ? engine.version : undefined,
      ...(preservedSnapshotId ? { snapshot: { id: preservedSnapshotId } } : {}),
    },
    metadata: {
      title: typeof metadata.title === 'string' ? metadata.title : 'Untitled',
      createdAt:
        typeof metadata.createdAt === 'string'
          ? metadata.createdAt
          : new Date().toISOString(),
      updatedAt:
        typeof metadata.updatedAt === 'string'
          ? metadata.updatedAt
          : new Date().toISOString(),
    },
    workbook: {
      activeSheetId,
      sheets,
      agentLog: normalizeAgentLog(workbook.agentLog),
      namedRanges: normalizeNamedRanges(workbook.namedRanges, sheets),
    },
  }
  if (isRecord(engine.snapshot) && isRecord(engine.snapshot.sheets)) {
    try {
      return documentFromUniverSnapshot(
        engine.snapshot as unknown as PureSheetsUniverSnapshot,
        fallbackDocument,
      )
    } catch {
      // Legacy PureSheets files may contain partial or stale engine data.
      // The schema workbook remains the compatibility source in that case.
    }
  }
  return withFreshUniverSnapshot(fallbackDocument)
}

function normalizeSheet(input: unknown): WorkbookSheet {
  if (!isRecord(input)) throw new Error('Sheet must be an object.')
  if (typeof input.id !== 'string' || typeof input.name !== 'string') {
    throw new Error('Sheet id and name are required.')
  }
  if (
    typeof input.rowCount !== 'number' ||
    typeof input.columnCount !== 'number' ||
    !Number.isInteger(input.rowCount) ||
    !Number.isInteger(input.columnCount) ||
    input.rowCount < 1 ||
    input.columnCount < 1
  ) {
    throw new Error('Sheet dimensions are required.')
  }
  if (!isRecord(input.cells)) throw new Error('Sheet cells must be an object.')
  const rowCount = Math.min(input.rowCount, MAX_PERSISTED_ROWS)
  const columnCount = Math.min(input.columnCount, MAX_SHEET_COLUMNS)
  const rawFrozenRows = optionalInteger(input.frozenRows)
  const rawFrozenCols = optionalInteger(input.frozenColumns)
  const rawFilterRow = optionalInteger(input.filterRow)
  const rawFilterCol = optionalInteger(input.filterColumn)
  return {
    id: input.id,
    name: input.name,
    rowCount,
    columnCount,
    cells: normalizeCells(input.cells),
    columnWidths: normalizeNumberRecord(input.columnWidths),
    rowHeights: normalizeNumberRecord(input.rowHeights),
    frozenRows:
      rawFrozenRows !== undefined
        ? Math.min(rowCount, Math.max(0, rawFrozenRows))
        : undefined,
    frozenColumns:
      rawFrozenCols !== undefined
        ? Math.min(columnCount, Math.max(0, rawFrozenCols))
        : undefined,
    merges: normalizeMerges(input.merges),
    tabColor: normalizeHexColor(input.tabColor),
    hidden: input.hidden === true ? true : undefined,
    showGridlines: input.showGridlines === false ? false : undefined,
    hiddenRows: normalizeIndexList(input.hiddenRows),
    hiddenColumns: normalizeIndexList(input.hiddenColumns),
    filterRow:
      rawFilterRow !== undefined && rawFilterRow >= 0 && rawFilterRow < rowCount
        ? rawFilterRow
        : undefined,
    filterColumn:
      rawFilterCol !== undefined &&
      rawFilterCol >= 0 &&
      rawFilterCol < columnCount
        ? rawFilterCol
        : undefined,
    filterHasHeader:
      typeof input.filterHasHeader === 'boolean'
        ? input.filterHasHeader
        : undefined,
    filterQuery: normalizeFilterQuery(input.filterQuery),
    columnFilters: normalizeColumnFilters(input.columnFilters),
    pivot: normalizePivotDefinition(input.pivot),
    charts: normalizeCharts(input.charts),
    comments: normalizeComments(input.comments),
    validations: normalizeValidations(input.validations),
    protection: normalizeProtection(input.protection),
    conditionalFormats: normalizeConditionalFormats(input.conditionalFormats),
    agentDraft: normalizeAgentDraft(input.agentDraft),
  }
}

function normalizeCells(
  input: Record<string, unknown>,
): Record<string, SheetCell> {
  const cells: Record<string, SheetCell> = {}
  for (const [key, value] of Object.entries(input)) {
    if (!isSafeCellKey(key)) continue
    if (!isRecord(value)) continue
    if (typeof value.value !== 'string') continue
    const kind = normalizeCellKind(value.kind)
    cells[key.trim().toUpperCase()] = {
      value: value.value,
      kind,
      style: normalizeCellStyle(value.style),
      ...(typeof value.hyperlink === 'string' && value.hyperlink.trim()
        ? { hyperlink: value.hyperlink.trim() }
        : {}),
      ...(value.checkbox === true ? { checkbox: true } : {}),
    }
  }
  return cells
}

function normalizeCellKind(value: unknown): CellValueKind {
  return value === 'number' ||
    value === 'formula' ||
    value === 'blank' ||
    value === 'text'
    ? value
    : 'text'
}

function normalizeCellStyle(input: unknown): CellStyle | undefined {
  if (!isRecord(input)) return undefined
  const style: CellStyle = {}
  if (typeof input.bold === 'boolean') style.bold = input.bold
  if (typeof input.italic === 'boolean') style.italic = input.italic
  if (typeof input.underline === 'boolean') style.underline = input.underline
  if (typeof input.strikethrough === 'boolean')
    style.strikethrough = input.strikethrough
  if (typeof input.fontFamily === 'string' && input.fontFamily.trim())
    style.fontFamily = input.fontFamily.trim()
  if (['wrap', 'clip', 'overflow'].includes(input.wrap as string)) style.wrap = input.wrap as CellStyle['wrap']
  const fontSize = clampFontSize(input.fontSize)
  if (fontSize !== undefined) style.fontSize = fontSize
  const textColor = normalizeHexColor(input.textColor)
  if (textColor) style.textColor = textColor
  const fillColor = normalizeHexColor(input.fillColor)
  if (fillColor) style.fillColor = fillColor
  // `border` was silently dropped here for legacy files without an embedded
  // engine snapshot (the snapshot path preserved it) — a legacy-load data
  // loss the round-trip baseline could not see. Keep it, and `decimals`.
  if (typeof input.border === 'boolean') style.border = input.border
  if (
    typeof input.decimals === 'number' &&
    Number.isInteger(input.decimals) &&
    input.decimals >= 0
  ) {
    style.decimals = Math.min(input.decimals, 8)
  }
  if (
    input.align === 'left' ||
    input.align === 'center' ||
    input.align === 'right'
  ) {
    style.align = input.align
  }
  if (
    input.numberFormat === 'text' ||
    input.numberFormat === 'number' ||
    input.numberFormat === 'currency' ||
    input.numberFormat === 'accounting' ||
    input.numberFormat === 'percent' ||
    input.numberFormat === 'date' ||
    input.numberFormat === 'time' ||
    input.numberFormat === 'scientific'
  ) {
    style.numberFormat = input.numberFormat
  }
  if (
    input.currencySymbol === '$' ||
    input.currencySymbol === '€' ||
    input.currencySymbol === '£' ||
    input.currencySymbol === '¥'
  ) {
    style.currencySymbol = input.currencySymbol
  }
  if (
    input.borderStyle === 'thin' ||
    input.borderStyle === 'medium' ||
    input.borderStyle === 'double'
  ) {
    style.borderStyle = input.borderStyle
  }
  if (typeof input.borderColor === 'string' && input.borderColor.trim()) {
    style.borderColor = input.borderColor.trim()
  }
  // Carried from Excel imports: the exact format code, partial borders, vertical alignment.
  if (typeof input.numberPattern === 'string' && input.numberPattern.trim())
    style.numberPattern = input.numberPattern.slice(0, 256)
  if (Array.isArray(input.borderSides)) {
    const sides = (['top', 'right', 'bottom', 'left'] as const).filter(side =>
      (input.borderSides as unknown[]).includes(side),
    )
    if (sides.length) style.borderSides = sides
  }
  if (input.verticalAlign === 'top' || input.verticalAlign === 'middle' || input.verticalAlign === 'bottom')
    style.verticalAlign = input.verticalAlign
  return Object.keys(style).length ? style : undefined
}

function normalizeMerges(input: unknown): string[] | undefined {
  if (!Array.isArray(input)) return undefined
  const merges = input.filter(
    (merge): merge is string =>
      typeof merge === 'string' && merge.includes(':') && isSafeRangeToken(merge),
  )
  return merges.length ? merges : undefined
}

function normalizeNumberRecord(
  input: unknown,
): Record<string, number> | undefined {
  if (!isRecord(input)) return undefined
  const values: Record<string, number> = {}
  for (const [key, value] of Object.entries(input)) {
    if (!/^(?:0|[1-9]\d*)$/.test(key)) continue
    const idx = Number(key)
    if (!Number.isSafeInteger(idx) || idx >= 1_048_576) continue
    if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) continue
    values[key] = Math.min(2000, Math.max(1, value))
  }
  return Object.keys(values).length ? values : undefined
}

function optionalInteger(input: unknown): number | undefined {
  return typeof input === 'number' && Number.isInteger(input)
    ? input
    : undefined
}

const VALID_CHART_TYPES = new Set<WorkbookChart['type']>([
  'bar',
  'stacked-bar',
  'horizontal-bar',
  'line',
  'area',
  'combo',
  'pie',
  'donut',
  'scatter',
])

function normalizeCharts(input: unknown): WorkbookChart[] | undefined {
  if (!Array.isArray(input)) return undefined
  const charts: WorkbookChart[] = input.flatMap(item => {
    if (!isRecord(item)) return []
    const chartType = item.type
    if (
      typeof item.id !== 'string' ||
      typeof item.title !== 'string' ||
      typeof chartType !== 'string' ||
      !VALID_CHART_TYPES.has(chartType as WorkbookChart['type']) ||
      typeof item.range !== 'string' ||
      !Array.isArray(item.labels) ||
      !Array.isArray(item.values)
    ) {
      return []
    }
    const values = item.values.filter(
      (value): value is number =>
        typeof value === 'number' && Number.isFinite(value),
    )
    if (values.length !== item.values.length) return []
    const series = Array.isArray(item.series)
      ? item.series.flatMap(s => {
          if (!isRecord(s) || typeof s.name !== 'string' || !Array.isArray(s.values))
            return []
          const nums = s.values.filter(
            (v): v is number => typeof v === 'number' && Number.isFinite(v),
          )
          if (nums.length !== s.values.length) return []
          return [
            {
              name: s.name,
              values: nums,
              color: typeof s.color === 'string' ? s.color : undefined,
            },
          ]
        })
      : undefined
    return [
      {
        id: item.id,
        title: item.title,
        type: chartType as WorkbookChart['type'],
        range: item.range,
        labels: item.labels.map(label => String(label)),
        values,
        ...(series && series.length ? { series } : {}),
        legend:
          item.legend === 'right' ||
          item.legend === 'bottom' ||
          item.legend === 'none'
            ? item.legend
            : undefined,
        showAxes: typeof item.showAxes === 'boolean' ? item.showAxes : undefined,
        xLabel: typeof item.xLabel === 'string' ? item.xLabel : undefined,
        yLabel: typeof item.yLabel === 'string' ? item.yLabel : undefined,
      },
    ]
  })
  return charts.length ? charts : undefined
}

function normalizeAgentDraft(input: unknown): AgentDraftMetadata | undefined {
  if (!isRecord(input)) return undefined
  if (
    typeof input.sourceSheetId !== 'string' ||
    typeof input.sourceSheetName !== 'string' ||
    typeof input.agentRunId !== 'string' ||
    typeof input.agentName !== 'string' ||
    typeof input.createdAt !== 'string' ||
    typeof input.selectedSourceRange !== 'string' ||
    typeof input.prompt !== 'string' ||
    !isQaStatus(input.qaStatus)
  ) {
    return undefined
  }
  return {
    sourceSheetId: input.sourceSheetId,
    sourceSheetName: input.sourceSheetName,
    agentRunId: input.agentRunId,
    agentName: input.agentName,
    createdAt: input.createdAt,
    selectedSourceRange: input.selectedSourceRange,
    prompt: input.prompt,
    qaStatus: input.qaStatus,
  }
}

function normalizeAgentLog(input: unknown): AgentLogEntry[] | undefined {
  if (!Array.isArray(input)) return undefined
  const entries = input.flatMap((item): AgentLogEntry[] => {
    if (!isRecord(item)) return []
    if (
      typeof item.id !== 'string' ||
      typeof item.at !== 'string' ||
      typeof item.agentName !== 'string' ||
      typeof item.tool !== 'string' ||
      typeof item.summary !== 'string' ||
      typeof item.sheetId !== 'string' ||
      typeof item.sheetName !== 'string'
    ) {
      return []
    }
    const cells = Array.isArray(item.cells)
      ? item.cells.flatMap((cell): AgentLogCellChange[] => {
          if (!isRecord(cell)) return []
          if (
            typeof cell.cell !== 'string' ||
            !isSafeCellKey(cell.cell) ||
            typeof cell.after !== 'string' ||
            (cell.kind !== 'text' &&
              cell.kind !== 'number' &&
              cell.kind !== 'formula' &&
              cell.kind !== 'blank')
          ) {
            return []
          }
          return [
            {
              cell: cell.cell.trim().toUpperCase(),
              after: cell.after,
              kind: cell.kind,
              ...(typeof cell.before === 'string'
                ? { before: cell.before }
                : {}),
            },
          ]
        })
      : undefined
    return [
      {
        id: item.id,
        at: item.at,
        agentName: item.agentName,
        tool: item.tool,
        summary: item.summary,
        sheetId: item.sheetId,
        sheetName: item.sheetName,
        ...(typeof item.range === 'string' ? { range: item.range } : {}),
        ...(cells?.length ? { cells } : {}),
      },
    ]
  })
  return entries.length ? entries : undefined
}

function isQaStatus(
  input: unknown,
): input is AgentDraftMetadata['qaStatus'] {
  return (
    input === 'pending' ||
    input === 'partially-approved' ||
    input === 'approved' ||
    input === 'rejected' ||
    input === 'discarded'
  )
}

function touchDocument(document: PureSheetsDocument): PureSheetsDocument {
  return withFreshUniverSnapshot({
    ...document,
    metadata: {
      ...document.metadata,
      updatedAt: new Date().toISOString(),
    },
  })
}

function isRecord(input: unknown): input is Record<string, unknown> {
  return typeof input === 'object' && input !== null && !Array.isArray(input)
}

function escapeHtml(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
