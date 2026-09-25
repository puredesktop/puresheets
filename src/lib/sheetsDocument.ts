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
  normalizeComments,
  normalizeConditionalFormats,
  normalizeFilterQuery,
  normalizeProtection,
  normalizeValidations,
} from './sheetMetadataNormalize'

const HTML_DATA_SCRIPT_ID = 'puresheets-workbook-data'

export function loadSheetsDocument(raw: string): PureSheetsDocument {
  const parsed = JSON.parse(raw) as unknown
  return validateDocument(parsed)
}

export function serializeSheetsDocument(document: PureSheetsDocument): string {
  return `${JSON.stringify(touchDocument(document), null, 2)}\n`
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
  const safeJson = JSON.stringify(touchDocument(document), null, 2)
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
    <script type="application/json" id="${HTML_DATA_SCRIPT_ID}">${safeJson}</script>
  </body>
</html>
`
}

export function isSheetsHtmlPath(path: string): boolean {
  return path.toLowerCase().endsWith('.sheets.html')
}

function validateDocument(input: unknown): PureSheetsDocument {
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
  const fallbackDocument: PureSheetsDocument = {
    app: 'PureSheets',
    version: 1,
    engine: {
      name: 'univer',
      version: typeof engine.version === 'string' ? engine.version : undefined,
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
    !Number.isInteger(input.columnCount)
  ) {
    throw new Error('Sheet dimensions are required.')
  }
  if (!isRecord(input.cells)) throw new Error('Sheet cells must be an object.')
  const rowCount = input.rowCount
  const columnCount = input.columnCount
  return {
    id: input.id,
    name: input.name,
    rowCount,
    columnCount,
    cells: normalizeCells(input.cells),
    columnWidths: normalizeNumberRecord(input.columnWidths),
    rowHeights: normalizeNumberRecord(input.rowHeights),
    frozenRows: optionalInteger(input.frozenRows),
    frozenColumns: optionalInteger(input.frozenColumns),
    filterRow: optionalInteger(input.filterRow),
    filterColumn: optionalInteger(input.filterColumn),
    filterQuery: normalizeFilterQuery(input.filterQuery),
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
  return Object.fromEntries(
    Object.entries(input).flatMap(([key, value]) => {
      if (!isRecord(value)) return []
      if (typeof value.value !== 'string') return []
      const kind = normalizeCellKind(value.kind)
      return [
        [
          key,
          {
            value: value.value,
            kind,
            style: normalizeCellStyle(value.style),
          },
        ],
      ]
    }),
  )
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
  if (['wrap', 'clip', 'overflow'].includes(input.wrap as string)) style.wrap = input.wrap as CellStyle['wrap']
  if (typeof input.fontSize === 'number' && Number.isFinite(input.fontSize)) {
    style.fontSize = input.fontSize
  }
  if (typeof input.textColor === 'string') style.textColor = input.textColor
  if (typeof input.fillColor === 'string') style.fillColor = input.fillColor
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
    input.numberFormat === 'percent' ||
    input.numberFormat === 'date'
  ) {
    style.numberFormat = input.numberFormat
  }
  return Object.keys(style).length ? style : undefined
}

function normalizeNumberRecord(
  input: unknown,
): Record<string, number> | undefined {
  if (!isRecord(input)) return undefined
  const values = Object.fromEntries(
    Object.entries(input).filter(
      (entry): entry is [string, number] =>
        typeof entry[1] === 'number' && Number.isFinite(entry[1]),
    ),
  )
  return Object.keys(values).length ? values : undefined
}

function optionalInteger(input: unknown): number | undefined {
  return typeof input === 'number' && Number.isInteger(input)
    ? input
    : undefined
}

function normalizeCharts(input: unknown): WorkbookChart[] | undefined {
  if (!Array.isArray(input)) return undefined
  const charts: WorkbookChart[] = input.flatMap(item => {
    if (!isRecord(item)) return []
    const chartType = item.type
    if (
      typeof item.id !== 'string' ||
      typeof item.title !== 'string' ||
      (chartType !== 'bar' &&
        chartType !== 'line' &&
        chartType !== 'area' &&
        chartType !== 'pie' &&
        chartType !== 'donut' &&
        chartType !== 'scatter') ||
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
    return [
      {
        id: item.id,
        title: item.title,
        type: chartType,
        range: item.range,
        labels: item.labels.map(label => String(label)),
        values,
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
            typeof cell.after !== 'string' ||
            typeof cell.kind !== 'string'
          ) {
            return []
          }
          return [
            {
              cell: cell.cell,
              after: cell.after,
              kind: cell.kind as AgentLogCellChange['kind'],
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
