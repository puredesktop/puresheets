export type CellValueKind = 'text' | 'number' | 'formula' | 'blank'

export interface CellStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  wrap?: 'wrap' | 'clip' | 'overflow'
  fontSize?: number
  textColor?: string
  fillColor?: string
  border?: boolean
  align?: 'left' | 'center' | 'right'
  numberFormat?: 'text' | 'number' | 'currency' | 'percent' | 'date'
  /**
   * Fixed fraction digits for number/currency/percent formats (the decimal
   * increase/decrease controls). Unset = the format's default (currency and
   * percent: 2; plain number: as entered).
   */
  decimals?: number
}

export interface SheetCell {
  value: string
  kind: CellValueKind
  style?: CellStyle
}

export type WorkbookChartType =
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'donut'
  | 'scatter'

export interface WorkbookChart {
  id: string
  title: string
  type: WorkbookChartType
  range: string
  /** Category labels; for scatter charts these hold the numeric x values. */
  labels: string[]
  values: number[]
  /** Chart editor options (Phase S4); absent = defaults. */
  legend?: 'right' | 'bottom' | 'none'
  showAxes?: boolean
  xLabel?: string
  yLabel?: string
}

/** One data-validation rule over a range (Phase S3). */
export interface ValidationRule {
  id: string
  range: string
  kind: 'list' | 'number-range' | 'date'
  /** Allowed values for `list` rules. */
  options?: string[]
  /** Inclusive bounds for `number-range` rules. */
  min?: number
  max?: number
}

/** One conditional-formatting rule over a range (Phase S3). */
export interface ConditionalFormatRule {
  id: string
  range: string
  condition: {
    kind: 'greater' | 'less' | 'equal' | 'between' | 'text-contains'
    value: string
    /** Upper bound for `between`. */
    value2?: string
  }
  style: CellStyle
}

/**
 * Sheet protection (Phase S3). Locked ranges are enforced at the model layer:
 * value mutations skip locked cells — including agent tool writes, which
 * skip locked cells and report the skip.
 */
export interface SheetProtection {
  lockedRanges: string[]
}

export interface WorkbookSheet {
  id: string
  name: string
  rowCount: number
  columnCount: number
  cells: Record<string, SheetCell>
  columnWidths?: Record<string, number>
  rowHeights?: Record<string, number>
  frozenRows?: number
  frozenColumns?: number
  filterRow?: number
  filterColumn?: number
  /** Whether the first row is a header that stays pinned/visible while
   *  filtering. Defaults to true (Excel-style); false filters row 0 too. */
  filterHasHeader?: boolean
  /** Optional contains-filter applied to the filter column (Phase S3). */
  filterQuery?: string
  charts?: WorkbookChart[]
  /** Cell comments/notes, keyed by cell (Phase S3). */
  comments?: Record<string, string>
  validations?: ValidationRule[]
  protection?: SheetProtection
  conditionalFormats?: ConditionalFormatRule[]
  agentDraft?: AgentDraftMetadata
}

/**
 * Metadata on a legacy "agent draft" sheet. Nothing creates these any more
 * (the in-app QA workflow was removed); the type survives so workbooks saved
 * by older versions still load, and draft sheets in them stay hidden from
 * the tab bar.
 */
export interface AgentDraftMetadata {
  sourceSheetId: string
  sourceSheetName: string
  agentRunId: string
  agentName: string
  createdAt: string
  selectedSourceRange: string
  prompt: string
  qaStatus:
    | 'pending'
    | 'partially-approved'
    | 'approved'
    | 'rejected'
    | 'discarded'
}

export interface AgentCellChange {
  cell: string
  value: string
  kind: CellValueKind
}

export interface PureSheetsDocument {
  app: 'PureSheets'
  version: 1
  engine: {
    name: 'univer'
    version?: string
    snapshot?: unknown
  }
  metadata: {
    title: string
    createdAt: string
    updatedAt: string
  }
  workbook: {
    activeSheetId: string
    sheets: WorkbookSheet[]
    /**
     * The agent log: every agent write (cell changes, charts, sheets,
     * comments) recorded with before/after, newest first. The log — not a
     * review queue — is the accountability mechanism for agent tool writes;
     * approval gates are the shell's permissions.
     */
    agentLog?: AgentLogEntry[]
  }
}

export interface AgentLogCellChange {
  cell: string
  before?: string
  after: string
  kind: CellValueKind
}

export interface AgentLogEntry {
  id: string
  at: string
  agentName: string
  tool: string
  summary: string
  sheetId: string
  sheetName: string
  range?: string
  cells?: AgentLogCellChange[]
}

export interface CellPosition {
  row: number
  column: number
}

export interface PureSheetsSettings {
  lastFilePath?: string
}
