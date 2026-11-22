export type CellValueKind = 'text' | 'number' | 'formula' | 'blank'

export interface CellStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  fontFamily?: string
  wrap?: 'wrap' | 'clip' | 'overflow'
  fontSize?: number
  textColor?: string
  fillColor?: string
  border?: boolean
  /**
   * Thin borders on only some sides (an imported Excel underline, say). Unset
   * with `border: true` means all four; setting `border` clears this.
   */
  borderSides?: Array<'top' | 'right' | 'bottom' | 'left'>
  borderStyle?: 'thin' | 'medium' | 'double'
  borderColor?: string
  align?: 'left' | 'center' | 'right'
  /** Vertical alignment; unset = middle. */
  verticalAlign?: 'top' | 'middle' | 'bottom'
  numberFormat?:
    | 'text'
    | 'number'
    | 'currency'
    | 'accounting'
    | 'percent'
    | 'date'
    | 'time'
    | 'scientific'
  currencySymbol?: '$' | '€' | '£' | '¥'
  /**
   * The exact number format code an Excel import carried (`#,##0.0;[Red]-#,##0.0`),
   * kept so it shows and exports as it was. `numberFormat` still names its family
   * for the toolbar; setting `numberFormat` or `decimals` clears this.
   */
  numberPattern?: string
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
  hyperlink?: string
  checkbox?: boolean
}

export type WorkbookChartType =
  | 'bar'
  | 'stacked-bar'
  | 'horizontal-bar'
  | 'line'
  | 'area'
  | 'combo'
  | 'pie'
  | 'donut'
  | 'scatter'

export interface ChartSeries {
  name: string
  values: number[]
  color?: string
}

export interface WorkbookChart {
  id: string
  title: string
  type: WorkbookChartType
  range: string
  /** Category labels; for scatter charts these hold the numeric x values. */
  labels: string[]
  /** Primary series values (kept in sync with `series[0].values` for backward compatibility). */
  values: number[]
  /** Multi-series data (Wave 5). When absent, chart renders single-series `values`. */
  series?: ChartSeries[]
  /** Chart editor options; absent = defaults. */
  legend?: 'right' | 'bottom' | 'none'
  showAxes?: boolean
  xLabel?: string
  yLabel?: string
}

/** One data-validation rule over a range. */
export interface ValidationRule {
  id: string
  range: string
  kind:
    | 'list'
    | 'number-range'
    | 'date'
    | 'text-length'
    | 'text-contains'
    | 'custom-formula'
  /** Allowed values for `list` rules. */
  options?: string[]
  /** Optional sheet/range reference backing a `list` rule (e.g. `A2:A10` or `Lists!A2:A10`). */
  sourceRange?: string
  /** Inclusive bounds for `number-range` or `text-length` rules. */
  min?: number
  max?: number
  /** Substring for `text-contains` rules. */
  textQuery?: string
  /** Custom formula for `custom-formula` rules (truthy = valid). */
  formula?: string
  /** When true, UI warns strongly before committing; default false (flag only). */
  strict?: boolean
}

/** One conditional-formatting rule over a range. */
export interface ConditionalFormatRule {
  id: string
  range: string
  condition: {
    kind:
      | 'greater'
      | 'less'
      | 'equal'
      | 'not-equal'
      | 'between'
      | 'text-contains'
      | 'blank'
      | 'not-blank'
      | 'duplicate'
      | 'top-n'
      | 'bottom-n'
      | 'color-scale'
      | 'data-bar'
      | 'custom-formula'
    value: string
    /** Upper bound for `between`. */
    value2?: string
  }
  style: CellStyle
  /** Optional 2- or 3-color gradient for `color-scale` rules. */
  colorScale?: { minColor: string; midColor?: string; maxColor: string }
  /** Optional bar accent color for `data-bar` rules. */
  dataBarColor?: string
}

/** Multi-column filter rule (Wave 3). */
export interface ColumnFilterRule {
  /** Distinct values explicitly selected; undefined = all values allowed. */
  selectedValues?: string[]
  operator?:
    | 'contains'
    | 'not-contains'
    | 'equals'
    | 'gt'
    | 'lt'
    | 'between'
    | 'blank'
    | 'not-blank'
  value?: string
  value2?: string
}

export type PivotAggregate =
  | 'sum'
  | 'count'
  | 'avg'
  | 'min'
  | 'max'
  | 'percent-of-total'

export interface PivotFieldRef {
  columnIndex: number
  label: string
}

export interface PivotValueRef {
  columnIndex: number
  label: string
  aggregate: PivotAggregate
}

export interface PivotFilterRef {
  columnIndex: number
  label: string
  selectedValues: string[]
}

/** Derived pivot-table projection owned by a destination sheet (Wave 5). */
export interface PivotDefinition {
  id: string
  sourceSheetId: string
  sourceRange: string
  rows: PivotFieldRef[]
  columns: PivotFieldRef[]
  values: PivotValueRef[]
  filters?: PivotFilterRef[]
  refreshedAt: string
  sourceSignature?: string
}

/** Workbook-level Named Range (Wave 2). */
export interface NamedRange {
  id: string
  name: string
  sheetId: string
  range: string
  comment?: string
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
  /** Merged ranges in A1 form (`A1:C1`); the top-left cell holds the value. */
  merges?: string[]
  tabColor?: string
  hidden?: boolean
  /** Whether gridlines are visible on this sheet (default true). */
  showGridlines?: boolean
  /** Explicitly hidden row indices (independent of filter rows). */
  hiddenRows?: number[]
  /** Explicitly hidden column indices. */
  hiddenColumns?: number[]
  filterRow?: number
  filterColumn?: number
  /** Whether the first row is a header that stays pinned/visible while
   *  filtering. Defaults to true (Excel-style); false filters row 0 too. */
  filterHasHeader?: boolean
  /** Optional contains-filter applied to the filter column (Phase S3). */
  filterQuery?: string
  /** Per-column smart filters (Wave 3). */
  columnFilters?: Record<string, ColumnFilterRule>
  /** Derived pivot table definition (Wave 5). */
  pivot?: PivotDefinition
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
    /** Workbook-level named ranges (Wave 2). */
    namedRanges?: NamedRange[]
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
