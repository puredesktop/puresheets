import type { PureSheetsUniverSnapshot } from '../lib/univerAdapter'
import type { CellStyle } from '../types'

/**
 * Contract between the shell and the mounted Univer editing surface.
 *
 * The Univer editor is the primary editing/rendering surface (Phase S0 of
 * MAIL_SHEETS_MASTER_PLAN.md): shell-level commands are routed through this
 * bridge first, and every bridge edit exports a fresh Univer snapshot which is
 * converted back into the `.sheets` document (see `documentFromUniverSnapshot`).
 * Callers fall back to the legacy workbook-model mutation only when the
 * surface is unavailable (tests, mount failure).
 */
export type UniverEditorBridge = {
  addSheet: (name: string) => Promise<boolean>
  applyRangeStyle: (
    range: string,
    patch: Partial<CellStyle>,
  ) => Promise<boolean>
  /** Reflect a filter on the live editor: show every row, then hide the given
   *  row indices. Display-only — never recorded as an edit. */
  applyRowVisibility: (
    hiddenRows: number[],
    rowCount: number,
  ) => Promise<boolean>
  clearRange: (range: string) => Promise<boolean>
  deleteColumn: (column: number) => Promise<boolean>
  deleteRow: (row: number) => Promise<boolean>
  deleteSheet: (sheetId: string) => Promise<boolean>
  duplicateActiveSheet: () => Promise<boolean>
  insertColumn: (column: number) => Promise<boolean>
  insertRow: (row: number) => Promise<boolean>
  pasteTabularData: (cell: string, text: string) => Promise<boolean>
  renameSheet: (sheetId: string, name: string) => Promise<boolean>
  selectSheet: (sheetId: string) => Promise<boolean>
  /** Move the editor's active selection to a range (used to highlight the
   *  column a filter applies to). Display-only — never recorded as an edit. */
  selectRange: (range: string) => Promise<boolean>
  setRangeValues: (
    range: string,
    values: UniverPasteValue[][],
  ) => Promise<boolean>
  setColumnWidth: (column: number, width: number) => Promise<boolean>
  /** Select a cell and scroll it into view in the live editor (used to make a
   *  Find match visible). Display-only — never recorded as an edit. */
  revealCell: (cell: string) => Promise<boolean>
  flushSnapshot: () => Promise<PureSheetsUniverSnapshot | null>
  setFrozenColumns: (columns: number) => Promise<boolean>
  setFrozenRows: (rows: number) => Promise<boolean>
  setRowHeight: (row: number, height: number) => Promise<boolean>
  autoFitRows: (sheetId: string, rows: number[]) => Promise<boolean>
  setCellValue: (cell: string, value: string) => Promise<boolean>
  // True while the user is actively typing in a cell (edit box open) but has
  // not yet committed with Enter/Tab.
  isCellEditing: () => boolean
  // Discard the uncommitted in-cell edit, returning the grid to committed state.
  abortEditing: () => Promise<boolean>
  // Commit (save) the uncommitted in-cell edit so its value lands in the model
  // — e.g. before Find/Replace, so the cell you're typing in is searched (#277).
  commitEditing: () => Promise<boolean>
}

export type UniverSelectionState = {
  cell: string
  range: string
  editValue: string
  displayValue: string
  preserveRange?: boolean
}

export type UniverFacadeWorksheet = {
  activate?: () => UniverFacadeWorksheet
  deleteColumns?: (columnPosition: number, howMany: number) => unknown
  deleteRows?: (rowPosition: number, howMany: number) => unknown
  getRange?: (range: string) => {
    activate?: () => unknown
    clearContent?: () => unknown
    getDisplayValue?: () => string
    getFormula?: () => string
    getRawValue?: () => unknown
    getValue?: () => unknown
    getValues?: () => unknown[][]
    setBackgroundColor?: (color: string) => unknown
    setBorder?: (type: unknown, style: unknown, color?: string) => unknown
    setFontColor?: (color: string | null) => unknown
    setFontLine?: (
      line: 'none' | 'underline' | 'line-through' | null,
    ) => unknown
    setFontSize?: (size: number | null) => unknown
    setFontStyle?: (style: 'normal' | 'italic' | null) => unknown
    setFontWeight?: (weight: 'normal' | 'bold' | null) => unknown
    setFormula?: (formula: string) => unknown
    // Univer's facade alignment vocabulary: right alignment is spelled
    // 'normal' — passing 'right' throws (`transformFacadeHorizontalAlignment`).
    setHorizontalAlignment?: (
      alignment: 'left' | 'center' | 'normal',
    ) => unknown
    setNumberFormat?: (pattern: string) => unknown
    setValue?: (value: string | number | boolean) => unknown
    setValues?: (
      values: Array<Array<string | number | boolean | null | { f: string }>>,
    ) => unknown
  }
  getSelection?: () => {
    getActiveRange?: () => {
      getRange?: () => UniverRangeLike
    } | null
    getCurrentCell?: () => {
      actualRow?: number
      actualColumn?: number
      row?: number
      column?: number
    } | null
  } | null
  getScrollState?: () => {
    offsetX?: number
    offsetY?: number
    sheetViewStartRow?: number
    sheetViewStartColumn?: number
  } | null
  getSheetId?: () => string
  getSheetName?: () => string
  scrollToCell?: (row: number, column: number, duration?: number) => unknown
  setActiveRange?: (range: unknown) => unknown
  hideRows?: (rowIndex: number, numRows?: number) => unknown
  showRows?: (rowIndex: number, numRows?: number) => unknown
  insertColumns?: (columnIndex: number, numColumns?: number) => unknown
  // Conditional formatting facade (from @univerjs/sheets-conditional-formatting).
  // All optional — the surface seeds rules best-effort and skips when absent.
  newConditionalFormattingRule?: () => UniverConditionalFormatBuilder
  addConditionalFormattingRule?: (rule: unknown) => unknown
  getConditionalFormattingRules?: () => Array<{ cfId?: string }>
  deleteConditionalFormattingRule?: (cfId: string) => unknown
  insertRows?: (rowIndex: number, numRows?: number) => unknown
  setColumnWidth?: (columnPosition: number, width: number) => unknown
  setFrozenColumns?: (columns: number) => unknown
  setFrozenRows?: (rows: number) => unknown
  setRowHeight?: (rowPosition: number, height: number) => unknown
  autoFitRow?: (rowPosition: number) => unknown
  setName?: (name: string) => unknown
}

export type UniverFacadeWorkbook = {
  deleteSheet?: (sheet: UniverFacadeWorksheet | string) => boolean
  duplicateActiveSheet?: () => UniverFacadeWorksheet
  getActiveSheet?: () => UniverFacadeWorksheet | null
  getSheetBySheetId?: (sheetId: string) => UniverFacadeWorksheet | null
  insertSheet?: (sheetName?: string) => UniverFacadeWorksheet
  isCellEditing?: () => boolean
  abortEditingAsync?: () => Promise<boolean>
  /** Commit (save=true) or discard the open in-cell edit. */
  endEditingAsync?: (save?: boolean) => Promise<boolean>
  save: () => PureSheetsUniverSnapshot
  setActiveSheet?: (
    sheet: UniverFacadeWorksheet | string,
  ) => UniverFacadeWorksheet
}

export type UniverRangeLike = {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

// Structural (duck-typed) view of Univer's ConditionalFormatHighlightRuleBuilder.
// Every method is optional and chainable; the surface guards each call.
export type UniverConditionalFormatBuilder = {
  whenNumberGreaterThan?: (value: number) => UniverConditionalFormatBuilder
  whenNumberLessThan?: (value: number) => UniverConditionalFormatBuilder
  whenNumberEqualTo?: (value: number) => UniverConditionalFormatBuilder
  whenNumberBetween?: (
    start: number,
    end: number,
  ) => UniverConditionalFormatBuilder
  whenTextEqualTo?: (text: string) => UniverConditionalFormatBuilder
  whenTextContains?: (text: string) => UniverConditionalFormatBuilder
  setBackground?: (color?: string) => UniverConditionalFormatBuilder
  setFontColor?: (color?: string) => UniverConditionalFormatBuilder
  setBold?: (isBold: boolean) => UniverConditionalFormatBuilder
  setItalic?: (isItalic: boolean) => UniverConditionalFormatBuilder
  setUnderline?: (isUnderline: boolean) => UniverConditionalFormatBuilder
  setRanges?: (ranges: UniverRangeLike[]) => UniverConditionalFormatBuilder
  build?: () => unknown
}

// `null` means "clear the target cell" (setValues treats it as an empty
// cell), which sort relies on to move blanks over previously filled cells.
export type UniverPasteValue = string | number | boolean | null | { f: string }
