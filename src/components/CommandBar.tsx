import { useEffect, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownAZ,
  ArrowDownToLine,
  ArrowRightToLine,
  ArrowUpAZ,
  BarChart3,
  Bold,
  Calculator,
  CaseSensitive,
  Clipboard,
  ClipboardType,
  Columns3,
  Copy,
  DecimalsArrowLeft,
  DecimalsArrowRight,
  Eraser,
  FileSpreadsheet,
  Filter,
  Grid2X2,
  History,
  Italic,
  Lock,
  LockOpen,
  Maximize2,
  MessageSquarePlus,
  Minimize2,
  PaintBucket,
  PaintRoller,
  PanelLeft,
  PanelTop,
  Redo2,
  Scissors,
  Search,
  Table,
  TableColumnsSplit,
  TableRowsSplit,
  Trash2,
  Type,
  Underline,
  Undo2,
} from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { ToolIconButton, ToolTextButton, ToolbarDivider } from './controls'
import { ToolbarMenu } from './ToolbarMenu'
import {
  ConditionalFormatControl,
  ValidationControl,
  type ValidationKind,
} from './DataRuleControls'
import type { TableStylePreset } from '../lib/workbookModel'
import type {
  CellStyle,
  ConditionalFormatRule,
  WorkbookChart,
  WorkbookChartType,
} from '../types'

export type FormulaTemplate =
  | 'SUM'
  | 'AVERAGE'
  | 'MIN'
  | 'MAX'
  | 'COUNT'
  | 'COUNTA'
  | 'IF'
  | 'ROUND'
  | 'CONCAT'
  | 'LEFT'
  | 'RIGHT'
  | 'VLOOKUP'
  | 'COUNTIF'
  | 'SUMIF'
  | 'AVERAGEIF'
  | 'RANK'
  | 'IFERROR'
  | 'TODAY'
  | 'NOW'
  | 'ADD'
  | 'SUBTRACT'
  | 'MULTIPLY'
  | 'DIVIDE'

const FORMULA_TEMPLATE_OPTIONS: Array<{
  value: FormulaTemplate
  label: string
}> = [
  { value: 'SUM', label: 'SUM' },
  { value: 'AVERAGE', label: 'AVERAGE' },
  { value: 'MIN', label: 'MIN' },
  { value: 'MAX', label: 'MAX' },
  { value: 'COUNT', label: 'COUNT' },
  { value: 'COUNTA', label: 'COUNTA' },
  { value: 'IF', label: 'IF' },
  { value: 'ROUND', label: 'ROUND' },
  { value: 'CONCAT', label: 'CONCAT' },
  { value: 'LEFT', label: 'LEFT' },
  { value: 'RIGHT', label: 'RIGHT' },
  { value: 'VLOOKUP', label: 'VLOOKUP' },
  { value: 'COUNTIF', label: 'COUNTIF' },
  { value: 'SUMIF', label: 'SUMIF' },
  { value: 'AVERAGEIF', label: 'AVERAGEIF' },
  { value: 'RANK', label: 'RANK' },
  { value: 'IFERROR', label: 'IFERROR' },
  { value: 'TODAY', label: 'TODAY' },
  { value: 'NOW', label: 'NOW' },
  { value: 'ADD', label: 'Add (+)' },
  { value: 'SUBTRACT', label: 'Subtract (-)' },
  { value: 'MULTIPLY', label: 'Multiply (*)' },
  { value: 'DIVIDE', label: 'Divide (/)' },
]

const FONT_SIZE_OPTIONS = ['10', '12', '14', '16', '18', '24'].map(size => ({
  value: size,
  label: size,
}))

const NUMBER_FORMAT_OPTIONS: Array<{
  value: NonNullable<CellStyle['numberFormat']>
  label: string
}> = [
  { value: 'text', label: 'Text' },
  { value: 'number', label: 'Number' },
  { value: 'currency', label: 'Currency' },
  { value: 'percent', label: 'Percent' },
  { value: 'date', label: 'Date' },
]

const TABLE_STYLE_OPTIONS: Array<{ value: TableStylePreset; label: string }> = [
  { value: 'header-row', label: 'Header row' },
  { value: 'banded-rows', label: 'Banded rows' },
  { value: 'totals-row', label: 'Totals row' },
  { value: 'grid-borders', label: 'Full grid borders' },
]

const CHART_TYPE_MENU_OPTIONS: Array<{
  value: WorkbookChartType
  label: string
}> = [
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
  { value: 'pie', label: 'Pie' },
  { value: 'donut', label: 'Donut' },
  { value: 'scatter', label: 'Scatter' },
]

export type ImportExportAction =
  | 'import-csv'
  | 'export-csv'
  | 'import-xlsx'
  | 'export-xlsx'

const IMPORT_EXPORT_OPTIONS: Array<{
  value: ImportExportAction
  label: string
}> = [
  { value: 'import-csv', label: 'Import CSV as new sheet…' },
  { value: 'export-csv', label: 'Export sheet as CSV…' },
  { value: 'import-xlsx', label: 'Import XLSX workbook…' },
  { value: 'export-xlsx', label: 'Export workbook as XLSX…' },
]

/**
 * Command bar: history/clipboard, text formatting, alignment, data tools,
 * structure, formula/number-format menus, and find/replace. Icon buttons with
 * tooltips throughout; dropdowns are custom menus (no native `<select>`).
 */
export function CommandBar({
  selectedStyle,
  canUndo,
  canRedo,
  findText,
  replaceText,
  onUndo,
  onRedo,
  onCopySelectedRange,
  onPasteFromClipboard,
  onToggleStyle,
  onApplyCellStyle,
  onClearSelectedRange,
  onSortSelectedRange,
  onToggleFilter,
  onToggleFirstRowFreeze,
  onToggleFirstColumnFreeze,
  firstRowFrozen,
  firstColumnFrozen,
  onCreateChart,
  onInsertRow,
  onDeleteRow,
  onInsertColumn,
  onDeleteColumn,
  onResizeColumn,
  onResizeRow,
  onPasteValuesOnly,
  onPasteFormattingOnly,
  onFillDown,
  onFillRight,
  onAdjustDecimals,
  onApplyTableStyle,
  onInsertFormulaTemplate,
  onFindTextChange,
  onReplaceTextChange,
  onFindNext,
  onReplaceMatches,
  matchCase,
  onToggleMatchCase,
  onReplaceNext,
  onReplaceAllInSheet,
  onLockSelectedRange,
  onUnlockSelectedRange,
  onToggleCommentBar,
  commentCount = 0,
  filterActive,
  filterQuery,
  filterColumnLabel = '',
  filterHasHeader = true,
  onToggleFilterHeader,
  onCommitFilterQuery,
  selectedRangeLabel,
  onApplyValidation,
  onClearValidation,
  onApplyConditionalFormat,
  onClearConditionalFormats,
  onImportExport,
  onToggleAgentActivity,
  agentLogCount = 0,
}: {
  selectedStyle: CellStyle | undefined
  canUndo: boolean
  canRedo: boolean
  findText: string
  replaceText: string
  onUndo: () => void
  onRedo: () => void
  onCopySelectedRange: (cut?: boolean) => Promise<void>
  onPasteFromClipboard: () => Promise<void>
  onToggleStyle: (style: 'bold' | 'italic' | 'underline') => void
  onApplyCellStyle: (patch: Partial<CellStyle>, message?: string) => void
  onClearSelectedRange: () => void
  onSortSelectedRange: (direction: 'asc' | 'desc') => void
  onToggleFilter: () => void
  onToggleFirstRowFreeze: () => void
  onToggleFirstColumnFreeze: () => void
  /** Reflect the frozen state so the buttons read "Unfreeze …" when active. */
  firstRowFrozen: boolean
  firstColumnFrozen: boolean
  onCreateChart: (type: WorkbookChart['type']) => void
  onInsertRow: () => void
  onDeleteRow: () => void
  onInsertColumn: () => void
  onDeleteColumn: () => void
  onResizeColumn: (delta: number) => void
  onResizeRow: (delta: number) => void
  onPasteValuesOnly: () => void
  onPasteFormattingOnly: () => void
  onFillDown: () => void
  onFillRight: () => void
  onAdjustDecimals: (delta: 1 | -1) => void
  onApplyTableStyle: (preset: TableStylePreset) => void
  onInsertFormulaTemplate: (template: FormulaTemplate) => void
  onFindTextChange: (value: string) => void
  onReplaceTextChange: (value: string) => void
  onFindNext: () => void
  onReplaceMatches: () => void
  matchCase: boolean
  onToggleMatchCase: () => void
  onReplaceNext: () => void
  onReplaceAllInSheet: () => void
  onLockSelectedRange: () => void
  onUnlockSelectedRange: () => void
  onToggleCommentBar: () => void
  /** Number of cells on the active sheet that carry a comment — shown as a
   *  badge on the comment button so the sheet's comments are visible at a
   *  glance without opening the bar. */
  commentCount?: number
  filterActive: boolean
  filterQuery: string
  /** Column label (e.g. "A") the filter applies to — shown so the user can
   *  tell which column is being filtered. */
  filterColumnLabel?: string
  /** Whether the first row is treated as a header (kept visible). */
  filterHasHeader?: boolean
  onToggleFilterHeader?: () => void
  onCommitFilterQuery: (query: string) => void
  selectedRangeLabel: string
  onApplyValidation: (
    kind: ValidationKind,
    options: { min?: number; max?: number },
  ) => void
  onClearValidation: () => void
  onApplyConditionalFormat: (
    condition: ConditionalFormatRule['condition'],
    style: CellStyle,
  ) => void
  onClearConditionalFormats: () => void
  onImportExport: (action: ImportExportAction) => void
  /** Opens the agent-activity panel — the read surface for the workbook's
   *  agent log (what the drawer agent changed, when, and why). */
  onToggleAgentActivity: () => void
  /** Number of agent-log entries — badged on the activity button so agent
   *  work on the workbook is visible at a glance. */
  agentLogCount?: number
}): React.ReactElement {
  const numberFormat = selectedStyle?.numberFormat ?? 'text'
  const numberFormatLabel =
    NUMBER_FORMAT_OPTIONS.find(option => option.value === numberFormat)
      ?.label ?? 'Text'
  const [filterDraft, setFilterDraft] = useState(filterQuery)
  useEffect(() => {
    setFilterDraft(filterQuery)
  }, [filterQuery])
  return (
    <Toolbar aria-label="Spreadsheet toolbar">
      <ToolbarRow>
      <ToolbarGroup aria-label="File history">
        <ToolIconButton
          title="Undo"
          onClick={onUndo}
          disabled={!canUndo}
          icon={Undo2}
        />
        <ToolIconButton
          title="Redo"
          onClick={onRedo}
          disabled={!canRedo}
          icon={Redo2}
        />
        <ToolIconButton
          title="Copy selected range"
          onClick={() => void onCopySelectedRange()}
          icon={Copy}
        />
        <ToolIconButton
          title="Cut selected range"
          onClick={() => void onCopySelectedRange(true)}
          icon={Scissors}
        />
        <ToolIconButton
          title="Paste from clipboard"
          onClick={() => void onPasteFromClipboard()}
          icon={Clipboard}
        />
        <ToolIconButton
          title="Paste values only"
          onClick={onPasteValuesOnly}
          icon={ClipboardType}
        />
        <ToolIconButton
          title="Paste formatting only"
          onClick={onPasteFormattingOnly}
          icon={PaintRoller}
        />
        <ToolIconButton
          title="Fill down from first row of selection"
          onClick={onFillDown}
          icon={ArrowDownToLine}
        />
        <ToolIconButton
          title="Fill right from first column of selection"
          onClick={onFillRight}
          icon={ArrowRightToLine}
        />
      </ToolbarGroup>
      <ToolbarDivider />
      <ToolbarGroup aria-label="Text formatting">
        <ToolIconButton
          title="Bold"
          onClick={() => onToggleStyle('bold')}
          icon={Bold}
        />
        <ToolIconButton
          title="Italic"
          onClick={() => onToggleStyle('italic')}
          icon={Italic}
        />
        <ToolIconButton
          title="Underline"
          onClick={() => onToggleStyle('underline')}
          icon={Underline}
        />
        <ToolbarMenu
          label="Font size"
          displayValue={String(selectedStyle?.fontSize ?? 12)}
          options={FONT_SIZE_OPTIONS}
          selectedValue={String(selectedStyle?.fontSize ?? 12)}
          minWidth={72}
          onSelect={value => onApplyCellStyle({ fontSize: Number(value) })}
        />
        <ColorControl title="Text color">
          <PlatformIcon icon={Type} size={15} strokeWidth={1.9} />
          <input
            aria-label="Text color"
            type="color"
            value={selectedStyle?.textColor ?? '#20231f'}
            onChange={event =>
              onApplyCellStyle({ textColor: event.currentTarget.value })
            }
          />
        </ColorControl>
        <ColorControl title="Fill color">
          <PlatformIcon icon={PaintBucket} size={15} strokeWidth={1.9} />
          <input
            aria-label="Fill color"
            type="color"
            value={selectedStyle?.fillColor ?? '#ffffff'}
            onChange={event =>
              onApplyCellStyle({ fillColor: event.currentTarget.value })
            }
          />
        </ColorControl>
        <ToolIconButton
          title="Toggle all borders"
          onClick={() =>
            onApplyCellStyle(
              { border: !selectedStyle?.border },
              'Updated borders',
            )
          }
          icon={Grid2X2}
        />
        <ToolIconButton
          title="Clear selected range"
          onClick={onClearSelectedRange}
          icon={Eraser}
        />
      </ToolbarGroup>
      <ToolbarDivider />
      <ToolbarGroup aria-label="Alignment">
        <ToolIconButton
          title="Align left"
          onClick={() => onApplyCellStyle({ align: 'left' })}
          icon={AlignLeft}
        />
        <ToolIconButton
          title="Align center"
          onClick={() => onApplyCellStyle({ align: 'center' })}
          icon={AlignCenter}
        />
        <ToolIconButton
          title="Align right"
          onClick={() => onApplyCellStyle({ align: 'right' })}
          icon={AlignRight}
        />
      </ToolbarGroup>
      <ToolbarDivider />
      <ToolbarGroup aria-label="Data tools">
        <ToolIconButton
          title="Sort selected range ascending"
          onClick={() => onSortSelectedRange('asc')}
          icon={ArrowUpAZ}
        />
        <ToolIconButton
          title="Sort selected range descending"
          onClick={() => onSortSelectedRange('desc')}
          icon={ArrowDownAZ}
        />
        <ToolIconButton
          title="Filter selected column"
          onClick={onToggleFilter}
          icon={Filter}
        />
        {filterActive && (
          <FilterControls>
            <FilterColumnChip title={`Filtering column ${filterColumnLabel}`}>
              Col {filterColumnLabel}
            </FilterColumnChip>
            <FindInput
              value={filterDraft}
              onChange={event => setFilterDraft(event.currentTarget.value)}
              onBlur={() => onCommitFilterQuery(filterDraft)}
              onKeyDown={event => {
                if (event.key === 'Enter') onCommitFilterQuery(filterDraft)
              }}
              placeholder="contains"
              aria-label="Filter query"
            />
            <HeaderCheck title="First row is a header (kept visible while filtering)">
              <input
                type="checkbox"
                checked={filterHasHeader}
                onChange={() => onToggleFilterHeader?.()}
              />
              Header row
            </HeaderCheck>
          </FilterControls>
        )}
        <ToolIconButton
          title={firstRowFrozen ? 'Unfreeze first row' : 'Freeze first row'}
          onClick={onToggleFirstRowFreeze}
          icon={PanelTop}
        />
        <ToolIconButton
          title={
            firstColumnFrozen ? 'Unfreeze first column' : 'Freeze first column'
          }
          onClick={onToggleFirstColumnFreeze}
          icon={PanelLeft}
        />
        <ToolIconButton
          title="Protect selected range"
          onClick={onLockSelectedRange}
          icon={Lock}
        />
        <ToolIconButton
          title="Unprotect selected range"
          onClick={onUnlockSelectedRange}
          icon={LockOpen}
        />
        <CommentButtonWrap>
          <ToolIconButton
            title={
              commentCount > 0
                ? `Comments (${commentCount}) — click to view`
                : 'Cell comment'
            }
            onClick={onToggleCommentBar}
            icon={MessageSquarePlus}
          />
          {commentCount > 0 ? (
            <CommentBadge aria-hidden="true">{commentCount}</CommentBadge>
          ) : null}
        </CommentButtonWrap>
        <ValidationControl
          selectedRangeLabel={selectedRangeLabel}
          onApply={onApplyValidation}
          onClear={onClearValidation}
        />
        <ConditionalFormatControl
          selectedRangeLabel={selectedRangeLabel}
          onApply={onApplyConditionalFormat}
          onClear={onClearConditionalFormats}
        />
        <ToolbarMenu
          label="Insert chart from selected range"
          displayValue="Chart"
          icon={BarChart3}
          options={CHART_TYPE_MENU_OPTIONS}
          minWidth={110}
          compact
          onSelect={value => onCreateChart(value as WorkbookChartType)}
        />
        <ToolbarMenu
          label="Import or export data"
          displayValue="Data"
          icon={FileSpreadsheet}
          options={IMPORT_EXPORT_OPTIONS}
          minWidth={200}
          compact
          onSelect={value => onImportExport(value as ImportExportAction)}
        />
        <CommentButtonWrap>
          <ToolIconButton
            title={
              agentLogCount > 0
                ? `Agent activity (${agentLogCount}) — click to view`
                : 'Agent activity'
            }
            onClick={onToggleAgentActivity}
            icon={History}
          />
          {agentLogCount > 0 ? (
            <CommentBadge aria-hidden="true">{agentLogCount}</CommentBadge>
          ) : null}
        </CommentButtonWrap>
      </ToolbarGroup>
      </ToolbarRow>
      <ToolbarRow>
            <ToolbarGroup aria-label="Structure">
        <ToolIconButton
          title="Insert row"
          onClick={onInsertRow}
          icon={TableRowsSplit}
        />
        <ToolIconButton
          title="Delete row"
          onClick={onDeleteRow}
          icon={Trash2}
        />
        <ToolbarDivider />
        <ToolIconButton
          title="Insert column"
          onClick={onInsertColumn}
          icon={TableColumnsSplit}
        />
        <ToolIconButton
          title="Delete column"
          onClick={onDeleteColumn}
          icon={Trash2}
        />
        <ToolbarDivider />
        <ToolIconButton
          title="Widen selected column"
          onClick={() => onResizeColumn(24)}
          icon={Columns3}
        />
        <ToolIconButton
          title="Narrow selected column"
          onClick={() => onResizeColumn(-24)}
          icon={Minimize2}
        />
        <ToolIconButton
          title="Taller selected row"
          onClick={() => onResizeRow(8)}
          icon={Maximize2}
        />
        <ToolIconButton
          title="Shorter selected row"
          onClick={() => onResizeRow(-8)}
          icon={Minimize2}
        />
      </ToolbarGroup>
      <ToolbarDivider />
      <ToolbarGroup aria-label="Number and find">
        <ToolbarMenu
          label="Insert formula or operation"
          displayValue="Formula"
          icon={Calculator}
          options={FORMULA_TEMPLATE_OPTIONS}
          minWidth={140}
          compact
          onSelect={value => onInsertFormulaTemplate(value as FormulaTemplate)}
        />
        <ToolbarMenu
          label="Number format"
          displayValue={numberFormatLabel}
          options={NUMBER_FORMAT_OPTIONS}
          selectedValue={numberFormat}
          minWidth={110}
          onSelect={value =>
            onApplyCellStyle({
              numberFormat: value as CellStyle['numberFormat'],
            })
          }
        />
        <ToolIconButton
          title="Increase decimal places"
          onClick={() => onAdjustDecimals(1)}
          icon={DecimalsArrowRight}
        />
        <ToolIconButton
          title="Decrease decimal places"
          onClick={() => onAdjustDecimals(-1)}
          icon={DecimalsArrowLeft}
        />
        <ToolbarMenu
          label="Table style"
          displayValue="Table"
          icon={Table}
          options={TABLE_STYLE_OPTIONS}
          minWidth={140}
          compact
          onSelect={value => onApplyTableStyle(value as TableStylePreset)}
        />
        <FindBox>
          <FindInput
            value={findText}
            onChange={event => onFindTextChange(event.currentTarget.value)}
            onKeyDown={event => {
              if (event.key === 'Enter') onFindNext()
            }}
            placeholder="Find"
            aria-label="Find cells"
          />
          <FindInput
            value={replaceText}
            onChange={event => onReplaceTextChange(event.currentTarget.value)}
            onKeyDown={event => {
              if (event.key === 'Enter') onReplaceNext()
            }}
            placeholder="Replace"
            aria-label="Replace with"
          />
          <ToolIconButton
            title={matchCase ? 'Match case: on' : 'Match case: off'}
            aria-pressed={matchCase}
            onClick={onToggleMatchCase}
            icon={CaseSensitive}
          />
          <ToolIconButton
            title="Find next"
            onClick={onFindNext}
            icon={Search}
          />
          {/* Labeled text buttons: the replace actions were undiscoverable as
              look-alike icons (two identical "Replace" icons) (#277). */}
          <FindActionButton
            type="button"
            title="Replace next match"
            aria-label="Replace next match"
            onClick={onReplaceNext}
          >
            Replace
          </FindActionButton>
          <FindActionButton
            type="button"
            title="Replace matches in selected range"
            aria-label="Replace matches in selected range"
            onClick={onReplaceMatches}
          >
            In range
          </FindActionButton>
          <FindActionButton
            type="button"
            title="Replace all in sheet"
            aria-label="Replace all in sheet"
            onClick={onReplaceAllInSheet}
          >
            All
          </FindActionButton>
        </FindBox>
      </ToolbarGroup>
      </ToolbarRow>
    </Toolbar>
  )
}

/** Typed loosely on purpose: styled-components' attrs rejects data-* literals. */
const selectChrome: Record<string, string> = { 'data-chrome': 'toolbar-select' }

const CommentButtonWrap = styled.div`
  position: relative;
  display: inline-flex;
`

const CommentBadge = styled.span`
  position: absolute;
  top: -3px;
  right: -3px;
  display: grid;
  place-items: center;
  min-width: 14px;
  height: 14px;
  padding: 0 3px;
  border-radius: 999px;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-family: var(--platform-typography-font-family-mono);
  font-size: 9px;
  font-weight: 700;
  line-height: 1;
  pointer-events: none;
`

const Toolbar = styled.div`
  display: flex;
  flex-direction: column;
  gap: 4px;
  align-items: flex-end;

  @media (max-width: 720px) {
    grid-column: 1 / -1;
    align-items: flex-start;
  }
`

const ToolbarRow = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  align-items: center;
  justify-content: flex-end;

  @media (max-width: 720px) {
    justify-content: flex-start;
  }
`

const ToolbarGroup = styled.div`
  display: inline-flex;
  flex-wrap: wrap;
  gap: 3px;
  align-items: center;
  min-width: 0;
`

const ColorControl = styled.label`
  position: relative;
  display: inline-grid;
  place-items: center;
  min-width: var(--pure-chrome-control-height);
  height: var(--pure-chrome-control-height);
  border: 0;
  border-radius: 7px;
  background: transparent;
  color: var(--sheets-ink);
  font-size: 12px;
  font-weight: 700;

  &:hover {
    background: var(--pure-chrome-hover);
  }

  input {
    position: absolute;
    inset: 0;
    width: 100%;
    height: 100%;
    opacity: 0;
  }

  &[title]:hover::after,
  &:focus-within::after {
    position: absolute;
    top: calc(100% + 6px);
    left: 50%;
    z-index: 50;
    max-width: 220px;
    padding: 5px 7px;
    border: 1px solid var(--platform-colors-border);
    border-radius: 0;
    background: var(--platform-colors-popover);
    box-shadow: var(--platform-shadow-md);
    color: var(--platform-colors-popover-text);
    content: attr(title);
    font-size: var(--pure-chrome-meta-size);
    font-weight: 600;
    line-height: 1.2;
    pointer-events: none;
    text-align: center;
    transform: translateX(-50%);
    white-space: nowrap;
  }
`

const FindBox = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  align-items: center;
  min-width: 0;
`

// Labelled toolbar controls for the replace actions — the icon-only versions
// were undiscoverable (#277).
const FindActionButton = styled(ToolTextButton)``

// A toolbar-height outlined input (the 28px toolbar select's box).
const FindInput = styled.input.attrs(selectChrome)`
  width: 88px;
  padding: 0 8px;
  font-weight: 400;

  @media (max-width: 720px) {
    width: 96px;
  }
`

const FilterControls = styled.div`
  display: inline-flex;
  align-items: center;
  gap: 6px;
`

const FilterColumnChip = styled.span`
  display: inline-grid;
  place-items: center;
  min-width: 34px;
  height: 24px;
  padding: 0 7px;
  border-radius: 0;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-family: var(--platform-typography-font-family-mono);
  font-size: var(--pure-chrome-meta-size);
  font-weight: 700;
  white-space: nowrap;
`

const HeaderCheck = styled.label`
  display: inline-flex;
  align-items: center;
  gap: 5px;
  color: var(--pure-chrome-muted);
  font-size: 12px;
  font-weight: 600;
  white-space: nowrap;
  cursor: pointer;

  input {
    margin: 0;
    cursor: pointer;
  }
`
