import { useEffect, useId, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  ArrowDownAZ,
  ArrowDownToLine,
  ArrowRightToLine,
  ArrowUpAZ,
  BarChart3,
  Bold,
  Calculator,
  CaseSensitive,
  ChevronUp,
  Clipboard,
  ClipboardType,
  Code2,
  Columns3,
  Combine,
  Command,
  Copy,
  DecimalsArrowLeft,
  DecimalsArrowRight,
  Eraser,
  Eye,
  EyeOff,
  FileSpreadsheet,
  Filter,
  Grid3X3,
  History,
  Italic,
  Lock,
  LockOpen,
  Maximize2,
  MessageSquarePlus,
  Minimize2,
  MoveHorizontal,
  MoveVertical,
  PaintRoller,
  PanelLeft,
  PanelRight,
  PanelTop,
  Printer,
  Redo2,
  RemoveFormatting,
  Scissors,
  Search,
  Strikethrough,
  Table,
  Table2,
  TableColumnsSplit,
  TableRowsSplit,
  Trash2,
  Underline,
  Undo2,
  WrapText,
  X,
} from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { ToolIconButton, ToolTextButton, ToolbarDivider } from './controls'
import { ToolbarMenu } from './ToolbarMenu'
import {
  ConditionalFormatControl,
  ValidationControl,
  type ValidationKind,
} from './DataRuleControls'
import {
  BorderPickerControl,
  ColorSwatchControl,
  CustomNumberFormatPopover,
  FreezePanesControl,
  type BorderPreset,
} from './StylePopovers'
import { SmartFilterControl } from './SmartFilterPopover'
import {
  DataCleanupControl,
  MultiSortControl,
} from './SortAndTransformPopovers'
import type { SortKey } from '../lib/dataTransforms'
import type { FormulaWorkbook, TableStylePreset } from '../lib/workbookModel'
import type {
  CellStyle,
  ColumnFilterRule,
  ConditionalFormatRule,
  WorkbookChart,
  WorkbookChartType,
  WorkbookSheet,
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
  | 'XLOOKUP'
  | 'COUNTIF'
  | 'SUMIF'
  | 'SUMIFS'
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
  { value: 'XLOOKUP', label: 'XLOOKUP' },
  { value: 'VLOOKUP', label: 'VLOOKUP' },
  { value: 'COUNTIF', label: 'COUNTIF' },
  { value: 'SUMIF', label: 'SUMIF' },
  { value: 'SUMIFS', label: 'SUMIFS' },
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

const FONT_FAMILY_OPTIONS = [
  { value: 'Arial', label: 'Arial (Default)' },
  { value: 'Inter', label: 'Inter' },
  { value: 'Georgia', label: 'Georgia' },
  { value: 'Courier New', label: 'Courier New' },
  { value: 'Helvetica Neue', label: 'Helvetica Neue' },
]

const FONT_SIZE_OPTIONS = ['9', '10', '11', '12', '14', '16', '18', '24'].map(
  size => ({ value: size, label: size }),
)

const VERTICAL_ALIGN_OPTIONS = [
  { value: 'top', label: 'Align top' },
  { value: 'middle', label: 'Align middle' },
  { value: 'bottom', label: 'Align bottom' },
]

const NUMBER_FORMAT_OPTIONS: Array<{
  value: NonNullable<CellStyle['numberFormat']> | 'custom'
  label: string
}> = [
  { value: 'text', label: 'Text' },
  { value: 'number', label: 'Number' },
  { value: 'currency', label: 'Currency ($)' },
  { value: 'accounting', label: 'Accounting' },
  { value: 'percent', label: 'Percent (%)' },
  { value: 'date', label: 'Date' },
  { value: 'time', label: 'Time' },
  { value: 'scientific', label: 'Scientific' },
  { value: 'custom', label: 'Currencies & custom…' },
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
  { value: 'stacked-bar', label: 'Stacked Bar' },
  { value: 'horizontal-bar', label: 'Horizontal Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
  { value: 'combo', label: 'Combo (Bar + Line)' },
  { value: 'pie', label: 'Pie' },
  { value: 'donut', label: 'Donut' },
  { value: 'scatter', label: 'Scatter' },
]

export type ImportExportAction =
  | 'import-csv'
  | 'export-csv'
  | 'export-selection-csv'
  | 'export-selection-tsv'
  | 'export-selection-json'
  | 'import-xlsx'
  | 'export-xlsx'
  | 'print-pdf'

const IMPORT_EXPORT_OPTIONS: Array<{
  value: ImportExportAction
  label: string
}> = [
  { value: 'import-csv', label: 'Import CSV as new sheet…' },
  { value: 'export-csv', label: 'Export sheet as CSV…' },
  { value: 'export-selection-csv', label: 'Export selection as CSV…' },
  { value: 'export-selection-tsv', label: 'Export selection as TSV…' },
  { value: 'export-selection-json', label: 'Export selection as JSON…' },
  { value: 'import-xlsx', label: 'Open Excel workbook (.xlsx)…' },
  { value: 'export-xlsx', label: 'Save as Excel (.xlsx)…' },
  { value: 'print-pdf', label: 'Print / Export PDF…' },
]

export function CommandBar({
  activeSheet,
  workbook,
  selectedCell = 'A1',
  selectedRow = 0,
  selectedColumn = 0,
  selectedStyle,
  isMerged = false,
  formatPainterActive = false,
  showGridlines = true,
  showFormulas = false,
  inspectorOpen = false,
  activeFilterCount = 0,
  canUndo,
  canRedo,
  findOpen: controlledFindOpen,
  onFindOpenChange,
  findFocusRequest = 0,
  findText,
  replaceText,
  findMatchCount = 0,
  findActiveIndex = -1,
  onUndo,
  onRedo,
  onCopySelectedRange,
  onPasteFromClipboard,
  onToggleStyle,
  onApplyCellStyle,
  onApplyBorderPreset,
  onClearSelectedRange,
  onClearFormatting,
  onToggleFormatPainter,
  onToggleMerge,
  onSortSelectedRange,
  sortHasHeader = false,
  onToggleSortHeader,
  onMultiSort,
  onToggleFilter,
  onApplyColumnFilter,
  onClearAllFilters,
  onToggleFirstRowFreeze,
  onToggleFirstColumnFreeze,
  onSetFreezePanes,
  firstRowFrozen,
  firstColumnFrozen,
  onCreateChart,
  onCreatePivot,
  onInsertRow,
  onDeleteRow,
  onInsertColumn,
  onDeleteColumn,
  onAutoFitColumns,
  onAutoFitRows,
  onHideRows,
  onUnhideRows,
  onHideColumns,
  onUnhideColumns,
  onToggleGridlines,
  onToggleShowFormulas,
  onPrintSheet,
  onResizeColumn,
  onResizeRow,
  onPasteValuesOnly,
  onPasteFormattingOnly,
  onPasteTranspose,
  onFillDown,
  onFillRight,
  onAdjustDecimals,
  onApplyTableStyle,
  onInsertFormulaTemplate,
  onRemoveDuplicates,
  onSplitTextToColumns,
  onTrimWhitespace,
  onToggleCheckboxes,
  onSetHyperlink,
  onFindTextChange,
  onReplaceTextChange,
  onFindNext,
  onFindPrevious,
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
  onToggleInspector,
  onOpenCommandPalette,
  agentLogCount = 0,
}: {
  activeSheet?: WorkbookSheet
  workbook?: FormulaWorkbook
  selectedCell?: string
  selectedRow?: number
  selectedColumn?: number
  selectedStyle: CellStyle | undefined
  isMerged?: boolean
  formatPainterActive?: boolean
  showGridlines?: boolean
  showFormulas?: boolean
  inspectorOpen?: boolean
  activeFilterCount?: number
  canUndo: boolean
  canRedo: boolean
  findOpen?: boolean
  onFindOpenChange?: (open: boolean) => void
  findFocusRequest?: number
  findText: string
  replaceText: string
  findMatchCount?: number
  findActiveIndex?: number
  onUndo: () => void
  onRedo: () => void
  onCopySelectedRange: (cut?: boolean) => Promise<void>
  onPasteFromClipboard: () => Promise<void>
  onToggleStyle: (style: 'bold' | 'italic' | 'underline' | 'strikethrough') => void
  onApplyCellStyle: (patch: Partial<CellStyle>, message?: string) => void
  onApplyBorderPreset?: (preset: BorderPreset, color?: string) => void
  onClearSelectedRange: () => void
  onClearFormatting?: () => void
  onToggleFormatPainter?: () => void
  onToggleMerge?: () => void
  onSortSelectedRange: (direction: 'asc' | 'desc') => void
  sortHasHeader?: boolean
  onToggleSortHeader?: () => void
  onMultiSort?: (keys: SortKey[], hasHeader: boolean) => void
  onToggleFilter: () => void
  onApplyColumnFilter?: (column: number, rule: ColumnFilterRule | null) => void
  onClearAllFilters?: () => void
  onToggleFirstRowFreeze: () => void
  onToggleFirstColumnFreeze: () => void
  onSetFreezePanes?: (rows: number, columns: number) => void
  firstRowFrozen: boolean
  firstColumnFrozen: boolean
  onCreateChart: (type: WorkbookChart['type']) => void
  onCreatePivot?: () => void
  onInsertRow: () => void
  onDeleteRow: () => void
  onInsertColumn: () => void
  onDeleteColumn: () => void
  onAutoFitColumns?: () => void
  onAutoFitRows?: () => void
  onHideRows?: () => void
  onUnhideRows?: () => void
  onHideColumns?: () => void
  onUnhideColumns?: () => void
  onToggleGridlines?: () => void
  onToggleShowFormulas?: () => void
  onPrintSheet?: () => void
  onResizeColumn: (delta: number) => void
  onResizeRow: (delta: number) => void
  onPasteValuesOnly: () => void
  onPasteFormattingOnly: () => void
  onPasteTranspose?: () => void
  onFillDown: () => void
  onFillRight: () => void
  onAdjustDecimals: (delta: 1 | -1) => void
  onApplyTableStyle: (preset: TableStylePreset) => void
  onInsertFormulaTemplate: (template: FormulaTemplate) => void
  onRemoveDuplicates?: (hasHeader: boolean) => void
  onSplitTextToColumns?: (delimiter: string) => void
  onTrimWhitespace?: () => void
  onToggleCheckboxes?: () => void
  onSetHyperlink?: (url: string, label?: string) => void
  onFindTextChange: (value: string) => void
  onReplaceTextChange: (value: string) => void
  onFindNext: () => void
  onFindPrevious?: () => void
  onReplaceMatches: () => void
  matchCase: boolean
  onToggleMatchCase: () => void
  onReplaceNext: () => void
  onReplaceAllInSheet: () => void
  onLockSelectedRange: () => void
  onUnlockSelectedRange: () => void
  onToggleCommentBar: () => void
  commentCount?: number
  filterActive: boolean
  filterQuery: string
  filterColumnLabel?: string
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
  onToggleAgentActivity: () => void
  onToggleInspector?: () => void
  onOpenCommandPalette?: () => void
  agentLogCount?: number
}): React.ReactElement {
  const numberFormat = selectedStyle?.numberFormat ?? 'text'
  const numberFormatLabel =
    NUMBER_FORMAT_OPTIONS.find(option => option.value === numberFormat)
      ?.label ?? 'Text'
  const toolbarId = useId()
  const [section, setSection] = useState<'Format' | 'Data' | 'Edit'>('Format')
  const [internalFindOpen, setInternalFindOpen] = useState(false)
  const [customNumOpen, setCustomNumOpen] = useState(false)
  const findOpen = controlledFindOpen ?? internalFindOpen
  const setFindOpen = (next: boolean | ((open: boolean) => boolean)): void => {
    const resolved = typeof next === 'function' ? next(findOpen) : next
    setInternalFindOpen(resolved)
    onFindOpenChange?.(resolved)
  }
  const findInputRef = useRef<HTMLInputElement>(null)
  const findTriggerRef = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    if (findOpen) findInputRef.current?.focus()
  }, [findOpen, findFocusRequest])
  const [filterDraft, setFilterDraft] = useState(filterQuery)
  useEffect(() => {
    setFilterDraft(filterQuery)
  }, [filterQuery])

  const anyFilterActive = filterActive || activeFilterCount > 0

  return (
    <Toolbar aria-label="Spreadsheet toolbar">
      <ToolbarHeader>
        <Sections role="tablist" aria-label="Spreadsheet tools">
          {(['Format', 'Data', 'Edit'] as const).map((name, index, names) => (
            <SectionButton
              key={name}
              id={`${toolbarId}-${name}-tab`}
              type="button"
              role="tab"
              aria-selected={section === name}
              aria-controls={`${toolbarId}-${name}-panel`}
              tabIndex={section === name ? 0 : -1}
              onClick={() => setSection(name)}
              onKeyDown={event => {
                const next =
                  event.key === 'ArrowRight'
                    ? (index + 1) % names.length
                    : event.key === 'ArrowLeft'
                    ? (index + names.length - 1) % names.length
                    : event.key === 'Home'
                    ? 0
                    : event.key === 'End'
                    ? names.length - 1
                    : -1
                if (next < 0) return
                event.preventDefault()
                setSection(names[next])
                const tabs =
                  event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>(
                    '[role="tab"]',
                  )
                tabs?.[next]?.focus()
              }}
            >
              {name}
              {name === 'Data' && anyFilterActive ? (
                <ActiveDot aria-label="Filter active" />
              ) : null}
            </SectionButton>
          ))}
        </Sections>
        <QuickActions>
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
          <ToolbarMenu
            label="Import or export data"
            displayValue="File"
            icon={FileSpreadsheet}
            options={IMPORT_EXPORT_OPTIONS}
            minWidth={210}
            onSelect={value => onImportExport(value as ImportExportAction)}
          />
          <ToolTextButton
            ref={findTriggerRef}
            type="button"
            aria-label="Find and replace"
            aria-expanded={findOpen}
            aria-controls={`${toolbarId}-find`}
            onClick={() => setFindOpen(open => !open)}
          >
            <PlatformIcon icon={Search} size={15} /> <span>Find</span>
          </ToolTextButton>
          {onOpenCommandPalette && (
            <ToolTextButton
              type="button"
              title="Command palette (⌘K)"
              aria-label="Command palette"
              onClick={onOpenCommandPalette}
            >
              <PlatformIcon icon={Command} size={14} /> <span>⌘K</span>
            </ToolTextButton>
          )}
          {onToggleInspector && (
            <ToolIconButton
              title={inspectorOpen ? 'Hide inspector' : 'Open inspector (Rules, Pivot, Names, Audit)'}
              aria-pressed={inspectorOpen}
              onClick={onToggleInspector}
              icon={PanelRight}
            />
          )}
        </QuickActions>
      </ToolbarHeader>

      <ToolbarRow
        role="tabpanel"
        id={`${toolbarId}-Format-panel`}
        aria-labelledby={`${toolbarId}-Format-tab`}
        hidden={section !== 'Format'}
      >
        <ToolbarGroup role="group" aria-label="Text formatting">
          {onToggleFormatPainter && (
            <ToolIconButton
              title={
                formatPainterActive
                  ? 'Format painter active — click a cell to apply'
                  : 'Format painter'
              }
              aria-pressed={formatPainterActive}
              onClick={onToggleFormatPainter}
              icon={PaintRoller}
            />
          )}
          <ToolIconButton
            title="Bold"
            aria-pressed={Boolean(selectedStyle?.bold)}
            onClick={() => onToggleStyle('bold')}
            icon={Bold}
          />
          <ToolIconButton
            title="Italic"
            aria-pressed={Boolean(selectedStyle?.italic)}
            onClick={() => onToggleStyle('italic')}
            icon={Italic}
          />
          <ToolIconButton
            title="Underline"
            aria-pressed={Boolean(selectedStyle?.underline)}
            onClick={() => onToggleStyle('underline')}
            icon={Underline}
          />
          <ToolIconButton
            title="Strikethrough"
            aria-pressed={Boolean(selectedStyle?.strikethrough)}
            onClick={() => onToggleStyle('strikethrough')}
            icon={Strikethrough}
          />
          <ToolbarMenu
            label="Font family"
            displayValue={selectedStyle?.fontFamily ?? 'Arial'}
            options={FONT_FAMILY_OPTIONS}
            selectedValue={selectedStyle?.fontFamily ?? 'Arial'}
            minWidth={132}
            onSelect={value =>
              onApplyCellStyle({ fontFamily: value === 'Arial' ? undefined : value })
            }
          />
          <ToolbarMenu
            label="Font size"
            displayValue={String(selectedStyle?.fontSize ?? 12)}
            options={FONT_SIZE_OPTIONS}
            selectedValue={String(selectedStyle?.fontSize ?? 12)}
            minWidth={72}
            onSelect={value => onApplyCellStyle({ fontSize: Number(value) })}
          />
          <ColorSwatchControl
            mode="text"
            value={selectedStyle?.textColor}
            onChange={color =>
              onApplyCellStyle(
                { textColor: color },
                color ? 'Updated text color' : 'Reset text color',
              )
            }
          />
          <ColorSwatchControl
            mode="fill"
            value={selectedStyle?.fillColor}
            onChange={color =>
              onApplyCellStyle(
                { fillColor: color },
                color ? 'Updated fill color' : 'Cleared fill color',
              )
            }
          />
          <BorderPickerControl
            hasBorder={Boolean(
              selectedStyle?.border || selectedStyle?.borderSides?.length,
            )}
            onToggleAllBorders={() =>
              onApplyCellStyle(
                { border: !selectedStyle?.border },
                'Updated borders',
              )
            }
            onApplyBorder={(preset, color) => {
              if (onApplyBorderPreset) onApplyBorderPreset(preset, color)
              else
                onApplyCellStyle(
                  { border: preset !== 'none' },
                  'Updated borders',
                )
            }}
          />
          {onClearFormatting && (
            <ToolIconButton
              title="Clear formatting"
              onClick={onClearFormatting}
              icon={RemoveFormatting}
            />
          )}
          <ToolIconButton
            title="Clear selected range"
            onClick={onClearSelectedRange}
            icon={Eraser}
          />
        </ToolbarGroup>
        <ToolbarDivider />
        <ToolbarGroup role="group" aria-label="Alignment">
          <ToolIconButton
            title="Align left"
            aria-pressed={selectedStyle?.align === 'left'}
            onClick={() => onApplyCellStyle({ align: 'left' })}
            icon={AlignLeft}
          />
          <ToolIconButton
            title="Align center"
            aria-pressed={selectedStyle?.align === 'center'}
            onClick={() => onApplyCellStyle({ align: 'center' })}
            icon={AlignCenter}
          />
          <ToolIconButton
            title="Align right"
            aria-pressed={selectedStyle?.align === 'right'}
            onClick={() => onApplyCellStyle({ align: 'right' })}
            icon={AlignRight}
          />
          <ToolbarMenu
            label="Vertical alignment"
            displayValue=""
            icon={AlignVerticalJustifyCenter}
            options={VERTICAL_ALIGN_OPTIONS}
            selectedValue={selectedStyle?.verticalAlign ?? 'middle'}
            minWidth={116}
            onSelect={value =>
              onApplyCellStyle({
                verticalAlign: value as CellStyle['verticalAlign'],
              })
            }
          />
          <ToolIconButton
            title="Wrap text"
            aria-pressed={selectedStyle?.wrap === 'wrap'}
            onClick={() =>
              onApplyCellStyle(
                { wrap: selectedStyle?.wrap === 'wrap' ? 'overflow' : 'wrap' },
                selectedStyle?.wrap === 'wrap'
                  ? 'Disabled text wrap'
                  : 'Enabled text wrap',
              )
            }
            icon={WrapText}
          />
          {onToggleMerge && (
            <ToolIconButton
              title={isMerged ? 'Unmerge cells' : 'Merge cells'}
              aria-pressed={isMerged}
              onClick={onToggleMerge}
              icon={Combine}
            />
          )}
        </ToolbarGroup>
        <ToolbarDivider />
        <ToolbarGroup role="group" aria-label="Numbers and formulas">
          <ToolbarMenu
            label="Insert formula or operation"
            displayValue="Formula"
            icon={Calculator}
            options={FORMULA_TEMPLATE_OPTIONS}
            minWidth={146}
            onSelect={value =>
              onInsertFormulaTemplate(value as FormulaTemplate)
            }
          />
          <NumFormatWrap>
            <ToolbarMenu
              label="Number format"
              displayValue={numberFormatLabel}
              options={NUMBER_FORMAT_OPTIONS}
              selectedValue={numberFormat}
              minWidth={160}
              onSelect={value => {
                if (value === 'custom') {
                  setCustomNumOpen(true)
                  return
                }
                onApplyCellStyle({
                  numberFormat: value as CellStyle['numberFormat'],
                })
              }}
            />
            <CustomNumberFormatPopover
              open={customNumOpen}
              currentStyle={selectedStyle}
              onApply={(patch, msg) => onApplyCellStyle(patch, msg)}
              onClose={() => setCustomNumOpen(false)}
            />
          </NumFormatWrap>
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
            onSelect={value => onApplyTableStyle(value as TableStylePreset)}
          />
        </ToolbarGroup>
      </ToolbarRow>

      <ToolbarRow
        role="tabpanel"
        id={`${toolbarId}-Data-panel`}
        aria-labelledby={`${toolbarId}-Data-tab`}
        hidden={section !== 'Data'}
      >
        <ToolbarGroup role="group" aria-label="Data tools">
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
          {onToggleSortHeader && (
            <HeaderCheck title="Keep first row pinned as header when sorting">
              <input
                type="checkbox"
                aria-label="Sort with header row"
                checked={sortHasHeader}
                onChange={onToggleSortHeader}
              />
              Sort header
            </HeaderCheck>
          )}
          {activeSheet && onMultiSort && (
            <MultiSortControl
              sheet={activeSheet}
              selectedRange={selectedRangeLabel}
              onApplySort={onMultiSort}
            />
          )}
          <ToolIconButton
            title="Filter selected column"
            aria-pressed={filterActive}
            onClick={onToggleFilter}
            icon={Filter}
          />
          {activeSheet && onApplyColumnFilter && onClearAllFilters && (
            <SmartFilterControl
              sheet={activeSheet}
              workbook={workbook}
              column={selectedColumn}
              activeFilterCount={activeFilterCount}
              onApplyColumnFilter={onApplyColumnFilter}
              onClearAllFilters={onClearAllFilters}
            />
          )}
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
          {onRemoveDuplicates &&
            onSplitTextToColumns &&
            onTrimWhitespace &&
            onToggleCheckboxes &&
            onSetHyperlink && (
              <DataCleanupControl
                selectedRange={selectedRangeLabel}
                onRemoveDuplicates={onRemoveDuplicates}
                onSplitTextToColumns={onSplitTextToColumns}
                onTrimWhitespace={onTrimWhitespace}
                onToggleCheckboxes={onToggleCheckboxes}
                onSetHyperlink={onSetHyperlink}
              />
            )}
          <ToolbarDivider />
          <ToolIconButton
            aria-pressed={firstRowFrozen}
            title={firstRowFrozen ? 'Unfreeze first row' : 'Freeze first row'}
            onClick={onToggleFirstRowFreeze}
            icon={PanelTop}
          />
          <ToolIconButton
            title={
              firstColumnFrozen
                ? 'Unfreeze first column'
                : 'Freeze first column'
            }
            aria-pressed={firstColumnFrozen}
            onClick={onToggleFirstColumnFreeze}
            icon={PanelLeft}
          />
          {onSetFreezePanes && (
            <FreezePanesControl
              frozenRows={activeSheet?.frozenRows ?? 0}
              frozenColumns={activeSheet?.frozenColumns ?? 0}
              selectedCell={selectedCell}
              selectedRow={selectedRow}
              selectedCol={selectedColumn}
              onSetFreeze={onSetFreezePanes}
            />
          )}
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
          {onCreatePivot && (
            <ToolTextButton
              type="button"
              title="Insert Pivot Table sheet from selected range"
              aria-label="Insert Pivot Table"
              onClick={onCreatePivot}
            >
              <PlatformIcon icon={Table2} size={14} /> <span>Pivot</span>
            </ToolTextButton>
          )}
          <ToolbarMenu
            label="Insert chart from selected range"
            displayValue="Chart"
            icon={BarChart3}
            options={CHART_TYPE_MENU_OPTIONS}
            minWidth={148}
            onSelect={value => onCreateChart(value as WorkbookChartType)}
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

      <ToolbarRow
        role="tabpanel"
        id={`${toolbarId}-Edit-panel`}
        aria-labelledby={`${toolbarId}-Edit-tab`}
        hidden={section !== 'Edit'}
      >
        <ToolbarGroup role="group" aria-label="Clipboard">
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
          {onPasteTranspose && (
            <ToolTextButton
              type="button"
              title="Paste transposed (flip rows & columns)"
              aria-label="Paste transposed"
              onClick={onPasteTranspose}
            >
              <span>Transpose</span>
            </ToolTextButton>
          )}
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
        <ToolbarGroup role="group" aria-label="Structure">
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
          {onAutoFitColumns && (
            <ToolIconButton
              title="Auto-fit selected columns"
              onClick={onAutoFitColumns}
              icon={MoveHorizontal}
            />
          )}
          {onAutoFitRows && (
            <ToolIconButton
              title="Auto-fit selected rows"
              onClick={onAutoFitRows}
              icon={MoveVertical}
            />
          )}
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
          {onHideRows && (
            <ToolIconButton
              title="Hide selected rows"
              onClick={onHideRows}
              icon={EyeOff}
            />
          )}
          {onUnhideRows && (activeSheet?.hiddenRows?.length ?? 0) > 0 && (
            <ToolIconButton
              title="Unhide all rows"
              onClick={onUnhideRows}
              icon={Eye}
            />
          )}
          {onHideColumns && (
            <ToolTextButton
              type="button"
              title="Hide selected columns"
              aria-label="Hide columns"
              onClick={onHideColumns}
            >
              <span>Hide col</span>
            </ToolTextButton>
          )}
          {onUnhideColumns && (activeSheet?.hiddenColumns?.length ?? 0) > 0 && (
            <ToolTextButton
              type="button"
              title="Unhide all columns"
              aria-label="Unhide columns"
              onClick={onUnhideColumns}
            >
              <span>Show cols</span>
            </ToolTextButton>
          )}
        </ToolbarGroup>
        <ToolbarDivider />
        <ToolbarGroup role="group" aria-label="View">
          {onToggleGridlines && (
            <ToolIconButton
              title={showGridlines ? 'Hide gridlines' : 'Show gridlines'}
              aria-pressed={showGridlines}
              onClick={onToggleGridlines}
              icon={Grid3X3}
            />
          )}
          {onToggleShowFormulas && (
            <ToolIconButton
              title={showFormulas ? 'Show values' : 'Show formulas (Ctrl+`)'}
              aria-pressed={showFormulas}
              onClick={onToggleShowFormulas}
              icon={Code2}
            />
          )}
          {onPrintSheet && (
            <ToolIconButton
              title="Print / Export PDF"
              onClick={onPrintSheet}
              icon={Printer}
            />
          )}
        </ToolbarGroup>
      </ToolbarRow>

      <FindBox
        id={`${toolbarId}-find`}
        hidden={!findOpen}
        onKeyDown={event => {
          if (event.key === 'Escape') {
            event.preventDefault()
            event.stopPropagation()
            setFindOpen(false)
            findTriggerRef.current?.focus()
          }
        }}
      >
        <FindInput
          ref={findInputRef}
          value={findText}
          onChange={event => onFindTextChange(event.currentTarget.value)}
          onKeyDown={event => {
            if (event.key === 'Enter') {
              event.preventDefault()
              if (event.shiftKey && onFindPrevious) onFindPrevious()
              else onFindNext()
            }
          }}
          placeholder="Find"
          aria-label="Find cells"
        />
        {findText.trim() !== '' && (
          <FindMatchCounter aria-live="polite" aria-label="Find match count">
            {findMatchCount > 0 && findActiveIndex >= 0
              ? `${findActiveIndex + 1} / ${findMatchCount}`
              : `0 / ${findMatchCount}`}
          </FindMatchCounter>
        )}
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
        {onFindPrevious && (
          <ToolIconButton
            title="Find previous"
            onClick={onFindPrevious}
            icon={ChevronUp}
          />
        )}
        <ToolIconButton title="Find next" onClick={onFindNext} icon={Search} />
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
        <ToolIconButton
          title="Close find and replace"
          icon={X}
          onClick={() => {
            setFindOpen(false)
            findTriggerRef.current?.focus()
          }}
        />
      </FindBox>
    </Toolbar>
  )
}

const selectChrome: Record<string, string> = { 'data-chrome': 'toolbar-select' }

const NumFormatWrap = styled.div`
  position: relative;
  display: inline-flex;
`

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
  border-radius: 2px;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-family: var(--platform-typography-font-family-mono);
  font-size: 9px;
  font-weight: 700;
  line-height: 1;
  pointer-events: none;
`

const Toolbar = styled.div`
  display: contents;
`

const ToolbarHeader = styled.div`
  grid-column: 2;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
`

const Sections = styled.div`
  display: flex;
  align-items: center;
  gap: 4px;
`

const SectionButton = styled(ToolTextButton)`
  && {
    border-radius: 0;
    padding: 0 10px;
    background: transparent;
    box-shadow: none;
    font-weight: 500;
    color: var(--pure-chrome-soft);
  }
  &&:hover {
    background: var(--n-hover);
    color: var(--sheets-ink);
  }
  &&[aria-selected='true'] {
    color: var(--sheets-ink);
    background: var(--p-chosen);
    box-shadow: var(--p-chosen-edge);
    font-weight: 600;
  }
  &:focus-visible {
    outline: 2px solid var(--sheets-accent);
    outline-offset: 2px;
  }
`

const ActiveDot = styled.span`
  width: 5px;
  height: 5px;
  border-radius: 50%;
  background: var(--sheets-accent);
`

const QuickActions = styled.div`
  display: flex;
  align-items: center;
  gap: 4px;
`

const ToolbarRow = styled.div`
  grid-column: 1 / -1;
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  align-items: center;
  min-height: 32px;
  &[hidden] {
    display: none;
  }
`

const ToolbarGroup = styled.div`
  display: inline-flex;
  flex-wrap: wrap;
  gap: 3px;
  align-items: center;
  min-width: 0;
`

const FindBox = styled.div`
  grid-column: 1 / -1;
  &[hidden] {
    display: none;
  }
  padding-top: 10px;
  border-top: 1px solid var(--pure-chrome-line);
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  align-items: center;
  min-width: 0;
`

const FindActionButton = styled(ToolTextButton)``

const FindMatchCounter = styled.span`
  padding: 0 6px;
  color: var(--pure-chrome-muted);
  font-family: var(--platform-typography-font-family-mono);
  font-size: var(--pure-chrome-meta-size);
  font-weight: 600;
  white-space: nowrap;
`

// A toolbar-height outlined input (the 28px toolbar select's box).
const FindInput = styled.input.attrs(selectChrome)`
  width: 140px;
  padding: 0 8px;
  font-weight: 400;

  @media (max-width: 720px) {
    width: 96px;
  }
`

const FilterControls = styled.div`
  display: inline-flex;
  flex-wrap: wrap;
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
