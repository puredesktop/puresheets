import { useMemo } from 'react'
import type React from 'react'
import type { ClipboardEvent, KeyboardEvent } from 'react'
import styled from 'styled-components'
import {
  cellKey,
  cellsInRange,
  columnLabel,
  displayCellValue,
  isCellLocked,
  isFormulaError,
  parseCellKey,
} from '../lib/workbookModel'
import { conditionalStyleForCell } from '../lib/sheetDataFeatures'
import type { CellStyle, WorkbookSheet } from '../types'

/**
 * Legacy custom React grid, retained as the fallback surface.
 *
 * SUPERSEDED (Phase S0): interactive editing and rendering now happen in the
 * Univer editor (`UniverSpreadsheetSurface`). This grid stays for:
 * - vitest runs, where the Univer canvas cannot mount (`MODE === 'test'`);
 * - a data-integrity fallback if the Univer surface fails to mount.
 * It renders through the legacy `evaluateFormula` compatibility shim in
 * `workbookModel.ts`. Retire it only after the Phase S0 exit gate (no data
 * loss on the Univer surface) is verified in production use.
 *
 * Phase S3 decorations rendered here: comment corner marks, validation
 * flags, protected-cell hinting, and conditional-format styles.
 */
export function CompatibilityGrid({
  sheet,
  visible,
  gridTemplate,
  visibleRows,
  selectedCell,
  normalizedRange,
  invalidCells,
  onPasteIntoCell,
  onPasteFromClipboard,
  onClearSelectedRange,
  onCopySelectedRange,
  onSelectCell,
  onSelectRange,
}: {
  sheet: WorkbookSheet
  visible: boolean
  gridTemplate: string
  visibleRows: number[]
  selectedCell: string
  normalizedRange: string | null
  /** Cell key → validation message, computed by the shell from the live
   *  workbook (the same map the Univer surface paints). */
  invalidCells: Map<string, string>
  onPasteIntoCell: (cell: string, text: string) => void
  onPasteFromClipboard: (cell?: string) => Promise<void>
  onClearSelectedRange: () => void
  onCopySelectedRange: (cut?: boolean) => Promise<void>
  onSelectCell: (cell: string) => void
  onSelectRange: (range: string) => void
}): React.ReactElement {
  const selectedHeaderCoordinates = useMemo(() => {
    const keys = normalizedRange
      ? cellsInRange(normalizedRange)
      : [selectedCell]
    const rows = new Set<number>()
    const columns = new Set<number>()
    keys.forEach(key => {
      const position = parseCellKey(key)
      if (!position) return
      rows.add(position.row)
      columns.add(position.column)
    })
    return { columns, rows }
  }, [normalizedRange, selectedCell])

  const conditionalStyles = useMemo(() => {
    const map = new Map<string, CellStyle>()
    for (const rule of sheet.conditionalFormats ?? []) {
      for (const key of cellsInRange(rule.range)) {
        if (map.has(key)) continue
        const style = conditionalStyleForCell(sheet, key)
        if (style) map.set(key, style)
      }
    }
    return map
  }, [sheet])
  const lockedCells = useMemo(() => {
    const locked = new Set<string>()
    for (const range of sheet.protection?.lockedRanges ?? []) {
      for (const key of cellsInRange(range)) locked.add(key)
    }
    return locked
  }, [sheet])

  return (
    <Grid
      $template={gridTemplate}
      $visible={visible}
      data-testid="compatibility-grid"
      role="grid"
      aria-hidden={!visible}
      aria-label={`${sheet.name} compatibility spreadsheet grid`}
    >
      <Corner />
      {Array.from({ length: sheet.columnCount }, (_, column) => (
        <HeaderCell
          key={column}
          $frozenColumn={Boolean(sheet.frozenColumns && column === 0)}
          $selected={selectedHeaderCoordinates.columns.has(column)}
          role="columnheader"
        >
          {columnLabel(column)}
        </HeaderCell>
      ))}
      {visibleRows.map(row => (
        <RowFragment
          key={row}
          row={row}
          sheet={sheet}
          selectedCell={selectedCell}
          selectedRowHeader={selectedHeaderCoordinates.rows.has(row)}
          frozenRow={Boolean(sheet.frozenRows && row === 0)}
          frozenColumn={Boolean(sheet.frozenColumns)}
          onPaste={onPasteIntoCell}
          onPasteFromClipboard={onPasteFromClipboard}
          onClearSelectedRange={onClearSelectedRange}
          onCopySelectedRange={onCopySelectedRange}
          onSelect={onSelectCell}
          onSelectRange={onSelectRange}
          selectedRange={normalizedRange}
          invalidCells={invalidCells}
          conditionalStyles={conditionalStyles}
          lockedCells={lockedCells}
        />
      ))}
    </Grid>
  )
}

function RowFragment({
  frozenColumn,
  frozenRow,
  row,
  sheet,
  selectedCell,
  selectedRowHeader,
  onPaste,
  onPasteFromClipboard,
  onClearSelectedRange,
  onCopySelectedRange,
  onSelect,
  onSelectRange,
  selectedRange,
  invalidCells,
  conditionalStyles,
  lockedCells,
}: {
  frozenColumn: boolean
  frozenRow: boolean
  row: number
  sheet: WorkbookSheet
  selectedCell: string
  selectedRowHeader: boolean
  selectedRange: string | null
  onPaste: (cell: string, text: string) => void
  onPasteFromClipboard: (cell?: string) => Promise<void>
  onClearSelectedRange: () => void
  onCopySelectedRange: (cut?: boolean) => Promise<void>
  onSelect: (cell: string) => void
  onSelectRange: (range: string) => void
  invalidCells: Map<string, string>
  conditionalStyles: Map<string, CellStyle>
  lockedCells: Set<string>
}): React.ReactElement {
  const selectedCells = selectedRange
    ? new Set(cellsInRange(selectedRange))
    : new Set<string>()

  return (
    <>
      <RowHeader $frozenRow={frozenRow} $selected={selectedRowHeader}>
        {row + 1}
      </RowHeader>
      {Array.from({ length: sheet.columnCount }, (_, column) => {
        const key = cellKey(row, column)
        const cell = sheet.cells[key]
        const displayValue = displayCellValue(sheet, key)
        const formulaError = isFormulaError(displayValue)
        const comment = sheet.comments?.[key]
        const invalidMessage = invalidCells.get(key)
        const locked = lockedCells.has(key) && isCellLocked(sheet, key)
        const style: CellStyle = {
          ...cell?.style,
          ...conditionalStyles.get(key),
        }
        const titleParts = [
          comment ? `Note: ${comment}` : undefined,
          invalidMessage ? `Invalid: ${invalidMessage}` : undefined,
          locked ? 'Protected' : undefined,
        ].filter((part): part is string => Boolean(part))
        return (
          <DataCell
            key={key}
            aria-label={key}
            title={titleParts.length ? titleParts.join('\n') : undefined}
            $selected={key === selectedCell}
            $inRange={selectedCells.has(key)}
            $error={formulaError}
            $bold={style.bold}
            $italic={style.italic}
            $underline={style.underline}
            $align={style.align}
            $fill={style.fillColor}
            $fontSize={style.fontSize}
            $height={sheet.rowHeights?.[String(row)]}
            $color={style.textColor}
            $bordered={style.border}
            $frozenRow={frozenRow}
            $frozenColumn={frozenColumn && column === 0}
            $hasComment={Boolean(comment)}
            $invalid={Boolean(invalidMessage)}
            $locked={locked}
            onClick={event => {
              if (event.shiftKey) onSelectRange(`${selectedCell}:${key}`)
              else onSelect(key)
            }}
            onKeyDown={(event: KeyboardEvent<HTMLButtonElement>) => {
              if (event.key === 'Delete' || event.key === 'Backspace') {
                event.preventDefault()
                onClearSelectedRange()
              }
              if (
                (event.metaKey || event.ctrlKey) &&
                event.key.toLowerCase() === 'c'
              ) {
                event.preventDefault()
                void onCopySelectedRange()
              }
              if (
                (event.metaKey || event.ctrlKey) &&
                event.key.toLowerCase() === 'x'
              ) {
                event.preventDefault()
                void onCopySelectedRange(true)
              }
              if (
                (event.metaKey || event.ctrlKey) &&
                event.key.toLowerCase() === 'v'
              ) {
                event.preventDefault()
                void onPasteFromClipboard(key)
              }
            }}
            onPaste={(event: ClipboardEvent<HTMLButtonElement>) => {
              const text = event.clipboardData.getData('text/plain')
              if (!text) return
              event.preventDefault()
              onPaste(key, text)
            }}
          >
            {displayValue}
          </DataCell>
        )
      })}
    </>
  )
}

const Grid = styled.div<{ $template: string; $visible: boolean }>`
  display: ${({ $visible }) => ($visible ? 'grid' : 'none')};
  grid-template-columns: 44px ${({ $template }) => $template};
  width: max-content;
  min-width: 100%;
`

const Corner = styled.div`
  position: sticky;
  top: 0;
  left: 0;
  z-index: 3;
  border-right: 1px solid var(--sheets-line);
  border-bottom: 1px solid var(--sheets-line);
  background: var(--sheets-surface);
`

const HeaderCell = styled.div<{ $frozenColumn?: boolean; $selected?: boolean }>`
  position: sticky;
  top: 0;
  left: ${({ $frozenColumn }) => ($frozenColumn ? '44px' : 'auto')};
  z-index: ${({ $frozenColumn }) => ($frozenColumn ? 4 : 2)};
  border-right: 1px solid var(--sheets-line);
  border-bottom: 1px solid var(--sheets-line);
  background: ${({ $selected }) =>
    $selected
      ? 'var(--app-bg, color-mix(in srgb, var(--app-acc, #007eaa) 10%, var(--sheets-surface)))'
      : 'var(--sheets-surface)'};
  color: ${({ $selected }) =>
    $selected ? 'var(--app-text, #005f80)' : 'var(--sheets-ink-faint)'};
  padding: 7px 8px;
  font-family: var(--sheets-mono);
  font-size: 11px;
  font-weight: 700;
  text-align: center;
`

const RowHeader = styled.div<{ $frozenRow?: boolean; $selected?: boolean }>`
  position: sticky;
  top: ${({ $frozenRow }) => ($frozenRow ? '32px' : 'auto')};
  left: 0;
  z-index: ${({ $frozenRow }) => ($frozenRow ? 4 : 1)};
  min-height: 32px;
  border-right: 1px solid var(--sheets-line);
  border-bottom: 1px solid var(--sheets-line);
  background: ${({ $selected }) =>
    $selected
      ? 'var(--app-bg, color-mix(in srgb, var(--app-acc, #007eaa) 10%, var(--sheets-surface)))'
      : 'var(--sheets-surface)'};
  padding: 7px 8px;
  color: ${({ $selected }) =>
    $selected ? 'var(--app-text, #005f80)' : 'var(--sheets-ink-faint)'};
  font-family: var(--sheets-mono);
  font-size: 11px;
  text-align: right;
`

const DataCell = styled.button<{
  $selected?: boolean
  $inRange?: boolean
  $error?: boolean
  $bold?: boolean
  $italic?: boolean
  $underline?: boolean
  $align?: 'left' | 'center' | 'right'
  $fill?: string
  $fontSize?: number
  $height?: number
  $color?: string
  $bordered?: boolean
  $frozenRow?: boolean
  $frozenColumn?: boolean
  $hasComment?: boolean
  $invalid?: boolean
  $locked?: boolean
}>`
  position: ${({ $frozenRow, $frozenColumn }) =>
    $frozenRow || $frozenColumn ? 'sticky' : 'relative'};
  top: ${({ $frozenRow }) => ($frozenRow ? '32px' : 'auto')};
  left: ${({ $frozenColumn }) => ($frozenColumn ? '44px' : 'auto')};
  z-index: ${({ $frozenRow, $frozenColumn }) =>
    $frozenRow && $frozenColumn ? 3 : $frozenRow || $frozenColumn ? 2 : 0};
  min-height: ${({ $height }) => $height ?? 32}px;
  border: 0;
  border-right: 1px solid
    ${({ $error }) => ($error ? '#e3a29b' : 'var(--sheets-line)')};
  border-bottom: 1px solid
    ${({ $error }) => ($error ? '#e3a29b' : 'var(--sheets-line)')};
  box-shadow: ${({ $bordered, $selected, $inRange }) => {
    if ($selected) return 'inset 0 0 0 2px var(--app-acc, #007eaa)'
    if ($inRange)
      return 'inset 0 0 0 1px color-mix(in srgb, var(--app-acc, #007eaa) 58%, transparent)'
    return $bordered ? 'inset 0 0 0 1px var(--sheets-line)' : 'none'
  }};
  outline: 0;
  background: ${({ $error, $fill, $selected, $inRange, $locked }) => {
    if ($error) return '#fff1f0'
    if ($fill) return $fill
    if ($selected || $inRange)
      return 'var(--app-bg, color-mix(in srgb, var(--app-acc, #007eaa) 9%, #ffffff))'
    if ($locked)
      return 'color-mix(in srgb, var(--sheets-surface) 55%, var(--sheets-canvas))'
    return 'var(--sheets-canvas)'
  }};
  color: ${({ $error, $color }) =>
    $error ? '#982d25' : $color ?? 'var(--sheets-ink)'};
  padding: 6px 8px;
  overflow: hidden;
  font-family: var(--sheets-mono);
  font-size: ${({ $fontSize }) => $fontSize ?? 12}px;
  font-style: ${({ $italic }) => ($italic ? 'italic' : 'normal')};
  font-weight: ${({ $error, $bold }) => ($error || $bold ? 700 : 400)};
  text-align: ${({ $align }) => $align ?? 'left'};
  text-decoration: ${({ $underline }) => ($underline ? 'underline' : 'none')};
  text-overflow: ellipsis;
  white-space: nowrap;

  /* Comment corner mark (quiet amber triangle, top-right). */
  ${({ $hasComment }) =>
    $hasComment
      ? `&::after {
          content: '';
          position: absolute;
          top: 0;
          right: 0;
          border-top: 6px solid #b98c1d;
          border-left: 6px solid transparent;
        }`
      : ''}

  /* Validation flag (dotted red underline along the cell bottom). */
  ${({ $invalid }) =>
    $invalid
      ? `&::before {
          content: '';
          position: absolute;
          right: 2px;
          bottom: 1px;
          left: 2px;
          border-bottom: 2px dotted #b3261e;
        }`
      : ''}
`
