import {
  agentToolErrorContent,
  formatAgentToolJson,
  readAgentToolStringArg,
} from '@purescience/platform-ui/bridge/agentToolHelpers'
import type { AgentToolHandlerResult } from '@purescience/platform-ui/bridge/react/usePlatformAgentTools'
import { usedRangeForSheet, workbookVersion } from '../lib/sheetsAgentContext'
import { getExchangeRate } from '../lib/exchangeRate'
import {
  addChartToSheet,
  addSheet,
  appendAgentLog,
  cellKey,
  deleteSheet,
  isCellLocked,
  parseRangeBounds,
  renameSheet,
  setRangeStyle,
  setColumnWidth,
  setRowHeight,
  parseCellKey,
  applyAgentCellChangesDirect,
  cellsInRange,
  computedCellValue,
  findMatches,
  normalizeRangeToken,
  setCellComment,
} from '../lib/workbookModel'
import { clampFontSize, normalizeHexColor } from '../lib/sheetMetadataNormalize'
import type {
  CellValueKind,
  PureSheetsDocument,
  WorkbookChart,
  WorkbookSheet,
} from '../types'
import {
  AgentSheetsToolError,
  type PureSheetsAtomicToolName,
  type SheetsAgentToolContext,
} from './catalog'

const CHART_TYPES: ReadonlySet<string> = new Set([
  'bar',
  'line',
  'area',
  'pie',
  'donut',
  'scatter',
])

const CELL_KINDS: ReadonlySet<string> = new Set([
  'text',
  'number',
  'formula',
  'blank',
])

const READ_CELL_CAP = 500
const CHANGE_CAP = 400
const FORMAT_CELL_CAP = 5_000
const MAX_COMMENT_LENGTH = 8_192

export async function autoFitRowsHandler(context: SheetsAgentToolContext, args: Record<string, unknown>): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const sheet = sheetByRef(document, readAgentToolStringArg(args, 'sheet'))
  if (!sheet) return agentToolErrorContent('No matching sheet. Read getSheetsContext.')
  if (sheet.protection?.lockedRanges?.length) return agentToolErrorContent('Row fitting affects entire rows and is refused on protected sheets.')
  const start = args.startRow, count = args.rowCount
  if (!Number.isInteger(start) || !Number.isInteger(count) || (start as number) < 1 || (count as number) < 1 || (count as number) > 100 || (start as number) + (count as number) - 1 > sheet.rowCount)
    return agentToolErrorContent('Provide a valid 1-based startRow and rowCount between 1 and 100 within the sheet.')
  const rows = Array.from({ length: count as number }, (_, i) => (start as number) - 1 + i)
  if (!await context.autoFitRows?.(sheet.id, rows)) return agentToolErrorContent('Live row measurement is unavailable; do not claim rows were fitted.')
  context.setDocument(current => appendAgentLog(current, {
    agentName: 'PureSheets Assistant', tool: 'autoFitRows',
    summary: `Fitted rows ${start}–${(start as number) + (count as number) - 1} using the live renderer.`,
    sheetId: sheet.id, sheetName: sheet.name,
  }))
  return { content: formatAgentToolJson({ sheet: sheet.name, startRow: start, rowCount: count, fitted: true }) }
}

export function requireDocument(
  context: SheetsAgentToolContext,
): PureSheetsDocument | { error: string } {
  if (!context.document) {
    return { error: 'The workbook is still loading; try again shortly.' }
  }
  return context.document
}

export function sheetByRef(
  document: PureSheetsDocument,
  ref: string | null | undefined,
): WorkbookSheet | undefined {
  if (!ref?.trim()) {
    return document.workbook.sheets.find(
      sheet => sheet.id === document.workbook.activeSheetId,
    )
  }
  const lowered = ref.trim().toLowerCase()
  return document.workbook.sheets.find(
    sheet => sheet.id === ref || sheet.name.toLowerCase() === lowered,
  )
}

export async function getSheetsContextHandler(
  context: SheetsAgentToolContext,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  return {
    content: formatAgentToolJson({
      title: document.metadata.title,
      filePath: context.filePath || null,
      version: await workbookVersion(document, context.filePath),
      activeSheetId: document.workbook.activeSheetId,
      selectedRange: context.selectedRange,
      sheets: document.workbook.sheets.map(sheet => ({
        id: sheet.id,
        name: sheet.name,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        usedRange: usedRangeForSheet(sheet),
        isAgentDraft: Boolean(sheet.agentDraft),
        lockedRanges: sheet.protection?.lockedRanges ?? [],
        charts: (sheet.charts ?? []).length,
      })),
    }),
  }
}

export async function readRangeHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const sheet = sheetByRef(document, readAgentToolStringArg(args, 'sheet'))
  if (!sheet) {
    return agentToolErrorContent(
      'No matching sheet. getSheetsContext lists them.',
    )
  }
  const rangeArg = readAgentToolStringArg(args, 'range')
  const range = rangeArg
    ? normalizeRangeToken(rangeArg)
    : usedRangeForSheet(sheet)
  if (!range) {
    return agentToolErrorContent(
      rangeArg
        ? `"${rangeArg}" is not a valid A1-style range.`
        : 'The sheet is empty and no range was given.',
    )
  }
  const [startKey, endKey] = range.split(':')
  const start = parseCellKey(startKey)!, end = parseCellKey(endKey)!
  const count = (end.row - start.row + 1) * (end.column - start.column + 1)
  if (count > READ_CELL_CAP) {
    return agentToolErrorContent(
      `${range} spans ${count} cells — read at most ${READ_CELL_CAP} at a time; narrow the range.`,
    )
  }
  const cells = cellsInRange(range)
    .filter(key => sheet.cells[key])
    .map(key => {
      const cell = sheet.cells[key]
      const computed = computedCellValue(sheet, key, document.workbook)
      return {
        cell: key,
        value: cell.value,
        kind: cell.kind,
        ...(computed !== cell.value ? { computed } : {}),
      }
    })
  return {
    content: formatAgentToolJson({
      version: await workbookVersion(document, context.filePath),
      filePath: context.filePath || null,
      sheetId: sheet.id,
      sheet: sheet.name,
      range,
      populated: cells.length,
      cells,
    }),
  }
}

export function findInWorkbookHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): AgentToolHandlerResult {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const query = readAgentToolStringArg(args, 'query')
  if (!query?.trim()) {
    return agentToolErrorContent('Pass "query": the text to find.')
  }
  const sheetRef = readAgentToolStringArg(args, 'sheet')
  const sheets = sheetRef
    ? [sheetByRef(document, sheetRef)].filter(
        (sheet): sheet is WorkbookSheet => Boolean(sheet),
      )
    : document.workbook.sheets
  if (sheetRef && sheets.length === 0) {
    return agentToolErrorContent(
      'No matching sheet. getSheetsContext lists them.',
    )
  }
  const matches: Array<{ sheet: string; cell: string; value: string }> = []
  for (const sheet of sheets) {
    for (const key of findMatches(sheet, query, {}, document.workbook)) {
      matches.push({
        sheet: sheet.name,
        cell: key,
        value: sheet.cells[key]?.value ?? '',
      })
      if (matches.length >= 50) break
    }
    if (matches.length >= 50) break
  }
  return {
    content: formatAgentToolJson({
      query,
      matches,
      ...(matches.length >= 50 ? { note: 'Capped at 50 matches.' } : {}),
    }),
  }
}

function readCellChangesArg(
  args: Record<string, unknown>,
): Array<{ cell: string; value: string; kind: CellValueKind }> | string {
  const raw = args.cellChanges
  if (!Array.isArray(raw) || raw.length === 0) {
    return 'Pass "cellChanges": an array of {cell, value, kind}.'
  }
  if (raw.length > CHANGE_CAP) {
    return `At most ${CHANGE_CAP} cell changes per transform; split larger work.`
  }
  const changes: Array<{ cell: string; value: string; kind: CellValueKind }> =
    []
  for (const item of raw) {
    if (!item || typeof item !== 'object') return 'Every cell change must be an object; nothing applied.'
    const candidate = item as Record<string, unknown>
    const cell = typeof candidate.cell === 'string' ? candidate.cell.trim() : ''
    if (typeof candidate.value !== 'string') return 'Every cell change needs a string value; nothing applied.'
    const value = candidate.value
    const kind =
      typeof candidate.kind === 'string' && CELL_KINDS.has(candidate.kind)
        ? (candidate.kind as CellValueKind)
        : undefined
    if (!cell || !kind) {
      return `Each cell change needs "cell" (A1 key) and "kind" (${[...CELL_KINDS].join(' | ')}).`
    }
    changes.push({ cell, value, kind })
  }
  return changes
}

export async function applyCellChangesHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const summary = readAgentToolStringArg(args, 'summary')
  if (!summary?.trim()) {
    return agentToolErrorContent(
      'Pass "summary": one sentence on what the change does and why.',
    )
  }
  if (!readAgentToolStringArg(args, 'sheet')?.trim()) return agentToolErrorContent('Provide the explicit target sheet id from getSheetsContext.')
  const sheet = sheetByRef(document, readAgentToolStringArg(args, 'sheet'))
  if (!sheet) {
    return agentToolErrorContent(
      'No matching sheet. getSheetsContext lists them.',
    )
  }
  const rangeArg = readAgentToolStringArg(args, 'range')
  const range =
    (rangeArg ? normalizeRangeToken(rangeArg) : usedRangeForSheet(sheet)) ??
    'A1:A1'
  const changes = readCellChangesArg(args)
  if (typeof changes === 'string') return agentToolErrorContent(changes)

  if (!rangeArg || !normalizeRangeToken(rangeArg)) return agentToolErrorContent('Provide an explicit valid target range.')
  const expectedVersion = readAgentToolStringArg(args, 'expectedVersion')
  if (!expectedVersion || expectedVersion !== await workbookVersion(document, context.filePath)) return agentToolErrorContent('Workbook changed or version missing. Read getSheetsContext/readRange again before writing.')
  const [startKey, endKey] = range.split(':')
  const start = parseCellKey(startKey)!, end = parseCellKey(endKey)!
  if (changes.some(change => {
    const position = parseCellKey(change.cell)
    return !position || position.row < start.row || position.row > end.row || position.column < start.column || position.column > end.column
  })) return agentToolErrorContent('Every changed cell must be inside the declared target range. Nothing applied.')
  let stale = false
  let applied = 0
  let skippedLocked = 0
  context.setDocument(current => {
    if (current !== document || (context.currentFilePath && context.currentFilePath() !== context.filePath)) { stale = true; return current }

    const result = applyAgentCellChangesDirect(
      current,
      sheet.id,
      range,
      summary.trim(),
      changes,
    )
    applied = result.applied
    skippedLocked = result.skippedLocked
    return result.document
  })
  if (stale) return agentToolErrorContent('Workbook changed during the write. Read it again; nothing applied.')
  if (!applied) {
    return agentToolErrorContent(
      skippedLocked
        ? `Nothing applied — ${skippedLocked} target cells are locked.`
        : 'Nothing applied — the changes were invalid or matched the current values.',
    )
  }
  return {
    content: formatAgentToolJson({
      sheet: sheet.name,
      applied,
      ...(skippedLocked ? { skippedLocked } : {}),
      note: 'Written to the sheet and recorded in the agent log with before/after.',
    }),
  }
}

export async function addChartHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const typeArg = readAgentToolStringArg(args, 'type') ?? 'bar'
  if (!CHART_TYPES.has(typeArg)) {
    return agentToolErrorContent(
      `type must be one of ${[...CHART_TYPES].join(', ')}.`,
    )
  }
  const sheet = sheetByRef(document, readAgentToolStringArg(args, 'sheet'))
  if (!sheet) {
    return agentToolErrorContent(
      'No matching sheet. getSheetsContext lists them.',
    )
  }
  const rangeArg = readAgentToolStringArg(args, 'range')
  const range = rangeArg ? normalizeRangeToken(rangeArg) : null
  if (!range) {
    return agentToolErrorContent('Pass "range": the A1 range to chart.')
  }
  const summary =
    readAgentToolStringArg(args, 'summary')?.trim() ||
    `Added a ${typeArg} chart from ${range}`
  let changed = false
  context.setDocument(current => {
    const next = addChartToSheet(current, sheet.id, range, typeArg as WorkbookChart['type'], summary)
    changed = next !== current
    return next
  })
  if (!changed) {
    return agentToolErrorContent(
      'The chart could not be built from that range — it needs labels and numeric values.',
    )
  }
  return {
    content: formatAgentToolJson({
      sheet: sheet.name,
      range,
      type: typeArg,
      note: 'Chart added and recorded in the agent log.',
    }),
  }
}

export async function addSheetHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const name = readAgentToolStringArg(args, 'name')?.trim()
  if (
    name &&
    document.workbook.sheets.some(
      sheet => sheet.name.toLowerCase() === name.toLowerCase(),
    )
  ) {
    return agentToolErrorContent(`A sheet named "${name}" already exists.`)
  }
  let created = ''
  let changed = false
  context.setDocument(current => {
    if (
      name &&
      current.workbook.sheets.some(
        sheet => sheet.name.toLowerCase() === name.toLowerCase(),
      )
    ) {
      return current
    }
    let next = addSheet(current, name || undefined)
    if (next === current) return current
    const sheet = next.workbook.sheets[next.workbook.sheets.length - 1]
    created = sheet?.name ?? ''
    if (sheet) {
      next = appendAgentLog(next, {
        agentName: 'PureSheets Assistant',
        tool: 'addSheet',
        summary: `Added sheet ${sheet.name}`,
        sheetId: sheet.id,
        sheetName: sheet.name,
      })
    }
    changed = true
    return next
  })
  if (!changed) {
    return agentToolErrorContent(
      name
        ? `A sheet named "${name}" already exists.`
        : 'Could not add sheet.',
    )
  }
  return {
    content: formatAgentToolJson({ sheet: created }),
  }
}

export async function setCellCommentHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const sheet = sheetByRef(document, readAgentToolStringArg(args, 'sheet'))
  if (!sheet) {
    return agentToolErrorContent(
      'No matching sheet. getSheetsContext lists them.',
    )
  }
  const cellArg = readAgentToolStringArg(args, 'cell')?.trim() ?? ''
  const parsedCell = !cellArg.includes(':') ? parseCellKey(cellArg) : null
  if (!parsedCell) {
    return agentToolErrorContent('Pass "cell": one A1 cell, e.g. "B4".')
  }
  const cell = cellKey(parsedCell.row, parsedCell.column)
  if (isCellLocked(sheet, cell)) {
    return agentToolErrorContent(
      `${cell} is protected — unlock the range before editing notes.`,
    )
  }
  const comment = readAgentToolStringArg(args, 'comment') ?? ''
  if (comment.length > MAX_COMMENT_LENGTH) {
    return agentToolErrorContent(
      `Comment exceeds the maximum length of ${MAX_COMMENT_LENGTH} characters.`,
    )
  }
  let changed = false
  let locked = false
  context.setDocument(current => {
    const target = current.workbook.sheets.find(
      candidate => candidate.id === sheet.id,
    )
    if (!target) return current
    if (isCellLocked(target, cell)) {
      locked = true
      return current
    }
    const updatedSheet = setCellComment(target, cell, comment)
    let next: PureSheetsDocument = {
      ...current,
      workbook: {
        ...current.workbook,
        sheets: current.workbook.sheets.map(candidate =>
          candidate.id === sheet.id ? updatedSheet : candidate,
        ),
      },
    }
    next = appendAgentLog(next, {
      agentName: 'PureSheets Assistant',
      tool: 'setCellComment',
      summary: comment.trim()
        ? `Noted ${cell} on ${sheet.name}`
        : `Cleared the note on ${cell} of ${sheet.name}`,
      sheetId: sheet.id,
      sheetName: sheet.name,
      range: cell,
    })
    changed = true
    return next
  })
  if (locked) {
    return agentToolErrorContent(
      `${cell} is protected — unlock the range before editing notes.`,
    )
  }
  if (!changed) {
    return agentToolErrorContent('Sheet is no longer available; nothing applied.')
  }
  return {
    content: formatAgentToolJson({
      sheet: sheet.name,
      cell,
      comment: comment.trim() || null,
    }),
  }
}

export function listSheetChangesHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): AgentToolHandlerResult {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const rawLimit = args.limit
  const limit =
    typeof rawLimit === 'number' && Number.isFinite(rawLimit)
      ? Math.max(1, Math.min(100, Math.floor(rawLimit)))
      : 25
  const entries = (document.workbook.agentLog ?? []).slice(0, limit)
  return {
    content: formatAgentToolJson({
      entries: entries.map(entry => ({
        at: entry.at,
        tool: entry.tool,
        agentName: entry.agentName,
        summary: entry.summary,
        sheet: entry.sheetName,
        ...(entry.range ? { range: entry.range } : {}),
        ...(entry.cells ? { cellsChanged: entry.cells.length } : {}),
      })),
      note: 'The agent log: every agent write with before/after, newest first. Check it before repeating work another agent already did.',
    }),
  }
}

const NUMBER_FORMATS: ReadonlySet<string> = new Set([
  'text',
  'number',
  'currency',
  'percent',
  'date',
])
const ALIGNMENTS: ReadonlySet<string> = new Set(['left', 'center', 'right'])

export async function formatRangeHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const sheet = sheetByRef(document, readAgentToolStringArg(args, 'sheet'))
  if (!sheet) {
    return agentToolErrorContent(
      'No matching sheet. getSheetsContext lists them.',
    )
  }
  const rangeArg = readAgentToolStringArg(args, 'range')
  const range = rangeArg ? normalizeRangeToken(rangeArg) : null
  const bounds = range ? parseRangeBounds(range) : null
  if (!range || !bounds) {
    return agentToolErrorContent('Pass "range": the A1 range to format.')
  }
  const spanCount =
    (bounds.end.row - bounds.start.row + 1) *
    (bounds.end.column - bounds.start.column + 1)
  if (!Number.isSafeInteger(spanCount) || spanCount > FORMAT_CELL_CAP) {
    return agentToolErrorContent(
      `${range} spans ${spanCount} cells — format at most ${FORMAT_CELL_CAP} cells at a time; narrow the range.`,
    )
  }
  const patch: Record<string, unknown> = {}
  if (args.wrap !== undefined && !['wrap', 'clip', 'overflow'].includes(args.wrap as string))
    return agentToolErrorContent('wrap must be wrap, clip or overflow.')
  if (['wrap', 'clip', 'overflow'].includes(args.wrap as string)) patch.wrap = args.wrap
  for (const [key, min, max] of [['columnWidth', 56, 360], ['rowHeight', 24, 160]] as const) {
    if (args[key] !== undefined && (typeof args[key] !== 'number' || !Number.isFinite(args[key]) || args[key] < min || args[key] > max))
      return agentToolErrorContent(`${key} must be a finite pixel size between ${min} and ${max}.`)
  }
  const resizing = args.columnWidth !== undefined || args.rowHeight !== undefined
  if (resizing && sheet.protection?.lockedRanges?.length)
    return agentToolErrorContent('Size changes affect entire rows/columns and are refused on protected sheets. Cell formatting can still be requested separately.')
  for (const key of ['bold', 'italic', 'underline', 'border'] as const) {
    if (typeof args[key] === 'boolean') patch[key] = args[key]
  }
  const fontSize = clampFontSize(args.fontSize)
  if (fontSize !== undefined) {
    patch.fontSize = fontSize
  }
  if (typeof args.decimals === 'number' && Number.isFinite(args.decimals)) {
    patch.decimals = Math.max(0, Math.min(8, Math.floor(args.decimals)))
  }
  for (const key of ['textColor', 'fillColor'] as const) {
    const normalizedColor = normalizeHexColor(args[key])
    if (normalizedColor) {
      patch[key] = normalizedColor
    }
  }
  const align = readAgentToolStringArg(args, 'align')
  if (align && ALIGNMENTS.has(align)) patch.align = align
  const numberFormat = readAgentToolStringArg(args, 'numberFormat')
  if (numberFormat && NUMBER_FORMATS.has(numberFormat)) {
    patch.numberFormat = numberFormat
  }
  if (!Object.keys(patch).length && !resizing) {
    return agentToolErrorContent(
      'Pass at least one formatting property: bold, italic, underline, border, fontSize, textColor, fillColor (#RRGGBB), align, numberFormat, decimals, wrap, columnWidth or rowHeight.',
    )
  }
  const summary =
    readAgentToolStringArg(args, 'summary')?.trim() ||
    `Formatted ${range} on ${sheet.name}`
  let changed = false
  context.setDocument(current => {
    const target = current.workbook.sheets.find(item => item.id === sheet.id)
    if (!target) return current
    if (resizing && target.protection?.lockedRanges?.length) return current
    let nextSheet = Object.keys(patch).length
      ? setRangeStyle(target, range, patch, FORMAT_CELL_CAP)
      : target
    if (typeof args.columnWidth === 'number') {
      for (let col = bounds.start.column; col <= bounds.end.column; col += 1) {
        nextSheet = setColumnWidth(nextSheet, col, args.columnWidth)
      }
    }
    if (typeof args.rowHeight === 'number') {
      for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
        nextSheet = setRowHeight(nextSheet, row, args.rowHeight)
      }
    }
    if (nextSheet === target) return current
    let next: PureSheetsDocument = {
      ...current,
      workbook: {
        ...current.workbook,
        sheets: current.workbook.sheets.map(item =>
          item.id === sheet.id ? nextSheet : item,
        ),
      },
    }
    next = appendAgentLog(next, {
      agentName: 'PureSheets Assistant',
      tool: 'formatRange',
      summary,
      sheetId: sheet.id,
      sheetName: sheet.name,
      range,
    })
    changed = true
    return next
  })
  if (!changed) {
    return agentToolErrorContent(
      'Formatting could not be applied (sheet missing or protected).',
    )
  }
  return {
    content: formatAgentToolJson({
      sheet: sheet.name,
      range,
      applied: { ...patch, ...(args.columnWidth === undefined ? {} : { columnWidth: args.columnWidth }), ...(args.rowHeight === undefined ? {} : { rowHeight: args.rowHeight }) },
      note: 'Formatting applied and recorded in the agent log.',
    }),
  }
}

export async function renameSheetHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const sheet = sheetByRef(document, readAgentToolStringArg(args, 'sheet'))
  if (!sheet) {
    return agentToolErrorContent(
      'No matching sheet. getSheetsContext lists them.',
    )
  }
  const name = readAgentToolStringArg(args, 'name')?.trim()
  if (!name) {
    return agentToolErrorContent('Pass "name": the new sheet name.')
  }
  if (
    document.workbook.sheets.some(
      candidate =>
        candidate.id !== sheet.id &&
        candidate.name.toLowerCase() === name.toLowerCase(),
    )
  ) {
    return agentToolErrorContent(`A sheet named "${name}" already exists.`)
  }
  const previous = sheet.name
  let changed = false
  context.setDocument(current => {
    let next = renameSheet(current, sheet.id, name)
    if (next === current) return current
    next = appendAgentLog(next, {
      agentName: 'PureSheets Assistant',
      tool: 'renameSheet',
      summary: `Renamed sheet ${previous} to ${name}`,
      sheetId: sheet.id,
      sheetName: name,
    })
    changed = true
    return next
  })
  if (!changed) {
    return agentToolErrorContent(
      previous === name
        ? `Sheet is already named "${name}".`
        : `Could not rename sheet "${previous}" to "${name}".`,
    )
  }
  return {
    content: formatAgentToolJson({ sheet: name, previous }),
  }
}

export async function deleteSheetHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const document = requireDocument(context)
  if ('error' in document) return agentToolErrorContent(document.error)
  const ref = readAgentToolStringArg(args, 'sheet')
  if (!ref?.trim()) {
    return agentToolErrorContent(
      'Pass "sheet": the sheet to delete, by name or id — never defaults to the active sheet.',
    )
  }
  const sheet = sheetByRef(document, ref)
  if (!sheet) {
    return agentToolErrorContent(
      'No matching sheet. getSheetsContext lists them.',
    )
  }
  const visibleSheets = document.workbook.sheets.filter(s => !s.agentDraft)
  if (!sheet.agentDraft && visibleSheets.length <= 1) {
    return agentToolErrorContent('The last sheet of a workbook cannot be deleted.')
  }
  const populated = Object.keys(sheet.cells).length
  let changed = false
  context.setDocument(current => {
    let next = deleteSheet(current, sheet.id)
    if (next === current) return current
    next = appendAgentLog(next, {
      agentName: 'PureSheets Assistant',
      tool: 'deleteSheet',
      summary: `Deleted sheet ${sheet.name} (${populated} populated cells)`,
      sheetId: sheet.id,
      sheetName: sheet.name,
    })
    changed = true
    return next
  })
  if (!changed) {
    return agentToolErrorContent('Could not delete that sheet.')
  }
  return {
    content: formatAgentToolJson({
      deleted: sheet.name,
      populatedCells: populated,
    }),
  }
}

/** Fetches dated source data; the drawer computes with it using normal sheet tools. */
export async function getExchangeRateHandler(_context: SheetsAgentToolContext, args: Record<string, unknown>): Promise<AgentToolHandlerResult> {
  try {
    const result = await getExchangeRate(readAgentToolStringArg(args, 'from') ?? '', readAgentToolStringArg(args, 'to') ?? '')
    return { content: formatAgentToolJson(result) }
  } catch (error) {
    return agentToolErrorContent(error instanceof Error ? error.message : 'Exchange-rate lookup failed. No workbook changes were made.')
  }
}

export async function createWorkbookHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const title = args.title
  if (typeof title !== 'string' || !title.trim()) {
    throw new AgentSheetsToolError('A non-empty workbook title is required.')
  }
  const create = context.createWorkbook
  if (!create) {
    throw new AgentSheetsToolError('Workbook creation is unavailable.')
  }
  return { content: formatAgentToolJson(await create(title.trim())) }
}

export async function openWorkbookHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const path = args.path
  if (typeof path !== 'string' || !path.trim()) {
    throw new AgentSheetsToolError('Provide the selected workbook path.')
  }
  const open = context.openWorkbook
  if (!open) {
    throw new AgentSheetsToolError('Workbook opening is unavailable.')
  }
  return { content: formatAgentToolJson(await open(path.trim())) }
}

export async function saveWorkbookAsXlsxHandler(
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
): Promise<AgentToolHandlerResult> {
  const path = args.path
  if (typeof path !== 'string' || !path.trim()) {
    throw new AgentSheetsToolError(
      'Provide the full path of the .xlsx file to write.',
    )
  }
  const save = context.saveAsXlsx
  if (!save) {
    throw new AgentSheetsToolError('Saving as Excel is unavailable.')
  }
  return {
    content: formatAgentToolJson(
      await save(path.trim(), args.overwrite === true),
    ),
  }
}

export async function completeWorkbookHandler(
  context: SheetsAgentToolContext,
): Promise<AgentToolHandlerResult> {
  const complete = context.completeWorkbook
  if (!complete) {
    throw new AgentSheetsToolError('Workbook saving is unavailable.')
  }
  return { content: formatAgentToolJson(await complete()) }
}

export type PureSheetsAtomicRunner = (
  context: SheetsAgentToolContext,
  args: Record<string, unknown>,
) => Promise<AgentToolHandlerResult> | AgentToolHandlerResult

export const PURESHEETS_ATOMIC_RUNNERS: Record<
  PureSheetsAtomicToolName,
  PureSheetsAtomicRunner
> = {
  getSheetsContext: context => getSheetsContextHandler(context),
  readRange: (context, args) => readRangeHandler(context, args),
  applyCellChanges: (context, args) => applyCellChangesHandler(context, args),
  completeWorkbook: context => completeWorkbookHandler(context),
  openWorkbook: (context, args) => openWorkbookHandler(context, args),
  createWorkbook: (context, args) => createWorkbookHandler(context, args),
  saveWorkbookAsXlsx: (context, args) =>
    saveWorkbookAsXlsxHandler(context, args),
  listSheetChanges: (context, args) => listSheetChangesHandler(context, args),
  findInWorkbook: (context, args) => findInWorkbookHandler(context, args),
  addSheet: (context, args) => addSheetHandler(context, args),
  renameSheet: (context, args) => renameSheetHandler(context, args),
  deleteSheet: (context, args) => deleteSheetHandler(context, args),
  formatRange: (context, args) => formatRangeHandler(context, args),
  autoFitRows: (context, args) => autoFitRowsHandler(context, args),
  setCellComment: (context, args) => setCellCommentHandler(context, args),
  addChart: (context, args) => addChartHandler(context, args),
  getExchangeRate: (context, args) => getExchangeRateHandler(context, args),
}
