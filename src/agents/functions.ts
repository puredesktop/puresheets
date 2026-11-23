import {
  defineAppFunction,
  type AppFunctionDefinition,
} from '@purescience/platform-ui/bridge/appFunctions'
import { invalidCellFlagsForSheet } from '../lib/sheetDataFeatures'
import { clampFontSize, normalizeHexColor } from '../lib/sheetMetadataNormalize'
import { usedRangeForSheet, workbookVersion } from '../lib/sheetsAgentContext'
import {
  addChartToSheet,
  addSheet,
  appendAgentLog,
  applyAgentCellChangesDirect,
  cellKey,
  columnIndex,
  columnLabel,
  computedCellValue,
  deriveCellKind,
  isCellLocked,
  isFormulaError,
  normalizeRangeToken,
  parseCellKey,
  parseRangeBounds,
  setCellComment,
  setColumnWidth,
  setRangeStyle,
  setRowHeight,
} from '../lib/workbookModel'
import type {
  AgentCellChange,
  CellStyle,
  PureSheetsDocument,
  WorkbookChart,
  WorkbookSheet,
} from '../types'
import { AgentSheetsToolError, type SheetsAgentToolContext } from './catalog'
import { requireDocument, sheetByRef } from './handlers'

const PROFILE_CELL_CAP = 25_000
const BATCH_CHANGE_CAP = 400
const FORMAT_CELL_CAP = 5_000
const MAX_COMMENT_LENGTH = 8_192

const CHART_TYPES: ReadonlySet<string> = new Set([
  'bar',
  'line',
  'area',
  'pie',
  'donut',
  'scatter',
])

const NUMBER_FORMATS: ReadonlySet<string> = new Set([
  'text',
  'number',
  'currency',
  'percent',
  'date',
])

const ALIGNMENTS: ReadonlySet<string> = new Set(['left', 'center', 'right'])
const WRAPS: ReadonlySet<string> = new Set(['wrap', 'clip', 'overflow'])

function assertDocument(ctx: SheetsAgentToolContext): PureSheetsDocument {
  const doc = requireDocument(ctx)
  if ('error' in doc) {
    throw new AgentSheetsToolError(doc.error)
  }
  return doc
}

function readString(
  obj: Record<string, unknown>,
  key: string,
): string | undefined {
  const val = obj[key]
  return typeof val === 'string' && val.trim() ? val.trim() : undefined
}

function clampInt(
  raw: unknown,
  fallback: number,
  min: number,
  max: number,
): number {
  if (typeof raw !== 'number' || !Number.isFinite(raw)) return fallback
  return Math.max(min, Math.min(max, Math.floor(raw)))
}

function parseNumericDisplay(raw: string): number | null {
  const trimmed = raw.trim()
  if (!trimmed) return null
  const isPercent = trimmed.endsWith('%')
  const stripped = (isPercent ? trimmed.slice(0, -1) : trimmed)
    .replace(/[$€£¥,\s]/g, '')
    .trim()
  if (!stripped) return null
  const n = Number(stripped)
  if (!Number.isFinite(n)) return null
  return isPercent ? n / 100 : n
}

function looksLikeNumericString(raw: string): boolean {
  const trimmed = raw.trim()
  if (!trimmed) return false
  if (!/^[-+]?[$€£¥]?\s*\d[\d,]*(?:\.\d+)?%?$/.test(trimmed)) return false
  return parseNumericDisplay(trimmed) !== null
}

function roundStat(value: number): number {
  return Math.round(value * 1_000_000) / 1_000_000
}

function extractReferencedSheetNames(formula: string): string[] {
  const refs: string[] = []
  const parts = formula.split(/("(?:[^"]|"")*")/)
  for (let i = 0; i < parts.length; i += 2) {
    const segment = parts[i]
    const regex =
      /'((?:[^']|'')+)'!|(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)!/g
    let match: RegExpExecArray | null
    while ((match = regex.exec(segment)) !== null) {
      const name =
        match[1] !== undefined
          ? match[1].replace(/''/g, "'")
          : match[2] ?? ''
      if (name) refs.push(name)
    }
  }
  return refs
}

function parseFormatSpec(
  raw: Record<string, unknown>,
  label: string,
): {
  patch: Partial<CellStyle>
  columnWidth?: number
  rowHeight?: number
} {
  const patch: Partial<CellStyle> = {}
  if (raw.wrap !== undefined) {
    if (typeof raw.wrap !== 'string' || !WRAPS.has(raw.wrap)) {
      throw new AgentSheetsToolError(
        `${label}: wrap must be wrap, clip, or overflow.`,
      )
    }
    patch.wrap = raw.wrap as CellStyle['wrap']
  }
  let columnWidth: number | undefined
  let rowHeight: number | undefined
  for (const [key, min, max] of [
    ['columnWidth', 56, 360],
    ['rowHeight', 24, 160],
  ] as const) {
    if (raw[key] !== undefined) {
      if (
        typeof raw[key] !== 'number' ||
        !Number.isFinite(raw[key]) ||
        raw[key] < min ||
        raw[key] > max
      ) {
        throw new AgentSheetsToolError(
          `${label}: ${key} must be a finite pixel size between ${min} and ${max}.`,
        )
      }
      if (key === 'columnWidth') columnWidth = raw[key]
      else rowHeight = raw[key]
    }
  }
  for (const key of ['bold', 'italic', 'underline', 'border'] as const) {
    if (typeof raw[key] === 'boolean') patch[key] = raw[key]
  }
  const fontSize = clampFontSize(raw.fontSize)
  if (fontSize !== undefined) patch.fontSize = fontSize
  if (typeof raw.decimals === 'number' && Number.isFinite(raw.decimals)) {
    patch.decimals = Math.max(0, Math.min(8, Math.floor(raw.decimals)))
  }
  for (const key of ['textColor', 'fillColor'] as const) {
    if (raw[key] !== undefined) {
      const normalized = normalizeHexColor(raw[key])
      if (!normalized) {
        throw new AgentSheetsToolError(
          `${label}: ${key} must be a valid hex color (#RGB or #RRGGBB).`,
        )
      }
      patch[key] = normalized
    }
  }
  if (raw.align !== undefined) {
    if (typeof raw.align !== 'string' || !ALIGNMENTS.has(raw.align)) {
      throw new AgentSheetsToolError(
        `${label}: align must be left, center, or right.`,
      )
    }
    patch.align = raw.align as CellStyle['align']
  }
  if (raw.numberFormat !== undefined) {
    if (
      typeof raw.numberFormat !== 'string' ||
      !NUMBER_FORMATS.has(raw.numberFormat)
    ) {
      throw new AgentSheetsToolError(
        `${label}: numberFormat must be text, number, currency, percent, or date.`,
      )
    }
    patch.numberFormat = raw.numberFormat as CellStyle['numberFormat']
  }
  return { patch, columnWidth, rowHeight }
}

function visibleSheets(doc: PureSheetsDocument): WorkbookSheet[] {
  return doc.workbook.sheets.filter(s => !s.agentDraft)
}

function resolveTargetSheets(
  doc: PureSheetsDocument,
  sheetRef: string | undefined,
): WorkbookSheet[] {
  if (!sheetRef) return visibleSheets(doc)
  const found = sheetByRef(doc, sheetRef)
  if (!found) {
    throw new AgentSheetsToolError(
      `No sheet matching "${sheetRef}". Available sheets: ${visibleSheets(doc)
        .map(s => s.name)
        .join(', ')}.`,
    )
  }
  return [found]
}

// ---------------------------------------------------------------------------
// 1. fn_workbookBriefing
// ---------------------------------------------------------------------------

export const fn_workbookBriefing = defineAppFunction<SheetsAgentToolContext>({
  name: 'fn_workbookBriefing',
  category: 'workbook',
  risk: 'read',
  requiresApproval: false,
  replaces: ['getSheetsContext', 'readRange', 'listSheetChanges'],
  description:
    'Inspect the open workbook in one call: returns title, version token for expectedVersion, active sheet, selection, named ranges, per-sheet structural metrics (used range, formula/error/validation/comment/chart counts, locked ranges, headers, and computed preview rows), and recent agent log entries. Replaces getSheetsContext + multiple readRange + listSheetChanges round-trips.',
  inputSchema: {
    type: 'object',
    properties: {
      sheet: {
        type: 'string',
        description:
          'Optional sheet name or id to brief. Omit to brief every visible sheet in the workbook.',
      },
      previewRows: {
        type: 'integer',
        minimum: 0,
        maximum: 20,
        description:
          'Number of leading rows per sheet to preview with computed values (default 5, max 20).',
      },
      maxPreviewColumns: {
        type: 'integer',
        minimum: 1,
        maximum: 20,
        description:
          'Maximum columns included in each sheet preview row (default 12, max 20).',
      },
      recentChangesLimit: {
        type: 'integer',
        minimum: 0,
        maximum: 25,
        description:
          'Number of recent agent log entries to include (default 5, max 25).',
      },
    },
  },
  run: async (ctx, args) => {
    const document = assertDocument(ctx)
    const sheetRef = readString(args, 'sheet')
    const previewRows = clampInt(args.previewRows, 5, 0, 20)
    const maxPreviewColumns = clampInt(args.maxPreviewColumns, 12, 1, 20)
    const recentChangesLimit = clampInt(args.recentChangesLimit, 5, 0, 25)
    const targets = resolveTargetSheets(document, sheetRef)
    const version = await workbookVersion(document, ctx.filePath)

    let totalPopulated = 0
    let totalFormulas = 0
    let totalErrors = 0
    let totalValidationFlags = 0

    const sheetBriefs = targets.map(sheet => {
      const usedRange = usedRangeForSheet(sheet)
      const bounds = usedRange ? parseRangeBounds(usedRange) : null
      let populatedCells = 0
      let formulaCount = 0
      let errorCount = 0
      for (const [key, cell] of Object.entries(sheet.cells)) {
        if (!cell.value.trim()) continue
        populatedCells += 1
        if (cell.kind === 'formula' || cell.value.trim().startsWith('=')) {
          formulaCount += 1
          const computed = computedCellValue(sheet, key, document.workbook)
          if (isFormulaError(computed)) errorCount += 1
        }
      }
      const validationFlagCount = invalidCellFlagsForSheet(
        sheet,
        document.workbook,
      ).length
      totalPopulated += populatedCells
      totalFormulas += formulaCount
      totalErrors += errorCount
      totalValidationFlags += validationFlagCount

      const headers: Array<{ col: string; cell: string; value: string }> = []
      const preview: Array<{ row: number; values: Record<string, string> }> = []
      if (bounds) {
        const endCol = Math.min(
          bounds.end.column,
          bounds.start.column + maxPreviewColumns - 1,
        )
        for (let c = bounds.start.column; c <= endCol; c += 1) {
          const k = cellKey(bounds.start.row, c)
          const raw = sheet.cells[k]?.value
          if (raw?.trim()) {
            headers.push({
              col: columnLabel(c),
              cell: k,
              value: computedCellValue(sheet, k, document.workbook),
            })
          }
        }
        const previewEndRow = Math.min(
          bounds.end.row,
          bounds.start.row + previewRows - 1,
        )
        for (let r = bounds.start.row; r <= previewEndRow; r += 1) {
          const rowValues: Record<string, string> = {}
          for (let c = bounds.start.column; c <= endCol; c += 1) {
            const k = cellKey(r, c)
            if (sheet.cells[k]?.value.trim()) {
              rowValues[columnLabel(c)] = computedCellValue(
                sheet,
                k,
                document.workbook,
              )
            }
          }
          if (Object.keys(rowValues).length > 0) {
            preview.push({ row: r + 1, values: rowValues })
          }
        }
      }

      return {
        id: sheet.id,
        name: sheet.name,
        usedRange,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        populatedCells,
        formulaCount,
        errorCount,
        validationFlagCount,
        commentCount: Object.keys(sheet.comments ?? {}).length,
        chartCount: (sheet.charts ?? []).length,
        mergeCount: (sheet.merges ?? []).length,
        lockedRanges: sheet.protection?.lockedRanges ?? [],
        frozenRows: sheet.frozenRows ?? 0,
        frozenColumns: sheet.frozenColumns ?? 0,
        headers,
        preview,
      }
    })

    const recentChanges = (document.workbook.agentLog ?? [])
      .slice(0, recentChangesLimit)
      .map(entry => ({
        at: entry.at,
        tool: entry.tool,
        summary: entry.summary,
        sheet: entry.sheetName,
        ...(entry.range ? { range: entry.range } : {}),
        ...(entry.cells ? { cellsChanged: entry.cells.length } : {}),
      }))

    const namedRanges = (document.workbook.namedRanges ?? []).map(nr => ({
      name: nr.name,
      sheetId: nr.sheetId,
      range: nr.range,
    }))

    return {
      mutated: false,
      summary: `Briefed "${document.metadata.title}": ${targets.length} sheet(s), ${totalPopulated} populated cells (${totalFormulas} formulas, ${totalErrors} errors, ${totalValidationFlags} validation flags).`,
      data: {
        title: document.metadata.title,
        bound: Boolean(ctx.filePath),
        filePath: ctx.filePath || null,
        version,
        activeSheetId: document.workbook.activeSheetId,
        selectedRange: ctx.selectedRange,
        namedRanges,
        totals: {
          sheets: targets.length,
          populatedCells: totalPopulated,
          formulas: totalFormulas,
          formulaErrors: totalErrors,
          validationFlags: totalValidationFlags,
        },
        sheets: sheetBriefs,
        recentChanges,
      },
    }
  },
})

// ---------------------------------------------------------------------------
// 2. fn_auditWorkbookFormulasAndData
// ---------------------------------------------------------------------------

export const fn_auditWorkbookFormulasAndData =
  defineAppFunction<SheetsAgentToolContext>({
    name: 'fn_auditWorkbookFormulasAndData',
    category: 'audit',
    risk: 'read',
    requiresApproval: false,
    replaces: ['getSheetsContext', 'readRange', 'findInWorkbook'],
    description:
      'Audit formulas, cross-sheet references, validation rules, and data quality across the workbook (or one sheet) in a single pass. Detects formula error values (#DIV/0!, #REF!, #NAME?, #VALUE!, #N/A, #NUM!), broken cross-sheet references, validation rule violations, hardcoded numeric overrides inside formula columns, and numeric strings stored as text.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description:
            'Optional sheet name or id to audit. Omit to audit every visible sheet in the workbook.',
        },
        maxFindings: {
          type: 'integer',
          minimum: 1,
          maximum: 200,
          description:
            'Maximum findings returned per category (default 50, max 200).',
        },
      },
    },
    run: async (ctx, args) => {
      const document = assertDocument(ctx)
      const targets = resolveTargetSheets(document, readString(args, 'sheet'))
      const maxFindings = clampInt(args.maxFindings, 50, 1, 200)
      const version = await workbookVersion(document, ctx.filePath)

      const knownSheetNames = new Set(
        document.workbook.sheets.map(s => s.name.toLowerCase()),
      )

      const formulaErrors: Array<{
        sheet: string
        cell: string
        formula: string
        computed: string
      }> = []
      const errorBreakdown: Record<string, number> = {}
      const brokenSheetRefs: Array<{
        sheet: string
        cell: string
        formula: string
        missingSheet: string
      }> = []
      const validationViolations: Array<{
        sheet: string
        cell: string
        value: string
        message: string
      }> = []
      const hardcodedOverrides: Array<{
        sheet: string
        cell: string
        value: string
        column: string
        columnFormulaCount: number
        columnNumericCount: number
      }> = []
      const numericTextCells: Array<{
        sheet: string
        cell: string
        value: string
        column: string
      }> = []
      const crossSheetEdges = new Map<string, number>()

      let totalFormulas = 0
      let totalFormulaErrors = 0
      let totalBrokenRefs = 0
      let totalValidationViolations = 0
      let totalHardcodedOverrides = 0
      let totalNumericTextCells = 0

      for (const sheet of targets) {
        // 1. Validation violations
        for (const flag of invalidCellFlagsForSheet(sheet, document.workbook)) {
          totalValidationViolations += 1
          if (validationViolations.length < maxFindings) {
            validationViolations.push({
              sheet: sheet.name,
              cell: flag.cell,
              value: flag.value,
              message: flag.message,
            })
          }
        }

        // Track per-column composition for anomaly detection
        const usedRange = usedRangeForSheet(sheet)
        const bounds = usedRange ? parseRangeBounds(usedRange) : null
        const headerRow =
          bounds && bounds.end.row > bounds.start.row ? bounds.start.row : -1

        const colBuckets = new Map<
          number,
          {
            formulas: number
            numbers: Array<{ key: string; value: string }>
            numericText: Array<{ key: string; value: string }>
            totalPopulated: number
          }
        >()

        for (const [key, cell] of Object.entries(sheet.cells)) {
          const raw = cell.value.trim()
          if (!raw) continue
          const pos = parseCellKey(key)
          const isFormula = cell.kind === 'formula' || raw.startsWith('=')

          if (isFormula) {
            totalFormulas += 1
            const computed = computedCellValue(sheet, key, document.workbook)
            if (isFormulaError(computed)) {
              totalFormulaErrors += 1
              const errCode = computed.trim().toUpperCase()
              errorBreakdown[errCode] = (errorBreakdown[errCode] ?? 0) + 1
              if (formulaErrors.length < maxFindings) {
                formulaErrors.push({
                  sheet: sheet.name,
                  cell: key,
                  formula: cell.value,
                  computed,
                })
              }
            }

            const refs = extractReferencedSheetNames(raw)
            for (const targetName of refs) {
              const lower = targetName.toLowerCase()
              if (!knownSheetNames.has(lower)) {
                totalBrokenRefs += 1
                if (brokenSheetRefs.length < maxFindings) {
                  brokenSheetRefs.push({
                    sheet: sheet.name,
                    cell: key,
                    formula: cell.value,
                    missingSheet: targetName,
                  })
                }
              } else if (lower !== sheet.name.toLowerCase()) {
                const edgeKey = `${sheet.name} -> ${targetName}`
                crossSheetEdges.set(
                  edgeKey,
                  (crossSheetEdges.get(edgeKey) ?? 0) + 1,
                )
              }
            }
          }

          if (pos && pos.row !== headerRow) {
            let bucket = colBuckets.get(pos.column)
            if (!bucket) {
              bucket = {
                formulas: 0,
                numbers: [],
                numericText: [],
                totalPopulated: 0,
              }
              colBuckets.set(pos.column, bucket)
            }
            bucket.totalPopulated += 1
            if (isFormula) {
              bucket.formulas += 1
            } else if (cell.kind === 'number') {
              bucket.numbers.push({ key, value: cell.value })
            } else if (cell.kind === 'text' && looksLikeNumericString(raw)) {
              bucket.numericText.push({ key, value: cell.value })
            }
          }
        }

        // Detect hardcoded numeric overrides inside formula-dominant columns & numeric strings in numeric columns
        for (const [colIdx, bucket] of colBuckets.entries()) {
          const colLetter = columnLabel(colIdx)
          const numericOrFormula = bucket.formulas + bucket.numbers.length
          if (
            bucket.formulas >= 3 &&
            bucket.numbers.length > 0 &&
            bucket.formulas / numericOrFormula >= 0.6
          ) {
            for (const item of bucket.numbers) {
              totalHardcodedOverrides += 1
              if (hardcodedOverrides.length < maxFindings) {
                hardcodedOverrides.push({
                  sheet: sheet.name,
                  cell: item.key,
                  value: item.value,
                  column: colLetter,
                  columnFormulaCount: bucket.formulas,
                  columnNumericCount: bucket.numbers.length,
                })
              }
            }
          }
          if (
            numericOrFormula >= 3 &&
            bucket.numericText.length > 0 &&
            numericOrFormula / bucket.totalPopulated >= 0.6
          ) {
            for (const item of bucket.numericText) {
              totalNumericTextCells += 1
              if (numericTextCells.length < maxFindings) {
                numericTextCells.push({
                  sheet: sheet.name,
                  cell: item.key,
                  value: item.value,
                  column: colLetter,
                })
              }
            }
          }
        }
      }

      const healthy =
        totalFormulaErrors === 0 &&
        totalBrokenRefs === 0 &&
        totalValidationViolations === 0

      return {
        mutated: false,
        summary: healthy
          ? `Audit passed across ${targets.length} sheet(s) (${totalFormulas} formulas): 0 formula errors, 0 broken sheet refs, 0 validation violations (${totalHardcodedOverrides} hardcoded override warning(s)).`
          : `Audit found ${totalFormulaErrors} formula error(s), ${totalBrokenRefs} broken sheet ref(s), ${totalValidationViolations} validation violation(s), and ${totalHardcodedOverrides} hardcoded override warning(s).`,
        data: {
          version,
          sheetsAudited: targets.map(s => s.name),
          healthy,
          counts: {
            formulas: totalFormulas,
            formulaErrors: totalFormulaErrors,
            brokenSheetRefs: totalBrokenRefs,
            validationViolations: totalValidationViolations,
            hardcodedOverrides: totalHardcodedOverrides,
            numericTextCells: totalNumericTextCells,
          },
          errorBreakdown,
          formulaErrors,
          brokenSheetRefs,
          validationViolations,
          hardcodedOverrides,
          numericTextCells,
          crossSheetDependencies: [...crossSheetEdges.entries()].map(
            ([edge, referenceCount]) => ({ edge, referenceCount }),
          ),
        },
      }
    },
  })

// ---------------------------------------------------------------------------
// 3. fn_profileSheetData
// ---------------------------------------------------------------------------

export const fn_profileSheetData = defineAppFunction<SheetsAgentToolContext>({
  name: 'fn_profileSheetData',
  category: 'analysis',
  risk: 'read',
  requiresApproval: false,
  replaces: ['readRange', 'findInWorkbook', 'getSheetsContext'],
  description:
    'Profile a sheet or A1 range (up to 25,000 cells) deterministically in one call: infers column types, non-empty/blank/distinct counts, exact numeric statistics (min, max, sum, mean, median), top categorical values, and optional grouped rollup by any column without writing helper formulas.',
  inputSchema: {
    type: 'object',
    properties: {
      sheet: {
        type: 'string',
        description:
          'Sheet name or id to profile. Defaults to the active sheet.',
      },
      range: {
        type: 'string',
        description:
          'Optional A1 range to profile. Omit to profile the full used range of the sheet.',
      },
      hasHeader: {
        type: 'boolean',
        description:
          'Whether the first row of the range contains column headers (default true).',
      },
      groupByColumn: {
        type: 'string',
        description:
          'Optional column letter (e.g. "A") or header name to compute a grouped rollup.',
      },
      valueColumn: {
        type: 'string',
        description:
          'Optional numeric column letter or header name to aggregate per group when groupByColumn is set (defaults to the first numeric column).',
      },
      topValuesLimit: {
        type: 'integer',
        minimum: 1,
        maximum: 20,
        description:
          'Maximum frequent values listed per column (default 5, max 20).',
      },
    },
  },
  run: async (ctx, args) => {
    const document = assertDocument(ctx)
    const sheet = sheetByRef(document, readString(args, 'sheet'))
    if (!sheet) {
      throw new AgentSheetsToolError(
        'No matching sheet. Call getSheetsContext or fn_workbookBriefing to list sheets.',
      )
    }
    const rangeArg = readString(args, 'range')
    const range = rangeArg
      ? normalizeRangeToken(rangeArg)
      : usedRangeForSheet(sheet)
    if (!range) {
      throw new AgentSheetsToolError(
        rangeArg
          ? `"${rangeArg}" is not a valid A1 range.`
          : `Sheet "${sheet.name}" is empty and no range was provided.`,
      )
    }
    const bounds = parseRangeBounds(range)!
    const rowSpan = bounds.end.row - bounds.start.row + 1
    const colSpan = bounds.end.column - bounds.start.column + 1
    const cellCount = rowSpan * colSpan
    if (cellCount > PROFILE_CELL_CAP) {
      throw new AgentSheetsToolError(
        `${range} spans ${cellCount} cells, exceeding the ${PROFILE_CELL_CAP}-cell profile cap. Narrow the range.`,
      )
    }

    const hasHeader = args.hasHeader !== false && rowSpan > 1
    const dataStartRow = hasHeader ? bounds.start.row + 1 : bounds.start.row
    const dataRowCount = bounds.end.row - dataStartRow + 1
    const topValuesLimit = clampInt(args.topValuesLimit, 5, 1, 20)
    const version = await workbookVersion(document, ctx.filePath)

    interface ColumnMeta {
      colIndex: number
      column: string
      header: string
      numericCount: number
    }

    const columnMetas: ColumnMeta[] = []
    const columns = []

    for (let c = bounds.start.column; c <= bounds.end.column; c += 1) {
      const colLetter = columnLabel(c)
      const headerCellKey = cellKey(bounds.start.row, c)
      const header = hasHeader
        ? computedCellValue(sheet, headerCellKey, document.workbook).trim() ||
          `Column ${colLetter}`
        : `Column ${colLetter}`

      let nonEmpty = 0
      let blanks = 0
      let formulas = 0
      let errors = 0
      let numericCount = 0
      let textCount = 0
      const nums: number[] = []
      const freq = new Map<string, number>()

      for (let r = dataStartRow; r <= bounds.end.row; r += 1) {
        const key = cellKey(r, c)
        const cell = sheet.cells[key]
        if (!cell || !cell.value.trim()) {
          blanks += 1
          continue
        }
        nonEmpty += 1
        const isFormula =
          cell.kind === 'formula' || cell.value.trim().startsWith('=')
        if (isFormula) formulas += 1
        const computed = computedCellValue(sheet, key, document.workbook).trim()
        if (isFormula && isFormulaError(computed)) {
          errors += 1
          continue
        }
        freq.set(computed, (freq.get(computed) ?? 0) + 1)
        const n = parseNumericDisplay(computed)
        if (n !== null) {
          numericCount += 1
          nums.push(n)
        } else {
          textCount += 1
        }
      }

      const inferredType =
        nonEmpty === 0
          ? 'empty'
          : numericCount > 0 && textCount === 0
            ? 'number'
            : textCount > 0 && numericCount === 0
              ? 'text'
              : 'mixed'

      let stats:
        | {
            min: number
            max: number
            sum: number
            mean: number
            median: number
          }
        | undefined
      if (nums.length > 0) {
        const sorted = [...nums].sort((a, b) => a - b)
        const sum = sorted.reduce((acc, v) => acc + v, 0)
        const mid = Math.floor(sorted.length / 2)
        const median =
          sorted.length % 2 === 1
            ? sorted[mid]
            : (sorted[mid - 1] + sorted[mid]) / 2
        stats = {
          min: roundStat(sorted[0]),
          max: roundStat(sorted[sorted.length - 1]),
          sum: roundStat(sum),
          mean: roundStat(sum / sorted.length),
          median: roundStat(median),
        }
      }

      const topValues = [...freq.entries()]
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
        .slice(0, topValuesLimit)
        .map(([value, count]) => ({ value, count }))

      columnMetas.push({
        colIndex: c,
        column: colLetter,
        header,
        numericCount,
      })

      columns.push({
        column: colLetter,
        header,
        inferredType,
        nonEmpty,
        blanks,
        distinctCount: freq.size,
        duplicateCount: Math.max(0, nonEmpty - freq.size),
        formulas,
        errors,
        ...(stats ? { stats } : {}),
        topValues,
      })
    }

    const resolveCol = (ref: string): ColumnMeta | undefined => {
      const trimmed = ref.trim()
      const lower = trimmed.toLowerCase()
      return (
        columnMetas.find(m => m.header.toLowerCase() === lower) ??
        columnMetas.find(m => m.column.toLowerCase() === lower)
      )
    }

    let groupedRollup: Record<string, unknown> | undefined
    const groupByArg = readString(args, 'groupByColumn')
    if (groupByArg) {
      const groupCol = resolveCol(groupByArg)
      if (!groupCol) {
        throw new AgentSheetsToolError(
          `Unknown groupByColumn "${groupByArg}". Available columns: ${columnMetas
            .map(m => `${m.column} (${m.header})`)
            .join(', ')}.`,
        )
      }
      const valueColArg = readString(args, 'valueColumn')
      const valueCol = valueColArg
        ? resolveCol(valueColArg)
        : columnMetas.find(
            m => m.colIndex !== groupCol.colIndex && m.numericCount > 0,
          )
      if (valueColArg && !valueCol) {
        throw new AgentSheetsToolError(
          `Unknown valueColumn "${valueColArg}". Available columns: ${columnMetas
            .map(m => `${m.column} (${m.header})`)
            .join(', ')}.`,
        )
      }

      const groups = new Map<
        string,
        {
          rowCount: number
          numericCount: number
          sum: number
          min: number
          max: number
        }
      >()
      for (let r = dataStartRow; r <= bounds.end.row; r += 1) {
        const gKey = cellKey(r, groupCol.colIndex)
        const groupLabel =
          computedCellValue(sheet, gKey, document.workbook).trim() || '(blank)'
        let agg = groups.get(groupLabel)
        if (!agg) {
          agg = {
            rowCount: 0,
            numericCount: 0,
            sum: 0,
            min: Infinity,
            max: -Infinity,
          }
          groups.set(groupLabel, agg)
        }
        agg.rowCount += 1
        if (valueCol) {
          const vKey = cellKey(r, valueCol.colIndex)
          const n = parseNumericDisplay(
            computedCellValue(sheet, vKey, document.workbook),
          )
          if (n !== null) {
            agg.numericCount += 1
            agg.sum += n
            if (n < agg.min) agg.min = n
            if (n > agg.max) agg.max = n
          }
        }
      }

      const sortedGroups = [...groups.entries()]
        .map(([group, agg]) => ({
          group,
          rowCount: agg.rowCount,
          ...(valueCol
            ? {
                numericCount: agg.numericCount,
                sum: roundStat(agg.sum),
                avg:
                  agg.numericCount > 0
                    ? roundStat(agg.sum / agg.numericCount)
                    : null,
                min: agg.numericCount > 0 ? roundStat(agg.min) : null,
                max: agg.numericCount > 0 ? roundStat(agg.max) : null,
              }
            : {}),
        }))
        .sort((a, b) => {
          const aSum = 'sum' in a && typeof a.sum === 'number' ? a.sum : a.rowCount
          const bSum = 'sum' in b && typeof b.sum === 'number' ? b.sum : b.rowCount
          return bSum - aSum || a.group.localeCompare(b.group)
        })

      groupedRollup = {
        groupBy: { column: groupCol.column, header: groupCol.header },
        ...(valueCol
          ? { valueColumn: { column: valueCol.column, header: valueCol.header } }
          : {}),
        totalGroups: sortedGroups.length,
        groups: sortedGroups.slice(0, 50),
      }
    }

    return {
      mutated: false,
      summary: `Profiled ${sheet.name}!${range} (${dataRowCount} data row(s), ${colSpan} column(s))${groupedRollup ? ` with grouped rollup by ${groupByArg}` : ''}.`,
      data: {
        version,
        sheet: sheet.name,
        sheetId: sheet.id,
        range,
        hasHeader,
        totalRows: rowSpan,
        dataRowCount,
        columnCount: colSpan,
        columns,
        ...(groupedRollup ? { groupedRollup } : {}),
      },
    }
  },
})

// ---------------------------------------------------------------------------
// 4. fn_buildStructuredSheetTable
// ---------------------------------------------------------------------------

export const fn_buildStructuredSheetTable =
  defineAppFunction<SheetsAgentToolContext>({
    name: 'fn_buildStructuredSheetTable',
    category: 'modeling',
    risk: 'write',
    requiresApproval: true,
    replaces: [
      'addSheet',
      'applyCellChanges',
      'formatRange',
      'setCellComment',
      'addChart',
      'autoFitRows',
    ],
    description:
      'Create or populate a structured spreadsheet table in one atomic, version-checked operation: optionally creates a new sheet, writes headers, data rows (values or =formulas), an optional auto-SUM summary row, header/column formatting, cell notes, an optional chart, and live row auto-fit, then returns computed readback and fresh version. Supports dryRun: true to validate before writing.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'Target sheet name/id, or new sheet name when createSheet is true.',
        },
        createSheet: {
          type: 'boolean',
          description: 'Create a new sheet before writing (default false; refuses duplicates).',
        },
        expectedVersion: {
          type: 'string',
          description: 'Exact workbook version; required when createSheet is false.',
        },
        startCell: {
          type: 'string',
          description: 'Top-left A1 anchor cell (default "A1").',
        },
        headers: {
          type: 'array',
          items: { type: 'string' },
          description: '1–50 column header labels.',
        },
        rows: {
          type: 'array',
          description: '2D row grid (strings, numbers, =formulas, booleans, null; max 400 total cells).',
          items: {
            type: 'array',
            items: { type: ['string', 'number', 'boolean', 'null'] },
          },
        },
        summaryRow: {
          type: 'object',
          description: 'Optional total row below rows: { label?, mode?: "auto-sum"|"custom", formulas? }.',
          properties: {
            label: { type: 'string' },
            mode: { type: 'string', enum: ['auto-sum', 'custom'] },
            formulas: { type: 'array', items: { type: 'string' } },
          },
        },
        styleHeader: {
          type: 'boolean',
          description: 'Style header row bold with fill and border (default true).',
        },
        columnFormats: {
          type: 'array',
          description: 'Per-column formatting by 0-based index, column letter, or header name.',
          items: {
            type: 'object',
            properties: {
              column: { type: ['integer', 'string'] },
              numberFormat: { type: 'string', enum: ['text', 'number', 'currency', 'percent', 'date'] },
              decimals: { type: 'number' },
              align: { type: 'string', enum: ['left', 'center', 'right'] },
              wrap: { type: 'string', enum: ['wrap', 'clip', 'overflow'] },
              columnWidth: { type: 'number', minimum: 56, maximum: 360 },
              bold: { type: 'boolean' },
            },
            required: ['column'],
          },
        },
        notes: {
          type: 'array',
          items: {
            type: 'object',
            properties: { cell: { type: 'string' }, comment: { type: 'string' } },
            required: ['cell', 'comment'],
          },
        },
        chart: {
          type: 'object',
          properties: {
            type: { type: 'string', enum: ['bar', 'line', 'area', 'pie', 'donut', 'scatter'] },
            range: { type: 'string' },
            summary: { type: 'string' },
          },
          required: ['type'],
        },
        autoFitRows: { type: 'boolean' },
        summary: {
          type: 'string',
          description: 'One-line summary recorded in the workbook agent log.',
        },
        dryRun: { type: 'boolean' },
      },
      required: ['sheet', 'headers', 'rows', 'summary'],
    },
    run: async (ctx, args) => {
      const document = assertDocument(ctx)
      const sheetArg = readString(args, 'sheet')
      if (!sheetArg) {
        throw new AgentSheetsToolError('Provide "sheet": the target or new sheet name.')
      }
      const summary = readString(args, 'summary')
      if (!summary) {
        throw new AgentSheetsToolError(
          'Provide "summary": a one-line description for the agent log.',
        )
      }
      const createNew = args.createSheet === true
      const dryRun = args.dryRun === true
      const startCellArg = readString(args, 'startCell') ?? 'A1'
      const startPos = !startCellArg.includes(':')
        ? parseCellKey(startCellArg)
        : null
      if (!startPos) {
        throw new AgentSheetsToolError(
          `"${startCellArg}" is not a valid single A1 startCell (e.g. "A1").`,
        )
      }

      if (
        !Array.isArray(args.headers) ||
        args.headers.length === 0 ||
        args.headers.length > 50 ||
        !args.headers.every(h => typeof h === 'string')
      ) {
        throw new AgentSheetsToolError(
          'Provide "headers": a non-empty array of 1 to 50 header strings.',
        )
      }
      const headers = args.headers as string[]

      if (
        !Array.isArray(args.rows) ||
        !args.rows.every(r => Array.isArray(r))
      ) {
        throw new AgentSheetsToolError(
          'Provide "rows": an array of row arrays.',
        )
      }
      const rows = args.rows as unknown[][]
      const colCount = headers.length

      // Validate target sheet & version
      let existingSheet: WorkbookSheet | undefined
      if (createNew) {
        if (
          document.workbook.sheets.some(
            s => s.name.toLowerCase() === sheetArg.toLowerCase(),
          )
        ) {
          throw new AgentSheetsToolError(
            `A sheet named "${sheetArg}" already exists.`,
          )
        }
      } else {
        existingSheet = sheetByRef(document, sheetArg)
        if (!existingSheet) {
          throw new AgentSheetsToolError(
            `No sheet matching "${sheetArg}". Pass createSheet: true to create it, or call getSheetsContext.`,
          )
        }
        const expectedVersion = readString(args, 'expectedVersion')
        const currentVersion = await workbookVersion(document, ctx.filePath)
        if (!expectedVersion || expectedVersion !== currentVersion) {
          throw new AgentSheetsToolError(
            'Workbook changed or expectedVersion missing. Read getSheetsContext or fn_workbookBriefing again before writing to an existing sheet.',
          )
        }
      }

      // Build cell changes
      const cellChanges: AgentCellChange[] = []
      const colHasNumericOrFormula = new Array<boolean>(colCount).fill(false)

      for (let c = 0; c < colCount; c += 1) {
        const val = headers[c]
        cellChanges.push({
          cell: cellKey(startPos.row, startPos.column + c),
          value: val,
          kind: deriveCellKind(val),
        })
      }

      for (let r = 0; r < rows.length; r += 1) {
        const row = rows[r]
        if (row.length > colCount) {
          throw new AgentSheetsToolError(
            `Row ${r + 1} has ${row.length} values, exceeding the ${colCount} declared headers.`,
          )
        }
        for (let c = 0; c < row.length; c += 1) {
          const raw = row[c]
          const strVal =
            raw === null || raw === undefined ? '' : String(raw)
          const kind = deriveCellKind(strVal)
          if (kind === 'number' || kind === 'formula') {
            colHasNumericOrFormula[c] = true
          }
          cellChanges.push({
            cell: cellKey(startPos.row + 1 + r, startPos.column + c),
            value: strVal,
            kind,
          })
        }
      }

      // Optional summary row
      let summaryRowIndex: number | null = null
      const summaryCells: Array<{ cell: string; value: string }> = []
      if (args.summaryRow !== undefined) {
        if (!args.summaryRow || typeof args.summaryRow !== 'object') {
          throw new AgentSheetsToolError('summaryRow must be an object.')
        }
        const sr = args.summaryRow as Record<string, unknown>
        const mode = readString(sr, 'mode') ?? 'auto-sum'
        if (mode !== 'auto-sum' && mode !== 'custom') {
          throw new AgentSheetsToolError(
            'summaryRow.mode must be "auto-sum" or "custom".',
          )
        }
        summaryRowIndex = startPos.row + 1 + rows.length
        const dataFirstRow = startPos.row + 1
        const dataLastRow = startPos.row + rows.length

        if (mode === 'custom') {
          if (
            !Array.isArray(sr.formulas) ||
            sr.formulas.length === 0 ||
            sr.formulas.length > colCount ||
            !sr.formulas.every(f => typeof f === 'string')
          ) {
            throw new AgentSheetsToolError(
              'summaryRow.formulas must be a non-empty array of strings when mode is "custom".',
            )
          }
          ;(sr.formulas as string[]).forEach((val, c) => {
            const k = cellKey(summaryRowIndex!, startPos.column + c)
            cellChanges.push({ cell: k, value: val, kind: deriveCellKind(val) })
            summaryCells.push({ cell: k, value: val })
          })
        } else {
          const label = readString(sr, 'label') ?? 'Total'
          const labelKey = cellKey(summaryRowIndex, startPos.column)
          cellChanges.push({
            cell: labelKey,
            value: label,
            kind: deriveCellKind(label),
          })
          summaryCells.push({ cell: labelKey, value: label })
          for (let c = 1; c < colCount; c += 1) {
            if (!colHasNumericOrFormula[c] || rows.length === 0) continue
            const colL = columnLabel(startPos.column + c)
            const formula = `=SUM(${colL}${dataFirstRow + 1}:${colL}${dataLastRow + 1})`
            const k = cellKey(summaryRowIndex, startPos.column + c)
            cellChanges.push({ cell: k, value: formula, kind: 'formula' })
            summaryCells.push({ cell: k, value: formula })
          }
        }
      }

      if (cellChanges.length > BATCH_CHANGE_CAP) {
        throw new AgentSheetsToolError(
          `Table requires ${cellChanges.length} cells, exceeding the ${BATCH_CHANGE_CAP}-cell cap per operation. Split into smaller tables or follow up with fn_applyFormulaAndFormatBatch.`,
        )
      }

      const lastRow =
        summaryRowIndex !== null
          ? summaryRowIndex
          : startPos.row + rows.length
      const lastCol = startPos.column + colCount - 1
      const tableRange = `${cellKey(startPos.row, startPos.column)}:${cellKey(lastRow, lastCol)}`
      const headerRange = `${cellKey(startPos.row, startPos.column)}:${cellKey(startPos.row, lastCol)}`
      const dataOnlyRange = `${cellKey(startPos.row, startPos.column)}:${cellKey(startPos.row + rows.length, lastCol)}`

      // Validate columnFormats
      interface ResolvedColumnFormat {
        colIndex: number
        patch: Partial<CellStyle>
        columnWidth?: number
      }
      const resolvedColFormats: ResolvedColumnFormat[] = []
      if (args.columnFormats !== undefined) {
        if (!Array.isArray(args.columnFormats)) {
          throw new AgentSheetsToolError('columnFormats must be an array.')
        }
        for (let i = 0; i < args.columnFormats.length; i += 1) {
          const item = args.columnFormats[i]
          if (!item || typeof item !== 'object') {
            throw new AgentSheetsToolError(
              `columnFormats[${i}] must be an object.`,
            )
          }
          const rec = item as Record<string, unknown>
          const colSpec = rec.column
          let colIndex = -1
          if (
            typeof colSpec === 'number' &&
            Number.isInteger(colSpec) &&
            colSpec >= 0 &&
            colSpec < colCount
          ) {
            colIndex = startPos.column + colSpec
          } else if (typeof colSpec === 'string' && colSpec.trim()) {
            const headerIdx = headers.findIndex(
              h => h.toLowerCase() === colSpec.trim().toLowerCase(),
            )
            if (headerIdx >= 0) {
              colIndex = startPos.column + headerIdx
            } else if (/^[A-Za-z]{1,3}$/.test(colSpec.trim())) {
              const absCol = columnIndex(colSpec.trim().toUpperCase())
              if (absCol >= startPos.column && absCol <= lastCol) {
                colIndex = absCol
              }
            }
          }
          if (colIndex < 0) {
            throw new AgentSheetsToolError(
              `columnFormats[${i}]: unknown column "${String(colSpec)}".`,
            )
          }
          const parsed = parseFormatSpec(rec, `columnFormats[${i}]`)
          resolvedColFormats.push({
            colIndex,
            patch: parsed.patch,
            columnWidth: parsed.columnWidth,
          })
        }
      }

      // Validate notes
      const parsedNotes: Array<{ cell: string; comment: string }> = []
      if (args.notes !== undefined) {
        if (!Array.isArray(args.notes)) {
          throw new AgentSheetsToolError('notes must be an array.')
        }
        for (let i = 0; i < args.notes.length; i += 1) {
          const item = args.notes[i]
          if (!item || typeof item !== 'object') {
            throw new AgentSheetsToolError(`notes[${i}] must be an object.`)
          }
          const rec = item as Record<string, unknown>
          const cellArg = readString(rec, 'cell') ?? ''
          const parsedCell = !cellArg.includes(':')
            ? parseCellKey(cellArg)
            : null
          if (!parsedCell) {
            throw new AgentSheetsToolError(
              `notes[${i}]: "${cellArg}" is not a valid single A1 cell.`,
            )
          }
          const comment = typeof rec.comment === 'string' ? rec.comment : ''
          if (comment.length > MAX_COMMENT_LENGTH) {
            throw new AgentSheetsToolError(
              `notes[${i}]: comment exceeds ${MAX_COMMENT_LENGTH} characters.`,
            )
          }
          parsedNotes.push({
            cell: cellKey(parsedCell.row, parsedCell.column),
            comment,
          })
        }
      }

      // Validate chart
      let chartSpec:
        | { type: WorkbookChart['type']; range: string; summary: string }
        | undefined
      if (args.chart !== undefined) {
        if (!args.chart || typeof args.chart !== 'object') {
          throw new AgentSheetsToolError('chart must be an object.')
        }
        const rec = args.chart as Record<string, unknown>
        const typeArg = readString(rec, 'type') ?? ''
        if (!CHART_TYPES.has(typeArg)) {
          throw new AgentSheetsToolError(
            `chart.type must be one of ${[...CHART_TYPES].join(', ')}.`,
          )
        }
        const rawRange = readString(rec, 'range') ?? dataOnlyRange
        const normRange = normalizeRangeToken(rawRange)
        if (!normRange) {
          throw new AgentSheetsToolError(
            `chart.range "${rawRange}" is not a valid A1 range.`,
          )
        }
        chartSpec = {
          type: typeArg as WorkbookChart['type'],
          range: normRange,
          summary:
            readString(rec, 'summary') ??
            `Added ${typeArg} chart from ${normRange}`,
        }
      }

      if (dryRun) {
        const lockedCollisions = existingSheet
          ? cellChanges
              .filter(c => isCellLocked(existingSheet!, c.cell))
              .map(c => c.cell)
          : []
        return {
          mutated: false,
          summary: `Dry run validated table on "${sheetArg}" (${tableRange}, ${cellChanges.length} cell(s), ${lockedCollisions.length} locked collision(s)).`,
          data: {
            dryRun: true,
            sheet: sheetArg,
            createSheet: createNew,
            range: tableRange,
            cellCount: cellChanges.length,
            summaryFormulas: summaryCells,
            columnFormatsCount: resolvedColFormats.length,
            notesCount: parsedNotes.length,
            chart: chartSpec ?? null,
            lockedCollisions,
          },
        }
      }

      // Commit atomically in one setDocument transaction
      const styleHeader = args.styleHeader !== false
      let stale = false
      let duplicateSheet = false
      let applied = 0
      let skippedLocked = 0
      let notesApplied = 0
      let notesSkippedLocked = 0
      let chartAdded = false
      let resizedSkippedProtected = false
      let committedDoc: PureSheetsDocument | null = null
      let targetSheetId = existingSheet?.id ?? ''
      let targetSheetName = existingSheet?.name ?? sheetArg

      ctx.setDocument(current => {
        if (
          current !== document ||
          (ctx.currentFilePath && ctx.currentFilePath() !== ctx.filePath)
        ) {
          stale = true
          return current
        }

        let next = current
        if (createNew) {
          if (
            next.workbook.sheets.some(
              s => s.name.toLowerCase() === sheetArg.toLowerCase(),
            )
          ) {
            duplicateSheet = true
            return current
          }
          next = addSheet(next, sheetArg)
          const created = next.workbook.sheets[next.workbook.sheets.length - 1]
          targetSheetId = created.id
          targetSheetName = created.name
          next = appendAgentLog(next, {
            agentName: 'PureSheets Assistant',
            tool: 'fn_buildStructuredSheetTable',
            summary: `Created sheet ${created.name}`,
            sheetId: created.id,
            sheetName: created.name,
          })
        }

        const writeResult = applyAgentCellChangesDirect(
          next,
          targetSheetId,
          tableRange,
          summary,
          cellChanges,
        )
        next = writeResult.document
        applied = writeResult.applied
        skippedLocked = writeResult.skippedLocked

        let sheetObj = next.workbook.sheets.find(s => s.id === targetSheetId)!
        const isProtected = Boolean(sheetObj.protection?.lockedRanges?.length)

        if (styleHeader) {
          sheetObj = setRangeStyle(
            sheetObj,
            headerRange,
            { bold: true, fillColor: '#f1f5f9', border: true },
            FORMAT_CELL_CAP,
          )
        }
        if (summaryRowIndex !== null) {
          const sumRange = `${cellKey(summaryRowIndex, startPos.column)}:${cellKey(summaryRowIndex, lastCol)}`
          sheetObj = setRangeStyle(
            sheetObj,
            sumRange,
            { bold: true, border: true },
            FORMAT_CELL_CAP,
          )
        }

        const dataBodyStartRow = startPos.row + 1
        for (const fmt of resolvedColFormats) {
          if (
            Object.keys(fmt.patch).length > 0 &&
            lastRow >= dataBodyStartRow
          ) {
            const colRange = `${cellKey(dataBodyStartRow, fmt.colIndex)}:${cellKey(lastRow, fmt.colIndex)}`
            sheetObj = setRangeStyle(
              sheetObj,
              colRange,
              fmt.patch,
              FORMAT_CELL_CAP,
            )
          }
          if (fmt.columnWidth !== undefined) {
            if (isProtected) {
              resizedSkippedProtected = true
            } else {
              sheetObj = setColumnWidth(sheetObj, fmt.colIndex, fmt.columnWidth)
            }
          }
        }

        for (const note of parsedNotes) {
          if (isCellLocked(sheetObj, note.cell)) {
            notesSkippedLocked += 1
            continue
          }
          sheetObj = setCellComment(sheetObj, note.cell, note.comment)
          notesApplied += 1
        }

        next = {
          ...next,
          workbook: {
            ...next.workbook,
            sheets: next.workbook.sheets.map(s =>
              s.id === targetSheetId ? sheetObj : s,
            ),
          },
        }

        if (chartSpec) {
          const withChart = addChartToSheet(
            next,
            targetSheetId,
            chartSpec.range,
            chartSpec.type,
            chartSpec.summary,
          )
          chartAdded = withChart !== next
          next = withChart
        }

        committedDoc = next
        return next
      })

      if (stale) {
        throw new AgentSheetsToolError(
          'Workbook changed during the write. Read it again; nothing applied.',
        )
      }
      if (duplicateSheet) {
        throw new AgentSheetsToolError(
          `A sheet named "${sheetArg}" already exists.`,
        )
      }
      if (!committedDoc) {
        throw new AgentSheetsToolError('Could not commit structured table.')
      }

      const finalDoc = committedDoc as PureSheetsDocument
      const finalSheet = finalDoc.workbook.sheets.find(
        s => s.id === targetSheetId,
      )!

      // Optional live row auto-fit
      let rowsAutoFitted = false
      if (args.autoFitRows === true && !finalSheet.protection?.lockedRanges?.length) {
        const rowIndices: number[] = []
        const capEnd = Math.min(lastRow, startPos.row + 99)
        for (let r = startPos.row; r <= capEnd; r += 1) rowIndices.push(r)
        rowsAutoFitted = Boolean(
          await ctx.autoFitRows?.(targetSheetId, rowIndices),
        )
      }

      // Compute readback for summary cells and any formula errors
      const summaryReadback = summaryCells.map(({ cell, value }) => ({
        cell,
        value,
        computed: computedCellValue(finalSheet, cell, finalDoc.workbook),
      }))
      const formulaErrors = cellChanges
        .filter(c => c.kind === 'formula')
        .map(c => ({
          cell: c.cell,
          formula: c.value,
          computed: computedCellValue(finalSheet, c.cell, finalDoc.workbook),
        }))
        .filter(item => isFormulaError(item.computed))

      const newVersion = await workbookVersion(finalDoc, ctx.filePath)

      return {
        mutated: true,
        summary: `Built table on ${targetSheetName}!${tableRange} (${applied} cell(s) written${summaryReadback.length ? `, ${summaryReadback.length} summary cell(s)` : ''}${chartAdded ? ', chart added' : ''}).`,
        data: {
          dryRun: false,
          sheet: targetSheetName,
          sheetId: targetSheetId,
          createdSheet: createNew,
          range: tableRange,
          version: newVersion,
          applied,
          skippedLocked,
          notesApplied,
          notesSkippedLocked,
          chartAdded,
          rowsAutoFitted,
          ...(resizedSkippedProtected
            ? {
                warning:
                  'Column width changes were skipped because the target sheet has locked ranges.',
              }
            : {}),
          summaryReadback,
          formulaErrors,
        },
      }
    },
  })

// ---------------------------------------------------------------------------
// 5. fn_applyFormulaAndFormatBatch
// ---------------------------------------------------------------------------

export const fn_applyFormulaAndFormatBatch =
  defineAppFunction<SheetsAgentToolContext>({
    name: 'fn_applyFormulaAndFormatBatch',
    category: 'modeling',
    risk: 'write',
    requiresApproval: true,
    replaces: [
      'applyCellChanges',
      'formatRange',
      'setCellComment',
      'readRange',
    ],
    description:
      'Apply a version-checked batch of cell edits (values/formulas up to 400 cells), multiple range formatting specs, and cell notes on an existing sheet in one atomic transaction, then return computed formula readback, error checks, and the fresh workbook version. Supports dryRun: true for preview.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: { type: 'string', description: 'Explicit target sheet name or id.' },
        expectedVersion: { type: 'string', description: 'Exact workbook version token.' },
        summary: { type: 'string', description: 'One-line summary for the workbook agent log.' },
        range: { type: 'string', description: 'Optional A1 bounding range for cellChanges.' },
        cellChanges: {
          type: 'array',
          description: 'Up to 400 cell edits ({ cell, value, kind? }).',
          items: {
            type: 'object',
            properties: {
              cell: { type: 'string' },
              value: { type: 'string' },
              kind: { type: 'string', enum: ['text', 'number', 'formula', 'blank'] },
            },
            required: ['cell', 'value'],
          },
        },
        formats: {
          type: 'array',
          description: 'Up to 25 range formatting specs (bold, italic, underline, border, fontSize, textColor, fillColor, align, numberFormat, decimals, wrap, columnWidth, rowHeight).',
          items: {
            type: 'object',
            properties: {
              range: { type: 'string' },
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              underline: { type: 'boolean' },
              border: { type: 'boolean' },
              fontSize: { type: 'number' },
              textColor: { type: 'string' },
              fillColor: { type: 'string' },
              align: { type: 'string', enum: ['left', 'center', 'right'] },
              numberFormat: { type: 'string', enum: ['text', 'number', 'currency', 'percent', 'date'] },
              decimals: { type: 'number' },
              wrap: { type: 'string', enum: ['wrap', 'clip', 'overflow'] },
              columnWidth: { type: 'number', minimum: 56, maximum: 360 },
              rowHeight: { type: 'number', minimum: 24, maximum: 160 },
            },
            required: ['range'],
          },
        },
        comments: {
          type: 'array',
          description: 'Up to 50 cell notes ({ cell, comment }); empty comment clears.',
          items: {
            type: 'object',
            properties: { cell: { type: 'string' }, comment: { type: 'string' } },
            required: ['cell', 'comment'],
          },
        },
        dryRun: { type: 'boolean' },
      },
      required: ['sheet', 'expectedVersion', 'summary'],
    },
    run: async (ctx, args) => {
      const document = assertDocument(ctx)
      const sheetRef = readString(args, 'sheet')
      if (!sheetRef) {
        throw new AgentSheetsToolError(
          'Provide explicit target "sheet" name or id.',
        )
      }
      const sheet = sheetByRef(document, sheetRef)
      if (!sheet) {
        throw new AgentSheetsToolError(
          `No sheet matching "${sheetRef}". Call getSheetsContext to list sheets.`,
        )
      }
      const summary = readString(args, 'summary')
      if (!summary) {
        throw new AgentSheetsToolError(
          'Provide "summary": one sentence describing this batch for the agent log.',
        )
      }
      const expectedVersion = readString(args, 'expectedVersion')
      const currentVersion = await workbookVersion(document, ctx.filePath)
      if (!expectedVersion || expectedVersion !== currentVersion) {
        throw new AgentSheetsToolError(
          'Workbook changed or expectedVersion missing. Read getSheetsContext/readRange/fn_workbookBriefing again before writing.',
        )
      }

      const rawChanges = Array.isArray(args.cellChanges)
        ? args.cellChanges
        : []
      const rawFormats = Array.isArray(args.formats) ? args.formats : []
      const rawComments = Array.isArray(args.comments) ? args.comments : []

      if (
        rawChanges.length === 0 &&
        rawFormats.length === 0 &&
        rawComments.length === 0
      ) {
        throw new AgentSheetsToolError(
          'Pass at least one of cellChanges, formats, or comments.',
        )
      }
      if (rawChanges.length > BATCH_CHANGE_CAP) {
        throw new AgentSheetsToolError(
          `At most ${BATCH_CHANGE_CAP} cellChanges per batch; split larger updates.`,
        )
      }
      if (rawFormats.length > 25) {
        throw new AgentSheetsToolError('At most 25 format entries per batch.')
      }
      if (rawComments.length > 50) {
        throw new AgentSheetsToolError('At most 50 comment entries per batch.')
      }

      // Validate cellChanges & bounding range
      const parsedChanges: AgentCellChange[] = []
      let minRow = Infinity
      let maxRow = -Infinity
      let minCol = Infinity
      let maxCol = -Infinity
      for (let i = 0; i < rawChanges.length; i += 1) {
        const item = rawChanges[i]
        if (!item || typeof item !== 'object') {
          throw new AgentSheetsToolError(
            `cellChanges[${i}] must be an object.`,
          )
        }
        const rec = item as Record<string, unknown>
        const cellArg = readString(rec, 'cell') ?? ''
        const pos = !cellArg.includes(':') ? parseCellKey(cellArg) : null
        if (!pos) {
          throw new AgentSheetsToolError(
            `cellChanges[${i}]: "${cellArg}" is not a valid A1 cell key.`,
          )
        }
        if (typeof rec.value !== 'string') {
          throw new AgentSheetsToolError(
            `cellChanges[${i}]: value must be a string.`,
          )
        }
        const canonicalCell = cellKey(pos.row, pos.column)
        parsedChanges.push({
          cell: canonicalCell,
          value: rec.value,
          kind: deriveCellKind(rec.value),
        })
        if (pos.row < minRow) minRow = pos.row
        if (pos.row > maxRow) maxRow = pos.row
        if (pos.column < minCol) minCol = pos.column
        if (pos.column > maxCol) maxCol = pos.column
      }

      const explicitRangeArg = readString(args, 'range')
      let targetRange: string | null = null
      if (explicitRangeArg) {
        const norm = normalizeRangeToken(explicitRangeArg)
        const bounds = norm ? parseRangeBounds(norm) : null
        if (!norm || !bounds) {
          throw new AgentSheetsToolError(
            `"${explicitRangeArg}" is not a valid A1 target range.`,
          )
        }
        for (const change of parsedChanges) {
          const pos = parseCellKey(change.cell)!
          if (
            pos.row < bounds.start.row ||
            pos.row > bounds.end.row ||
            pos.column < bounds.start.column ||
            pos.column > bounds.end.column
          ) {
            throw new AgentSheetsToolError(
              `Cell ${change.cell} lies outside declared target range ${norm}. Nothing applied.`,
            )
          }
        }
        targetRange = norm
      } else if (parsedChanges.length > 0) {
        targetRange = `${cellKey(minRow, minCol)}:${cellKey(maxRow, maxCol)}`
      }

      // Validate formats
      interface ParsedFormatOp {
        range: string
        bounds: NonNullable<ReturnType<typeof parseRangeBounds>>
        patch: Partial<CellStyle>
        columnWidth?: number
        rowHeight?: number
      }
      const parsedFormats: ParsedFormatOp[] = []
      const isProtected = Boolean(sheet.protection?.lockedRanges?.length)
      for (let i = 0; i < rawFormats.length; i += 1) {
        const item = rawFormats[i]
        if (!item || typeof item !== 'object') {
          throw new AgentSheetsToolError(`formats[${i}] must be an object.`)
        }
        const rec = item as Record<string, unknown>
        const rArg = readString(rec, 'range') ?? ''
        const norm = rArg ? normalizeRangeToken(rArg) : null
        const bounds = norm ? parseRangeBounds(norm) : null
        if (!norm || !bounds) {
          throw new AgentSheetsToolError(
            `formats[${i}]: "${rArg}" is not a valid A1 range.`,
          )
        }
        const span =
          (bounds.end.row - bounds.start.row + 1) *
          (bounds.end.column - bounds.start.column + 1)
        if (!Number.isSafeInteger(span) || span > FORMAT_CELL_CAP) {
          throw new AgentSheetsToolError(
            `formats[${i}]: ${norm} spans ${span} cells (cap ${FORMAT_CELL_CAP}).`,
          )
        }
        const parsed = parseFormatSpec(rec, `formats[${i}]`)
        const resizing =
          parsed.columnWidth !== undefined || parsed.rowHeight !== undefined
        if (!Object.keys(parsed.patch).length && !resizing) {
          throw new AgentSheetsToolError(
            `formats[${i}]: specify at least one style or size property.`,
          )
        }
        if (resizing && isProtected) {
          throw new AgentSheetsToolError(
            `formats[${i}]: columnWidth/rowHeight changes are refused on protected sheets.`,
          )
        }
        parsedFormats.push({
          range: norm,
          bounds,
          patch: parsed.patch,
          columnWidth: parsed.columnWidth,
          rowHeight: parsed.rowHeight,
        })
      }

      // Validate comments
      const parsedComments: Array<{ cell: string; comment: string }> = []
      for (let i = 0; i < rawComments.length; i += 1) {
        const item = rawComments[i]
        if (!item || typeof item !== 'object') {
          throw new AgentSheetsToolError(`comments[${i}] must be an object.`)
        }
        const rec = item as Record<string, unknown>
        const cellArg = readString(rec, 'cell') ?? ''
        const pos = !cellArg.includes(':') ? parseCellKey(cellArg) : null
        if (!pos) {
          throw new AgentSheetsToolError(
            `comments[${i}]: "${cellArg}" is not a valid A1 cell.`,
          )
        }
        const comment = typeof rec.comment === 'string' ? rec.comment : ''
        if (comment.length > MAX_COMMENT_LENGTH) {
          throw new AgentSheetsToolError(
            `comments[${i}]: comment exceeds ${MAX_COMMENT_LENGTH} characters.`,
          )
        }
        parsedComments.push({
          cell: cellKey(pos.row, pos.column),
          comment,
        })
      }

      const dryRun = args.dryRun === true
      if (dryRun) {
        const lockedCellSkips = parsedChanges
          .filter(c => isCellLocked(sheet, c.cell))
          .map(c => c.cell)
        const lockedCommentSkips = parsedComments
          .filter(c => isCellLocked(sheet, c.cell))
          .map(c => c.cell)
        return {
          mutated: false,
          summary: `Dry run validated batch on ${sheet.name}: ${parsedChanges.length} cell edit(s), ${parsedFormats.length} format range(s), ${parsedComments.length} note(s).`,
          data: {
            dryRun: true,
            sheet: sheet.name,
            range: targetRange,
            plannedCellChanges: parsedChanges.length,
            plannedFormats: parsedFormats.length,
            plannedComments: parsedComments.length,
            lockedCellSkips,
            lockedCommentSkips,
          },
        }
      }

      let stale = false
      let appliedCells = 0
      let skippedLockedCells = 0
      let appliedFormats = 0
      let appliedComments = 0
      let skippedLockedComments = 0
      let committedDoc: PureSheetsDocument | null = null

      ctx.setDocument(current => {
        if (
          current !== document ||
          (ctx.currentFilePath && ctx.currentFilePath() !== ctx.filePath)
        ) {
          stale = true
          return current
        }
        let next = current
        let loggedViaDirectCells = false

        if (parsedChanges.length > 0 && targetRange) {
          const res = applyAgentCellChangesDirect(
            next,
            sheet.id,
            targetRange,
            summary,
            parsedChanges,
          )
          next = res.document
          appliedCells = res.applied
          skippedLockedCells = res.skippedLocked
          loggedViaDirectCells = res.applied > 0
        }

        const foundSheet = next.workbook.sheets.find(s => s.id === sheet.id)
        if (!foundSheet) return current
        let targetSheet: WorkbookSheet = foundSheet

        for (const fmt of parsedFormats) {
          let updated: WorkbookSheet = targetSheet
          if (Object.keys(fmt.patch).length > 0) {
            updated = setRangeStyle(
              updated,
              fmt.range,
              fmt.patch,
              FORMAT_CELL_CAP,
            )
          }
          if (fmt.columnWidth !== undefined) {
            for (
              let col = fmt.bounds.start.column;
              col <= fmt.bounds.end.column;
              col += 1
            ) {
              updated = setColumnWidth(updated, col, fmt.columnWidth)
            }
          }
          if (fmt.rowHeight !== undefined) {
            for (
              let row = fmt.bounds.start.row;
              row <= fmt.bounds.end.row;
              row += 1
            ) {
              updated = setRowHeight(updated, row, fmt.rowHeight)
            }
          }
          if (updated !== targetSheet) {
            targetSheet = updated
            appliedFormats += 1
          }
        }

        for (const note of parsedComments) {
          if (isCellLocked(targetSheet, note.cell)) {
            skippedLockedComments += 1
            continue
          }
          targetSheet = setCellComment(targetSheet, note.cell, note.comment)
          appliedComments += 1
        }

        if (appliedFormats > 0 || appliedComments > 0) {
          next = {
            ...next,
            workbook: {
              ...next.workbook,
              sheets: next.workbook.sheets.map(s =>
                s.id === sheet.id ? targetSheet! : s,
              ),
            },
          }
          if (!loggedViaDirectCells) {
            next = appendAgentLog(next, {
              agentName: 'PureSheets Assistant',
              tool: 'fn_applyFormulaAndFormatBatch',
              summary,
              sheetId: sheet.id,
              sheetName: sheet.name,
              range: targetRange ?? parsedFormats[0]?.range ?? parsedComments[0]?.cell,
            })
          }
        }

        committedDoc = next
        return next
      })

      if (stale) {
        throw new AgentSheetsToolError(
          'Workbook changed during the write. Read it again; nothing applied.',
        )
      }
      if (!committedDoc) {
        throw new AgentSheetsToolError('Batch could not be committed.')
      }

      const finalDoc = committedDoc as PureSheetsDocument
      const finalSheet = finalDoc.workbook.sheets.find(s => s.id === sheet.id)!

      const formulaReadback = parsedChanges
        .filter(c => c.kind === 'formula')
        .map(c => {
          const computed = computedCellValue(
            finalSheet,
            c.cell,
            finalDoc.workbook,
          )
          return {
            cell: c.cell,
            formula: c.value,
            computed,
            isError: isFormulaError(computed),
          }
        })
      const formulaErrorCount = formulaReadback.filter(r => r.isError).length
      const didMutate =
        appliedCells > 0 || appliedFormats > 0 || appliedComments > 0
      const newVersion = await workbookVersion(finalDoc, ctx.filePath)

      return {
        mutated: didMutate,
        summary: `Applied batch on ${sheet.name}: ${appliedCells} cell(s), ${appliedFormats} format range(s), ${appliedComments} note(s)${formulaErrorCount ? ` (${formulaErrorCount} formula error(s) detected)` : ''}.`,
        data: {
          dryRun: false,
          sheet: sheet.name,
          sheetId: sheet.id,
          range: targetRange,
          version: newVersion,
          appliedCells,
          skippedLockedCells,
          appliedFormats,
          appliedComments,
          skippedLockedComments,
          formulaReadback,
          formulaErrorCount,
        },
      }
    },
  })

export const PURESHEETS_FUNCTIONS: readonly AppFunctionDefinition<SheetsAgentToolContext>[] =
  [
    fn_workbookBriefing,
    fn_auditWorkbookFormulasAndData,
    fn_profileSheetData,
    fn_buildStructuredSheetTable,
    fn_applyFormulaAndFormatBatch,
  ] as const
