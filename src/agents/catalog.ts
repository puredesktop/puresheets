import type { PureSheetsDocument } from '../types'

/**
 * Exposed PureSheets operations (must stay in sync 1:1 with `plugin.json` -> `app.agents.tools[].name`).
 *
 * 3-tier exposure (built on `@purescience/platform-ui/bridge/appFunctions`):
 *   1. 5 composite `function` workflows (`fn_*`) — replace 3–12 atomic round-trips each
 *   2. 4 core atomic tools (~24% of 17) — indispensable direct read/write/save verbs
 *   3. 3 universal escape-hatch meta-tools (`discoverAppTools`, `callExtendedReadTool`, `callExtendedTool`)
 *      providing full on-demand access to the remaining 13 specialist atomic tools.
 */
export const PURESHEETS_FUNCTION_NAMES = [
  'fn_workbookBriefing',
  'fn_auditWorkbookFormulasAndData',
  'fn_profileSheetData',
  'fn_buildStructuredSheetTable',
  'fn_applyFormulaAndFormatBatch',
] as const

export const PURESHEETS_CORE_TOOL_NAMES = [
  'getSheetsContext',
  'readRange',
  'applyCellChanges',
  'completeWorkbook',
] as const

export const PURESHEETS_META_TOOL_NAMES = [
  'discoverAppTools',
  'callExtendedReadTool',
  'callExtendedTool',
] as const

export const EXPOSED_TOOL_NAMES = [
  ...PURESHEETS_FUNCTION_NAMES,
  ...PURESHEETS_CORE_TOOL_NAMES,
  ...PURESHEETS_META_TOOL_NAMES,
] as const

/** Keep in sync with `plugin.json` -> `app.agents.tools[].name`. */
export const PURESHEETS_AGENT_TOOL_NAMES = EXPOSED_TOOL_NAMES

export const PURESHEETS_EXTENDED_TOOL_NAMES = [
  'openWorkbook',
  'createWorkbook',
  'saveWorkbookAsXlsx',
  'listSheetChanges',
  'findInWorkbook',
  'addSheet',
  'renameSheet',
  'deleteSheet',
  'formatRange',
  'autoFitRows',
  'setCellComment',
  'addChart',
  'getExchangeRate',
] as const

export const PURESHEETS_ALL_ATOMIC_TOOL_NAMES = [
  ...PURESHEETS_CORE_TOOL_NAMES,
  ...PURESHEETS_EXTENDED_TOOL_NAMES,
] as const

export type PureSheetsAtomicToolName =
  (typeof PURESHEETS_ALL_ATOMIC_TOOL_NAMES)[number]

export const PURESHEETS_AGENT_LOG_LABEL = 'sheets'

export class AgentSheetsToolError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'AgentSheetsToolError'
  }
}

/**
 * Reads and writes are direct. Every agent write lands on the sheet
 * immediately and is recorded in the workbook's agent log with
 * before/after — the log, not a review queue, is the accountability
 * mechanism; approval gates are the shell's permissions. Locked cells
 * are never mutated.
 */
export interface SheetsAgentToolContext {
  document: PureSheetsDocument | null
  setDocument: (
    updater: (current: PureSheetsDocument) => PureSheetsDocument,
  ) => void
  currentFilePath?: () => string
  filePath: string
  selectedRange: string
  openWorkbook?: (path: string) => Promise<{ title: string; sourcePath: string; filePath: string | null; warnings: string[] }>
  saveAsXlsx?: (path: string, overwrite: boolean) => Promise<{ path: string; sheets: string[]; warnings: string[] }>
  createWorkbook?: (title: string) => Promise<{ title: string; filePath: null }>
  autoFitRows?: (sheetId: string, rows: number[]) => Promise<boolean>
  completeWorkbook?: () => Promise<{ taskOutcome: string; artifactPaths: string[]; observations: string[] }>
}
