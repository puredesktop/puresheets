import type { PureSheetsDocument } from '../types'

/** Keep in sync with `plugin.json` -> `app.agents.tools[].name`. */
export const PURESHEETS_AGENT_TOOL_NAMES = [
  'getSheetsContext',
  'createWorkbook',
  'autoFitRows',
  'completeWorkbook',
  'readRange',
  'findInWorkbook',
  'applyCellChanges',
  'getExchangeRate',
  'addChart',
  'listSheetChanges',
  'addSheet',
  'renameSheet',
  'deleteSheet',
  'formatRange',
  'setCellComment',
] as const

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
  createWorkbook?: (title: string) => Promise<{ title: string; filePath: null }>
  autoFitRows?: (sheetId: string, rows: number[]) => Promise<boolean>
  completeWorkbook?: () => Promise<{ taskOutcome: string; artifactPaths: string[]; observations: string[] }>
}
