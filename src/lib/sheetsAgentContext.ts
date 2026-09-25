import type { PureSheetsDocument, WorkbookSheet } from '../types'
import { cellKey, parseCellKey } from './workbookModel'

export function usedRangeForSheet(sheet: WorkbookSheet): string | null {
  const positions = Object.entries(sheet.cells)
    .filter(([, cell]) => cell.value.trim())
    .map(([key]) => parseCellKey(key))
    .filter((position): position is NonNullable<typeof position> => Boolean(position))
  if (!positions.length) return null
  const minRow = Math.min(...positions.map(position => position.row))
  const maxRow = Math.max(...positions.map(position => position.row))
  const minColumn = Math.min(...positions.map(position => position.column))
  const maxColumn = Math.max(...positions.map(position => position.column))
  return `${cellKey(minRow, minColumn)}:${cellKey(maxRow, maxColumn)}`
}


/** Includes every source sheet and file binding, so cross-sheet calculations reject stale data. */
export async function workbookVersion(document: PureSheetsDocument, filePath: string): Promise<string> {
  const bytes = new TextEncoder().encode(JSON.stringify({ filePath, metadata: document.metadata, workbook: document.workbook }))
  const digest = await crypto.subtle.digest('SHA-256', bytes)
  return Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('')
}
