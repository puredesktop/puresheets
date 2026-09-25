import type { PureSheetsUniverSnapshot } from './univerAdapter'

/** The model owns explicit sheet changes; the renderer owns in-sheet position. */
export function requestedSheetId(snapshot: PureSheetsUniverSnapshot): string | null {
  const custom = snapshot.custom as { puresheets?: { activeSheetId?: unknown } } | undefined
  const id = custom?.puresheets?.activeSheetId
  return typeof id === 'string' && snapshot.sheets[id] ? id : null
}
