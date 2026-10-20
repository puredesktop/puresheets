/**
 * Tags for the shared agent activity layer: the refs an agent's tool call names, so the layer can mark what on
 * screen the call touches. Attributes only — no behaviour.
 *
 * The sheet tools take `sheet` (a sheet id or name), addSheet/renameSheet take `name` (the sheet's name), and reads
 * answer with sheet `id`s and the workbook's `filePath`. openWorkbook takes `path`. Cells live on Univer's canvas, so
 * they cannot be tagged; range- and cell-addressed calls land on their sheet's tab (and the grid, when it is the
 * active sheet) through their `sheet` argument.
 *
 * Refs are space-separated, so a name or path containing whitespace can never match and is left out.
 */

export interface AgentRefAttributes {
  'data-agent-ref': string
  'data-agent-label': string
}

function usable(value: string | null | undefined): value is string {
  return Boolean(value && value.trim() && !/\s/.test(value))
}

function sheetRefs(sheet: { id: string; name: string }): string[] {
  const refs: string[] = []
  if (usable(sheet.id)) refs.push(sheet.id, `sheet:${sheet.id}`)
  if (usable(sheet.name)) refs.push(`sheet:${sheet.name}`, `name:${sheet.name}`)
  return refs
}

/** A sheet tab: its id (plain, as reads answer it, and as `sheet:`), and its name as `sheet:` and `name:`. */
export function sheetAgentRefs(sheet: { id: string; name: string }): AgentRefAttributes {
  return { 'data-agent-ref': sheetRefs(sheet).join(' '), 'data-agent-label': sheet.name }
}

/** The workbook's grid: the bound file path (plain, `path:`, `filePath:`) plus the sheet it is showing. */
export function workbookAgentRefs(
  filePath: string,
  title: string,
  activeSheet: { id: string; name: string } | undefined,
): AgentRefAttributes {
  const refs: string[] = []
  if (usable(filePath)) refs.push(filePath, `path:${filePath}`, `filePath:${filePath}`)
  if (activeSheet) refs.push(...sheetRefs(activeSheet))
  return {
    'data-agent-ref': refs.join(' '),
    'data-agent-label': activeSheet ? `${title || 'Workbook'} — ${activeSheet.name}` : title || 'Workbook',
  }
}
