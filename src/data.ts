import { csvFromSheet } from './lib/csv'
import { readWorkbookFromPath } from './lib/workbookFiles'
import { computedCellValue, parseCellKey } from './lib/workbookModel'

export interface SheetDataset {
  id: string
  name: string
  csv: string
  error?: string
}

/** Read-only, saved-data interchange. Consumers never mount the Sheets editor. */
export async function readWorkbookDatasets(path: string, readText: (path: string) => Promise<string>) {
  const { document, recoveredFromBackup } = await readWorkbookFromPath(path, readText)
  const sheets: SheetDataset[] = document.workbook.sheets.map(original => {
    try {
      const sheet = { ...original, cells: { ...original.cells } }
      let rows = 0
      let columns = 0
      for (const [key, cell] of Object.entries(sheet.cells)) {
        if (!cell.value) { delete sheet.cells[key]; continue }
        const position = parseCellKey(key)
        if (!position) continue
        rows = Math.max(rows, position.row + 1)
        columns = Math.max(columns, position.column + 1)
        if (cell.kind === 'formula') {
          const value = computedCellValue(original, key)
          if (/^#[A-Z0-9/?!]+$/i.test(value.trim())) {
            throw new Error(`Formula at ${key} could not be exported. Check it in PureSheets or import a values-only dataset.`)
          }
          sheet.cells[key] = { ...cell, kind: 'text', value }
        }
      }
      if (rows * columns > 250_000) throw new Error('Sheet is too large to import. Export a smaller data range first.')
      return { id: sheet.id, name: sheet.name, csv: csvFromSheet(sheet) }
    } catch (error) {
      return { id: original.id, name: original.name, csv: '', error: error instanceof Error ? error.message : String(error) }
    }
  })
  return { title: document.metadata.title, sheets, recoveredFromBackup }
}
