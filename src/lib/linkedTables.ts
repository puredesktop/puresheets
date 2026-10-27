import {
  snapshotFromTableSource,
  type LinkedTableReference,
  type LinkedTableSource,
} from '@purescience/platform-editor/linkedTable'
import { displayCellValue } from './workbookModel'
import type { PureSheetsDocument } from '../types'

/** Derived publication data, rebuilt at save time. The workbook remains the
 * only editable source; writing apps never need to run its calculation engine. */
export function workbookLinkedTables(
  document: PureSheetsDocument,
): LinkedTableSource {
  return {
    version: 1,
    tables: document.workbook.sheets
      .filter(sheet => !sheet.agentDraft)
      .map(sheet => ({
        id: sheet.id,
        title: `${document.metadata.title} · ${sheet.name}`,
        cells: Object.fromEntries(
          Object.entries(sheet.cells).map(([key, cell]) => [
            key,
            {
              text: displayCellValue(sheet, key),
              bold: cell.style?.bold,
              italic: cell.style?.italic,
              color: cell.style?.textColor,
              background: cell.style?.fillColor,
              align: cell.style?.align,
            },
          ]),
        ),
      })),
  }
}
export function workbookLinkedSnapshot(
  document: PureSheetsDocument,
  reference: LinkedTableReference,
) {
  return snapshotFromTableSource(workbookLinkedTables(document), reference)
}
