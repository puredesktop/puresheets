import {
  isSheetsHtmlPath,
  loadSheetsDocument,
  loadSheetsHtmlDocument,
  serializeSheetsDocument,
  serializeSheetsHtmlDocument,
} from './sheetsDocument'
import type { PureSheetsDocument } from '../types'

/**
 * Workbook import/export flows: how a `.sheets` package, a legacy flat
 * `.sheets` file, or a `.sheets.html` file on disk becomes a
 * `PureSheetsDocument` and back. PureDesktop owns this persistence layer;
 * the Univer surface only ever sees the in-memory document/snapshot.
 */

export interface LoadedWorkbookFile {
  document: PureSheetsDocument
  /** True when the path names a `.sheets` package folder (workbook.json). */
  isPackage: boolean
  /** True when the main file was unreadable and the `.bak` sibling was used. */
  recoveredFromBackup: boolean
}

/**
 * Read a workbook from disk. `.sheets` names either a package folder or a
 * legacy flat file — the package content file is probed first. Flat files
 * fall back to the rolling `.bak` sibling written by pre-lifecycle saves when
 * the main file is unreadable (the lifecycle itself writes atomically, so new
 * saves cannot tear — recovery only matters for legacy files).
 */
export async function readWorkbookFromPath(
  resourcePath: string,
  readTextFile: (path: string) => Promise<string>,
): Promise<LoadedWorkbookFile & { path: string }> {
  const clean = resourcePath.replace(/\/+$/, '')
  const readFlatWithRecovery = async (
    parse: (raw: string) => PureSheetsDocument,
    path: string,
  ): Promise<{ doc: PureSheetsDocument; recovered: boolean }> => {
    const raw = await readTextFile(path)
    try {
      return { doc: parse(raw), recovered: false }
    } catch (parseError) {
      try {
        return {
          doc: parse(await readTextFile(`${path}.bak`)),
          recovered: true,
        }
      } catch {
        throw parseError
      }
    }
  }
  if (isSheetsHtmlPath(clean)) {
    const result = await readFlatWithRecovery(loadSheetsHtmlDocument, clean)
    return {
      document: result.doc,
      isPackage: false,
      recoveredFromBackup: result.recovered,
      path: clean,
    }
  }
  if (clean.endsWith('.sheets')) {
    try {
      const document = loadSheetsDocument(
        await readTextFile(`${clean}/workbook.json`),
      )
      return {
        document,
        isPackage: true,
        recoveredFromBackup: false,
        path: clean,
      }
    } catch {
      const result = await readFlatWithRecovery(loadSheetsDocument, clean)
      return {
        document: result.doc,
        isPackage: false,
        recoveredFromBackup: result.recovered,
        path: clean,
      }
    }
  }
  const result = await readFlatWithRecovery(loadSheetsDocument, clean)
  return {
    document: result.doc,
    isPackage: false,
    recoveredFromBackup: result.recovered,
    path: clean,
  }
}

export interface WorkbookFilePayload {
  /** File name inside a package, or null for "write the bound path itself". */
  name: string | null
  content: string
}

/**
 * Serialize a workbook for the unified document lifecycle. New workbooks are
 * `.sheets` package folders (manifest + workbook.json); legacy flat `.sheets`
 * and `.sheets.html` files opened from disk are written back in place in
 * their original single-file format.
 */
export function workbookFilesForSave(
  document: PureSheetsDocument,
  options: { boundPath: string | null; isPackage: boolean },
): WorkbookFilePayload[] {
  if (options.isPackage) {
    return [
      {
        name: 'manifest.json',
        content: `${JSON.stringify(
          {
            schemaVersion: 1,
            kind: 'purescience.sheets.document',
            packageSuffix: '.sheets',
            title: document.metadata.title,
            savedAt: new Date().toISOString(),
          },
          null,
          2,
        )}\n`,
      },
      { name: 'workbook.json', content: serializeSheetsDocument(document) },
    ]
  }
  return [
    {
      name: null,
      content:
        options.boundPath && isSheetsHtmlPath(options.boundPath)
          ? serializeSheetsHtmlDocument(document)
          : serializeSheetsDocument(document),
    },
  ]
}
