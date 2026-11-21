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
    try {
      return { doc: parse(await readTextFile(path)), recovered: false }
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
  if (/\.sheets$/i.test(clean)) {
    let raw: string
    try {
      raw = await readTextFile(`${clean}/workbook.json`)
    } catch (error) {
      // Only an absent package member (or a flat file where a directory was
      // probed) justifies trying the legacy flat format. Preserve real I/O errors.
      if (
        !/ENOENT|ENOTDIR|not a directory|no such file|not found|does not exist/i.test(
          String(error),
        )
      )
        throw error
      const result = await readFlatWithRecovery(loadSheetsDocument, clean)
      return {
        document: result.doc,
        isPackage: false,
        recoveredFromBackup: result.recovered,
        path: clean,
      }
    }
    return {
      document: loadSheetsDocument(raw),
      isPackage: true,
      recoveredFromBackup: false,
      path: clean,
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

/** Reject NUL bytes, encoded/raw parent-directory traversal segments, and optional relative paths. */
export function validateSafeFilePath(
  path: string,
  options: { requireAbsolute?: boolean } = {},
): string {
  const trimmed = path.trim()
  if (!trimmed) throw new Error('A file path is required.')
  if (trimmed.includes('\0') || /%00/i.test(trimmed)) {
    throw new Error('File path contains an invalid NUL byte.')
  }
  const hasTraversalSegment = (candidate: string): boolean =>
    candidate.split(/[\\/]/).some(segment => segment === '..')
  if (
    hasTraversalSegment(trimmed) ||
    /(?:^|[\\/])(?:%2e|\.)(?:%2e|\.)(?:$|[\\/])/i.test(trimmed)
  ) {
    throw new Error('File path must not contain ".." traversal segments.')
  }
  try {
    const decoded = decodeURIComponent(trimmed)
    if (decoded.includes('\0')) {
      throw new Error('File path contains an invalid NUL byte.')
    }
    if (hasTraversalSegment(decoded)) {
      throw new Error('File path must not contain ".." traversal segments.')
    }
  } catch (error) {
    if (error instanceof Error && error.message.startsWith('File path ')) {
      throw error
    }
  }
  if (
    options.requireAbsolute &&
    !(
      trimmed.startsWith('/') ||
      /^[A-Za-z]:[\\/]/.test(trimmed) ||
      trimmed.startsWith('\\\\')
    )
  ) {
    throw new Error('An absolute file path is required.')
  }
  return trimmed
}
