import { expect, it, vi } from 'vitest'
import { readWorkbookFromPath, workbookFilesForSave } from './workbookFiles'
import { createBlankWorkbook } from './workbookModel'
import { serializeSheetsHtmlDocument } from './sheetsDocument'
const workbook = createBlankWorkbook('Recovered')
it('recovers a missing or corrupt legacy main file from its backup', async () => {
  for (const missing of [true, false]) {
    const read = vi.fn(async (path: string) => {
      if (path.endsWith('/workbook.json')) throw Error('ENOTDIR')
      if (path.endsWith('.bak')) return JSON.stringify(workbook)
      if (missing) throw Error('ENOENT')
      return '{broken'
    })
    const loaded = await readWorkbookFromPath('/Legacy.sheets', read)
    expect(loaded.document.metadata.title).toBe('Recovered')
    expect(loaded.recoveredFromBackup).toBe(true)
    expect(loaded.isPackage).toBe(false)
  }
})
it('recovers unreadable HTML from an HTML backup and preserves its original format', async () => {
  const read = vi.fn(async (path: string) => {
    if (path.endsWith('.bak')) return serializeSheetsHtmlDocument(workbook)
    throw Error('ENOENT')
  })
  const loaded = await readWorkbookFromPath('/Legacy.sheets.html', read)
  expect(loaded.recoveredFromBackup).toBe(true)
  const saved = workbookFilesForSave(loaded.document, {
    boundPath: loaded.path,
    isPackage: loaded.isPackage,
  })
  expect(saved[0].name).toBeNull()
  expect(saved[0].content).toContain('<!doctype html>')
})
it('preserves actual package errors and never guesses a flat file after corrupt package JSON', async () => {
  for (const error of [
    Error('Permission denied'),
    Error('Connection timed out'),
  ]) {
    const read = vi.fn(async () => {
      throw error
    })
    await expect(readWorkbookFromPath('/Broken.sheets', read)).rejects.toBe(
      error,
    )
    expect(read).toHaveBeenCalledTimes(1)
  }
  const read = vi.fn(async () => '{broken')
  await expect(readWorkbookFromPath('/Broken.sheets', read)).rejects.toThrow()
  expect(read).toHaveBeenCalledTimes(1)
})
it('recognizes uppercase package suffixes', async () => {
  const read = vi.fn(async () => JSON.stringify(workbook))
  const loaded = await readWorkbookFromPath('/Upper.SHEETS/', read)
  expect(loaded.isPackage).toBe(true)
  expect(read).toHaveBeenCalledWith('/Upper.SHEETS/workbook.json')
})
