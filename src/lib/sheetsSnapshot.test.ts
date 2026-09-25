import { describe, expect, it } from 'vitest'
import { sheetsSnapshotHtml } from './sheetsSnapshot'
import type { PureSheetsDocument } from '../types'

function doc(cells: Record<string, { value: string; kind: 'text' | 'number' }>): PureSheetsDocument {
  return {
    app: 'PureSheets',
    version: 1,
    engine: { name: 'univer' },
    metadata: { title: 'Q3', createdAt: '', updatedAt: '' },
    workbook: {
      activeSheetId: 's1',
      sheets: [
        {
          id: 's1',
          name: 'Sheet 1',
          rowCount: 20,
          columnCount: 10,
          cells,
        },
      ],
    },
  } as PureSheetsDocument
}

describe('sheetsSnapshotHtml', () => {
  it('renders values into the grid, headers bold, numbers right-aligned', () => {
    const html = sheetsSnapshotHtml(
      doc({
        A1: { value: 'Region', kind: 'text' },
        B2: { value: '1200', kind: 'number' },
      }),
    )!
    expect(html).toContain('Region')
    expect(html).toContain('font-weight="700"')
    expect(html).toContain('text-anchor="end"')
  })

  it('escapes values and returns null for empty sheets', () => {
    expect(
      sheetsSnapshotHtml(doc({ A1: { value: '<b>&x', kind: 'text' } })),
    ).toContain('&lt;b&gt;&amp;x')
    expect(sheetsSnapshotHtml(doc({}))).toBeNull()
  })

  it('previews a populated sheet when the active sheet is empty without changing selection', () => {
    const document = doc({})
    document.workbook.sheets.push({ ...document.workbook.sheets[0], id: 's2', name: 'Procedure', cells: { A1: { value: 'Monday', kind: 'text' } } })
    expect(sheetsSnapshotHtml(document)).toContain('Monday')
    expect(document.workbook.activeSheetId).toBe('s1')
    document.workbook.sheets[0].cells.A1 = { value: 'Active', kind: 'text' }
    expect(sheetsSnapshotHtml(document)).toContain('Active')
    expect(sheetsSnapshotHtml(document)).not.toContain('Monday')
  })
})
