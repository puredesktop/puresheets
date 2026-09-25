/**
 * Agent write invariants for the direct-write model.
 *
 * `applyAgentCellChangesDirect` is the only cell-write path for agent tools:
 * every change passes the trust boundary (valid keys, bounded growth, capped
 * batch), locked cells are skipped and reported (never mutated), and each
 * write is recorded in the workbook's agent log with before/after — the log,
 * not a review queue, is the accountability mechanism.
 */
import { describe, expect, it } from 'vitest'
import {
  applyAgentCellChangesDirect,
  createBlankWorkbook,
  lockRange,
} from './workbookModel'
import { loadSheetsDocument, serializeSheetsDocument } from './sheetsDocument'

describe('direct agent writes with log', () => {
  it('applies cell changes directly and records before/after in the agent log', () => {
    let document = createBlankWorkbook('Log test')
    const sheet = document.workbook.sheets[0]
    const first = applyAgentCellChangesDirect(
      document,
      sheet.id,
      'A1:B1',
      'Seed the header row.',
      [
        { cell: 'A1', value: 'Region', kind: 'text' },
        { cell: 'B1', value: '=1+1', kind: 'formula' },
      ],
    )
    expect(first.applied).toBe(2)
    document = first.document
    expect(document.workbook.sheets[0].cells.A1?.value).toBe('Region')
    expect(document.workbook.agentLog?.[0]).toMatchObject({
      tool: 'applyCellChanges',
      summary: 'Seed the header row.',
      sheetName: sheet.name,
    })
    expect(document.workbook.agentLog?.[0]?.cells).toHaveLength(2)

    const second = applyAgentCellChangesDirect(
      document,
      sheet.id,
      'A1:A1',
      'Rename the header.',
      [{ cell: 'A1', value: 'Territory', kind: 'text' }],
    )
    expect(
      second.document.workbook.agentLog?.[0]?.cells?.[0],
    ).toMatchObject({ cell: 'A1', before: 'Region', after: 'Territory' })
  })

  it('skips locked cells in direct writes and reports the skip', () => {
    let document = createBlankWorkbook('Lock test')
    const sheet = document.workbook.sheets[0]
    const locked = lockRange(sheet, 'A1:A1')
    document = {
      ...document,
      workbook: {
        ...document.workbook,
        sheets: [locked, ...document.workbook.sheets.slice(1)],
      },
    }
    const result = applyAgentCellChangesDirect(
      document,
      sheet.id,
      'A1:B1',
      'Attempt over a locked cell.',
      [
        { cell: 'A1', value: 'nope', kind: 'text' },
        { cell: 'B1', value: 'yes', kind: 'text' },
      ],
    )
    expect(result.applied).toBe(1)
    expect(result.skippedLocked).toBe(1)
    expect(result.document.workbook.sheets[0].cells.A1).toBeUndefined()
  })

  it('rejects invalid or out-of-bounds cell changes at the trust boundary', () => {
    const document = createBlankWorkbook('Trust boundary')
    const sheet = document.workbook.sheets[0]
    const result = applyAgentCellChangesDirect(
      document,
      sheet.id,
      'A1:B2',
      'Mixed valid and invalid keys.',
      [
        { cell: 'A1', value: 'alpha', kind: 'text' },
        { cell: 'b2', value: 'beta', kind: 'text' }, // lowercase key normalizes
        { cell: '../etc/passwd', value: 'junk', kind: 'text' }, // invalid key
        { cell: 'ZZZ99999', value: 'far away', kind: 'text' }, // beyond bounds
        { cell: 'B2', value: 'duplicate', kind: 'text' }, // dupe of b2
      ],
    )
    expect(result.applied).toBe(2)
    const written = result.document.workbook.sheets[0].cells
    expect(written.A1?.value).toBe('alpha')
    expect(written.B2?.value).toBe('beta')
    expect(Object.keys(written)).not.toContain('../etc/passwd')
    expect(Object.keys(written)).not.toContain('ZZZ99999')
  })

  it('the agent log survives a save/reload round trip', () => {
    const document = createBlankWorkbook('Persistence')
    const sheet = document.workbook.sheets[0]
    const written = applyAgentCellChangesDirect(
      document,
      sheet.id,
      'A1:A1',
      'Persisted write.',
      [{ cell: 'A1', value: '42', kind: 'number' }],
    ).document
    const reloaded = loadSheetsDocument(serializeSheetsDocument(written))
    expect(reloaded.workbook.agentLog?.[0]).toMatchObject({
      tool: 'applyCellChanges',
      summary: 'Persisted write.',
    })
    expect(reloaded.workbook.agentLog?.[0]?.cells?.[0]).toMatchObject({
      cell: 'A1',
      after: '42',
    })
  })
})
