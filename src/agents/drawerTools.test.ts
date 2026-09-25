import { describe, expect, it } from 'vitest'
import { getSheetsContextHandler, readRangeHandler, applyCellChangesHandler } from './handlers'
import { createBlankWorkbook, lockRange } from '../lib/workbookModel'
import type { SheetsAgentToolContext } from './catalog'

function fixture() {
  const context: SheetsAgentToolContext = {
    document: createBlankWorkbook('Drawer test'), filePath: '/test.sheets', selectedRange: 'A1:A1',
    setDocument: update => { context.document = update(context.document!) },
  }
  return context
}
const parse = (result: { content: string }) => JSON.parse(result.content)

describe('drawer workbook tools', () => {
  it('reads, writes a formula and reads its computed result with log and fresh version', async () => {
    const context = fixture()
    const before = parse(await getSheetsContextHandler(context))
    const write = parse(await applyCellChangesHandler(context, { sheet: before.sheets[0].id, range: 'A1:B1', expectedVersion: before.version, summary: 'Source plus total', cellChanges: [{ cell: 'A1', value: '12', kind: 'number' }, { cell: 'B1', value: '=SUM(A1:A1)', kind: 'formula' }] }))
    expect(write.applied).toBe(2)
    const read = parse(await readRangeHandler(context, { range: 'A1:B1' }))
    expect(read.cells[1]).toMatchObject({ cell: 'B1', computed: '12' })
    expect(read.version).not.toBe(before.version)
    expect(context.document!.workbook.agentLog?.[0].cells).toHaveLength(2)
  })
  it('rejects a stale snapshot, including file-binding changes', async () => {
    const context = fixture()
    const before = parse(await getSheetsContextHandler(context))
    context.filePath = '/different.sheets'
    const original = JSON.stringify(context.document)
    const result = await applyCellChangesHandler(context, { sheet: before.sheets[0].id, range: 'A1:A1', expectedVersion: before.version, summary: 'Late result', cellChanges: [{ cell: 'A1', value: 'bad', kind: 'text' }] })
    expect(result.isError).toBe(true)
    expect(JSON.stringify(context.document)).toBe(original)
  })
  it('rejects a concurrent change inside the write callback', async () => {
    const context = fixture()
    const before = parse(await getSheetsContextHandler(context))
    context.setDocument = update => { context.document = update(createBlankWorkbook('New workbook')) }
    const result = await applyCellChangesHandler(context, { sheet: before.sheets[0].id, range: 'A1:A1', expectedVersion: before.version, summary: 'Late result', cellChanges: [{ cell: 'A1', value: 'bad', kind: 'text' }] })
    expect(result.isError).toBe(true)
    expect(context.document!.workbook.sheets[0].cells).toEqual({})
  })
  it('rejects changes outside the declared range and malformed batches atomically', async () => {
    const context = fixture()
    const before = parse(await getSheetsContextHandler(context))
    for (const cellChanges of [[{ cell: 'B1', value: 'bad', kind: 'text' }], [{ cell: 'A1', value: 'good', kind: 'text' }, null]]) {
      const result = await applyCellChangesHandler(context, { sheet: before.sheets[0].id, range: 'A1:A1', expectedVersion: before.version, summary: 'Invalid batch', cellChanges })
      expect(result.isError).toBe(true)
      expect(context.document!.workbook.sheets[0].cells).toEqual({})
    }
  })
  it('keeps locked cells unchanged and bounds reads before allocation', async () => {
    const context = fixture()
    const sheet = context.document!.workbook.sheets[0]
    context.document!.workbook.sheets[0] = lockRange(sheet, 'A1:A1')
    const before = parse(await getSheetsContextHandler(context))
    const result = await applyCellChangesHandler(context, { sheet: sheet.id, range: 'A1:A1', expectedVersion: before.version, summary: 'Locked', cellChanges: [{ cell: 'A1', value: 'bad', kind: 'text' }] })
    expect(result.isError).toBe(true)
    expect((await readRangeHandler(context, { range: 'A1:XFD1048576' })).isError).toBe(true)
  })
})
