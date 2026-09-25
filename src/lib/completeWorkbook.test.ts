import { describe, expect, it, vi } from 'vitest'
import { completeWorkbook } from './completeWorkbook'

const setup = () => ({
  save: vi.fn(async () => ({ path: '/new.sheets', contentPath: '/new.sheets/workbook.json', content: '{"cells":1}' })),
  read: vi.fn(async () => '{"cells":1}'),
  unchanged: () => true,
})
describe('durable workbook completion', () => {
  it('returns the actual saved workbook paths', async () => {
    const io = setup()
    expect((await completeWorkbook(io)).artifactPaths).toEqual(['/new.sheets', '/new.sheets/workbook.json'])
    expect(io.read).toHaveBeenCalledWith('/new.sheets/workbook.json')
  })
  it('propagates save failure without a receipt', async () => {
    const io = setup(); io.save.mockRejectedValue(new Error('disk full'))
    await expect(completeWorkbook(io)).rejects.toThrow('disk full')
    expect(io.read).not.toHaveBeenCalled()
  })
  it('rejects different saved bytes', async () => {
    const io = setup(); io.read.mockResolvedValue('old workbook')
    await expect(completeWorkbook(io)).rejects.toThrow('differs')
  })
  it('rejects edits or document switches during saving', async () => {
    await expect(completeWorkbook({ ...setup(), unchanged: () => false })).rejects.toThrow('changed')
  })
  it('rejects unreadable output', async () => {
    const io = setup(); io.read.mockRejectedValue(new Error('unreadable'))
    await expect(completeWorkbook(io)).rejects.toThrow('unreadable')
  })
})
