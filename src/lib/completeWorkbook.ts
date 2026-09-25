/** App-owned durable completion. A source file is never a workbook receipt. */
export async function completeWorkbook(io: {
  save: () => Promise<{ path: string; contentPath: string; content: string }>
  read: (path: string) => Promise<string>
  unchanged: () => boolean
}) {
  const saved = await io.save()
  if (!saved.path || !saved.contentPath || !saved.content.trim())
    throw new Error('Workbook has no saved content.')
  if (await io.read(saved.contentPath) !== saved.content)
    throw new Error('Saved workbook differs from the current workbook. Retry completion after refreshing.')
  if (!io.unchanged())
    throw new Error('Workbook changed during completion. Complete the current workbook again.')
  return {
    taskOutcome: 'Workbook saved and its content verified by PureSheets.',
    artifactPaths: [...new Set([saved.path, saved.contentPath])],
    observations: ['Saved bytes match the current workbook. This does not establish semantic or visual correctness.'],
  }
}
