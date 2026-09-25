// @vitest-environment happy-dom
import { act } from 'react'
import type React from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createBlankWorkbook } from '../lib/workbookModel'
import { PureSheetsShell } from './PureSheetsShell'
import type { SheetsAgentToolContext } from '../agents/catalog'

let agentContext: SheetsAgentToolContext
const surfaceRender = vi.hoisted(() => vi.fn())
vi.mock('./UniverSpreadsheetSurface', async importOriginal => {
  const actual = await importOriginal<typeof import('./UniverSpreadsheetSurface')>()
  return {
    ...actual,
    UniverSpreadsheetSurface: (props: React.ComponentProps<typeof actual.UniverSpreadsheetSurface>) => {
      surfaceRender(props)
      return <actual.UniverSpreadsheetSurface {...props} />
    },
  }
})
vi.mock('../hooks/useSheetsAgentTools', () => ({
  useSheetsAgentTools: (_ready: boolean, context: SheetsAgentToolContext) => { agentContext = context },
}))

const readTextFile = vi.fn()
const writeTextFile = vi.fn()
const readClipboard = vi.fn()
const writeClipboard = vi.fn()
const readBinaryFile = vi.fn()
const updatePureSheetsSettings = vi.fn()
const createPlatformDraft = vi.fn(async (..._args: unknown[]) => ({
  path: '/workspace/PureDrafts/Untitled — Test.sheets',
}))
const autosavePlatformDocument = vi.fn(async (..._args: unknown[]) => ({
  savedAt: '2026-07-02T12:00:00.000Z',
}))
const promotePlatformDocument = vi.fn(async (...args: unknown[]) => ({
  path: `/workspace/PureDesktop/${
    (args[0] as { title?: string } | undefined)?.title ?? 'Untitled'
  }.sheets`,
}))
const renamePlatformDocument = vi.fn(async (...args: unknown[]) => ({
  path: `/workspace/PureDesktop/${
    (args[0] as { title?: string } | undefined)?.title ?? 'Untitled'
  }.sheets`,
}))

;(
  globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true

vi.mock('@purescience/platform-ui/bridge/documents.mjs', () => ({
  createPlatformDraft: (...args: unknown[]) => createPlatformDraft(...args),
  autosavePlatformDocument: (...args: unknown[]) =>
    autosavePlatformDocument(...args),
  promotePlatformDocument: (...args: unknown[]) =>
    promotePlatformDocument(...args),
  renamePlatformDocument: (...args: unknown[]) =>
    renamePlatformDocument(...args),
  duplicatePlatformDocument: vi.fn(async () => ({ path: '/workspace/copy' })),
  listPlatformDocumentsByType: vi.fn(async () => []),
  listPlatformRecentDocuments: vi.fn(async () => []),
  touchPlatformRecentDocument: vi.fn(async () => undefined),
  suggestPlatformDocumentLocation: vi.fn(async () => ({ dir: '/workspace' })),
}))

vi.mock('../bridge/platformBridge', () => ({
  readBinaryFile: (...args: unknown[]) => readBinaryFile(...args),
  readClipboard: (...args: unknown[]) => readClipboard(...args),
  readTextFile: (...args: unknown[]) => readTextFile(...args),
  writeClipboard: (...args: unknown[]) => writeClipboard(...args),
  writeTextFile: (...args: unknown[]) => writeTextFile(...args),
  updatePureSheetsSettings: (...args: unknown[]) =>
    updatePureSheetsSettings(...args),
}))

// Roots are unmounted in afterEach. Without this, each test's mounted Univer
// editor is never torn down, so its global editor state leaks into later tests
// (a prior sheet's cells bleeding into an unrelated document), making the suite
// order-dependent and flaky.
const mountedRoots: Root[] = []

function render(ui: React.ReactElement): {
  container: HTMLDivElement
  root: Root
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  mountedRoots.push(root)
  act(() => {
    root.render(ui)
  })
  return { container, root }
}

function triggerSave(): void {
  window.dispatchEvent(
    new KeyboardEvent('keydown', {
      bubbles: true,
      key: 's',
      metaKey: true,
    }),
  )
}

async function flushAsync(count = 16): Promise<void> {
  for (let index = 0; index < count; index += 1) {
    await Promise.resolve()
  }
}

/** Workbook JSON from the last lifecycle save — the `workbook.json` entry of
 * a `.sheets` package payload (or the sole file of a legacy flat write). */
function lastSavedContent(): string {
  type FilePayload = { name: string | null; content?: string }
  const pickWorkbook = (files?: FilePayload[]): string | undefined => {
    if (!files?.length) return undefined
    const workbook = files.find(file => file.name === 'workbook.json')
    return (workbook ?? files[0])?.content
  }
  const autosaveCall = autosavePlatformDocument.mock.calls.at(-1) as
    | [{ files?: FilePayload[] }]
    | undefined
  const fromAutosave = pickWorkbook(autosaveCall?.[0]?.files)
  if (fromAutosave !== undefined) return String(fromAutosave)
  const draftCall = createPlatformDraft.mock.calls.at(-1) as
    | [{ files?: FilePayload[] }]
    | undefined
  return String(pickWorkbook(draftCall?.[0]?.files) ?? '')
}

describe('agent workbook completion integration', () => {
  it('keeps the latest explicit open when an older read finishes afterward', async () => {
    let finishOld!: (value: string) => void
    const old = createBlankWorkbook('Slow old file')
    const latest = createBlankWorkbook('Latest file')
    readTextFile.mockImplementation((path: string) => path.startsWith('/workspace/old.sheets')
      ? new Promise<string>(resolve => { finishOld = resolve })
      : Promise.resolve(JSON.stringify(latest)))
    const handledOld = vi.fn(), handledLatest = vi.fn()
    const { root } = render(<PureSheetsShell initialDocument={createBlankWorkbook()}
      resource={{ path: '/workspace/old.sheets' }} onResourceHandled={handledOld} />)
    await act(async () => { await flushAsync() })
    await act(async () => {
      root.render(<PureSheetsShell initialDocument={createBlankWorkbook()}
        resource={{ path: '/workspace/latest.sheets' }} onResourceHandled={handledLatest} />)
    })
    await act(async () => { await flushAsync() })
    await act(async () => { finishOld(JSON.stringify(old)); await flushAsync() })
    expect(agentContext.document?.metadata.title).toBe('Latest file')
    expect(agentContext.filePath).toBe('/workspace/latest.sheets')
    expect(handledOld).not.toHaveBeenCalled()
    expect(handledLatest).toHaveBeenCalledTimes(1)
    expect(updatePureSheetsSettings).toHaveBeenCalledTimes(1)
  })

  it('does not let a delayed boot read replace an explicitly created workbook', async () => {
    const old = createBlankWorkbook('Delayed old workbook')
    old.workbook.sheets[0].cells.A1 = { value: 'Old content', kind: 'text' }
    let resolveRead!: (value: string) => void
    readTextFile.mockImplementation(() => new Promise<string>(resolve => { resolveRead = resolve }))
    render(<PureSheetsShell initialDocument={createBlankWorkbook()} resource={{ path: '/workspace/old.sheets' }} />)
    await act(async () => { await flushAsync() })
    expect(readTextFile).toHaveBeenCalled()
    await act(async () => { await agentContext.createWorkbook!('New mission output') })
    await act(async () => { resolveRead(JSON.stringify(old)); await flushAsync() })
    expect(agentContext.document?.metadata.title).toBe('New mission output')
    expect(agentContext.filePath).toBe('')
    expect(agentContext.document?.workbook.sheets[0].cells).toEqual({})
    expect(autosavePlatformDocument).not.toHaveBeenCalled()
    expect(updatePureSheetsSettings).not.toHaveBeenCalled()
  })

  it('keeps the current workbook when flushing edits before creation fails', async () => {
    render(<PureSheetsShell initialDocument={createBlankWorkbook('Keep current')} />)
    readTextFile.mockImplementation(async () => lastSavedContent())
    await act(async () => { await agentContext.completeWorkbook!() })
    const existingPath = agentContext.filePath
    await act(async () => {
      agentContext.setDocument(current => ({ ...current, metadata: { ...current.metadata, title: 'Unsaved edit' } }))
    })
    autosavePlatformDocument.mockRejectedValueOnce(new Error('cannot flush'))
    await act(async () => {
      await expect(agentContext.createWorkbook!('Must not appear')).rejects.toThrow('cannot flush')
    })
    expect(agentContext.filePath).toBe(existingPath)
    expect(agentContext.document?.metadata.title).toBe('Unsaved edit')
  })

  it('creates a separate workbook from an existing file and saves only the new content there', async () => {
    const old = createBlankWorkbook('Existing workbook')
    old.workbook.sheets[0].cells.A1 = { value: 'Keep this', kind: 'text' }
    readTextFile.mockResolvedValue(JSON.stringify(old))
    render(<PureSheetsShell initialDocument={createBlankWorkbook()} resource={{ path: '/workspace/existing.sheets' }} />)
    await act(async () => { await flushAsync() })
    expect(agentContext.filePath).toBe('/workspace/existing.sheets')
    await act(async () => { await agentContext.createWorkbook!('Fresh deliverable') })
    expect(autosavePlatformDocument).not.toHaveBeenCalled()
    expect(agentContext.filePath).toBe('')
    expect(agentContext.document?.metadata.title).toBe('Fresh deliverable')
    expect(agentContext.document?.workbook.sheets[0].cells).toEqual({})
    await act(async () => {
      agentContext.setDocument(current => ({ ...current, workbook: { ...current.workbook,
        sheets: current.workbook.sheets.map(sheet => ({ ...sheet, cells: { A1: { value: 'New output', kind: 'text' as const } } })) } }))
      await flushAsync()
    })
    readTextFile.mockImplementation(async () => lastSavedContent())
    await act(async () => { await agentContext.completeWorkbook!() })
    expect(agentContext.filePath).toBe('/workspace/PureDrafts/Untitled — Test.sheets')
    expect(lastSavedContent()).toContain('New output')
    for (const [args] of autosavePlatformDocument.mock.calls as [{ path: string; files: unknown }][]) {
      if (args.path === '/workspace/existing.sheets') expect(JSON.stringify(args.files)).not.toContain('New output')
    }
  })

  it('creates a package and verifies the saved current workbook', async () => {
    render(<PureSheetsShell initialDocument={createBlankWorkbook()} />)
    readTextFile.mockImplementation(async () => lastSavedContent())
    // Serialization touches timestamps. Readback must compare the written
    // payload, not a new serialization after storage latency.
    autosavePlatformDocument.mockImplementationOnce(async () => {
      await new Promise(resolve => setTimeout(resolve, 20))
      return { savedAt: new Date().toISOString() }
    })
    let receipt: Awaited<ReturnType<NonNullable<SheetsAgentToolContext['completeWorkbook']>>> | undefined
    await act(async () => { receipt = await agentContext.completeWorkbook!() })
    expect(receipt?.artifactPaths).toEqual([
      '/workspace/PureDrafts/Untitled — Test.sheets',
      '/workspace/PureDrafts/Untitled — Test.sheets/workbook.json',
    ])
    expect(JSON.parse(lastSavedContent()).workbook.sheets.length).toBeGreaterThan(0)
  })

  it('cannot issue a receipt when persistence rejects', async () => {
    render(<PureSheetsShell initialDocument={createBlankWorkbook()} />)
    autosavePlatformDocument.mockRejectedValueOnce(new Error('disk full'))
    await act(async () => {
      await expect(agentContext.completeWorkbook!()).rejects.toThrow('disk full')
    })
  })

  it('rejects readback that does not match the saved workbook', async () => {
    render(<PureSheetsShell initialDocument={createBlankWorkbook()} />)
    readTextFile.mockResolvedValueOnce('stale bytes')
    await act(async () => {
      await expect(agentContext.completeWorkbook!()).rejects.toThrow('differs')
    })
  })

  it('rejects a workbook edit while saved bytes are being verified', async () => {
    render(<PureSheetsShell initialDocument={createBlankWorkbook()} />)
    let finishRead!: (value: string) => void
    readTextFile.mockImplementationOnce(() => new Promise<string>(resolve => { finishRead = resolve }))
    let completion!: Promise<unknown>
    await act(async () => {
      completion = agentContext.completeWorkbook!()
      await flushAsync()
    })
    const rejection = expect(completion).rejects.toThrow('changed')
    await act(async () => {
      agentContext.setDocument(current => ({ ...current, metadata: { ...current.metadata, title: 'Changed during completion' } }))
    })
    await act(async () => { finishRead(lastSavedContent()); await rejection })
  })
})

afterEach(() => {
  // Unmount every rendered shell — zombie instances keep window listeners
  // (⌘S, lifecycle flush) alive and pollute later tests' save payloads.
  for (const root of mountedRoots.splice(0)) {
    act(() => root.unmount())
  }
  document.body.innerHTML = ''
  vi.clearAllMocks()
  vi.useRealTimers()
})

describe('PureSheetsShell file safety', () => {
  it('refreshes the engine for agent edits and a new workbook with the same title', async () => {
    render(<PureSheetsShell initialDocument={createBlankWorkbook()} />)
    const lastSurface = () => surfaceRender.mock.calls.at(-1)![0]
    const initialRevision = lastSurface().revision
    const id = lastSurface().snapshot.id
    await act(async () => {
      agentContext.setDocument(current => ({ ...current, workbook: { ...current.workbook,
        sheets: current.workbook.sheets.map(sheet => ({ ...sheet, cells: { A1: { value: 'Agent content', kind: 'text' as const } } })) } }))
    })
    expect(lastSurface().revision).toBe(initialRevision + 1)
    expect(lastSurface().snapshot.id).toBe(id)
    expect(lastSurface().snapshot.sheets['sheet-1'].cellData[0][0].v).toBe('Agent content')
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'n', metaKey: true }))
      await flushAsync()
    })
    expect(lastSurface().revision).toBe(initialRevision + 2)
    expect(lastSurface().snapshot.id).toBe(id)
    expect(lastSurface().snapshot.sheets['sheet-1'].cellData).toEqual({})
  })

  it('requests a canvas reload on each successful same-ID reopen, but not a failed read', async () => {
    const initialDocument = createBlankWorkbook('Same workbook')
    readTextFile.mockResolvedValue(JSON.stringify(initialDocument))
    updatePureSheetsSettings.mockResolvedValue(undefined)
    const { root } = render(<PureSheetsShell initialDocument={initialDocument} />)
    const lastSurface = () => surfaceRender.mock.calls.at(-1)![0]
    const initialRevision = lastSurface().revision
    const open = async () => {
      act(() => root.render(<PureSheetsShell initialDocument={initialDocument} />))
      await act(async () => {
        root.render(<PureSheetsShell initialDocument={initialDocument} resource={{ path: '/workspace/same.sheets' }} />)
        await flushAsync()
      })
    }
    await open()
    const firstId = lastSurface().snapshot.id
    expect(lastSurface().revision).toBe(initialRevision + 1)
    await open()
    expect(lastSurface().snapshot.id).toBe(firstId)
    expect(lastSurface().revision).toBe(initialRevision + 2)
    readTextFile.mockRejectedValue(new Error('read failed'))
    await open()
    expect(lastSurface().snapshot.id).toBe(firstId)
    expect(lastSurface().revision).toBe(initialRevision + 2)
  })

  it('keeps a corrupt opened spreadsheet untouched and recoverable', async () => {
    readTextFile.mockResolvedValue('{not valid sheets json')
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)
    const onResourceHandled = vi.fn()

    render(
      <PureSheetsShell
        initialDocument={createBlankWorkbook('Unsaved')}
        resource={{ path: '/workspace/broken.sheets' }}
        onResourceHandled={onResourceHandled}
      />,
    )

    await act(async () => {
      await Promise.resolve()
    })

    expect(readTextFile).toHaveBeenCalledWith('/workspace/broken.sheets')
    expect(onResourceHandled).toHaveBeenCalled()
    expect(document.body.textContent).toContain('JSON')

    await act(async () => {
      triggerSave()
      await Promise.resolve()
    })

    expect(writeTextFile).not.toHaveBeenCalled()
    expect(updatePureSheetsSettings).not.toHaveBeenCalled()
  })

  it('normalizes the editable selected range and uses it for chart actions', async () => {
    let workbookDocument = createBlankWorkbook('Range test')
    workbookDocument = {
      ...workbookDocument,
      workbook: {
        ...workbookDocument.workbook,
        sheets: [
          {
            ...workbookDocument.workbook.sheets[0],
            cells: {
              A1: { value: 'North', kind: 'text' },
              B1: { value: '12', kind: 'number' },
              A2: { value: 'South', kind: 'text' },
              B2: { value: '8', kind: 'number' },
            },
          },
        ],
      },
    }
    render(<PureSheetsShell initialDocument={workbookDocument} />)

    const rangeInput = document.querySelector(
      'input[aria-label="Selected range"]',
    )
    if (!(rangeInput instanceof HTMLInputElement)) {
      throw new Error('Selected range input not found')
    }
    const chartMenuTrigger = document.querySelector(
      'button[aria-label="Insert chart from selected range"]',
    )
    if (!(chartMenuTrigger instanceof HTMLButtonElement)) {
      throw new Error('Chart menu trigger not found')
    }

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(rangeInput, 'b2:a1')
      rangeInput.dispatchEvent(new Event('input', { bubbles: true }))
      rangeInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    expect(rangeInput.value).toBe('A1:B2')

    await act(async () => {
      chartMenuTrigger.click()
    })
    const barOption = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[data-platform-menu-surface] [role="menuitem"]',
      ),
    ).find(item => item.textContent?.trim() === 'Bar')
    if (!barOption) throw new Error('Bar chart menu item not found')
    await act(async () => {
      barOption.click()
    })

    expect(document.body.textContent).toContain('Created bar chart from A1:B2')
    expect(document.body.textContent).toContain('A1:B2 · bar')
  })

  it('toggles first-row and first-column freeze controls', async () => {
    render(
      <PureSheetsShell initialDocument={createBlankWorkbook('Freeze test')} />,
    )

    const freezeRow = document.querySelector('button[title="Freeze first row"]')
    const freezeColumn = document.querySelector(
      'button[title="Freeze first column"]',
    )
    if (!(freezeRow instanceof HTMLButtonElement)) {
      throw new Error('Freeze first row button not found')
    }
    if (!(freezeColumn instanceof HTMLButtonElement)) {
      throw new Error('Freeze first column button not found')
    }

    await act(async () => {
      freezeRow.click()
      freezeColumn.click()
    })

    expect(document.body.textContent).toContain('Row frozen.')
    expect(document.body.textContent).toContain('Column frozen.')
  })

  it('renders the enterprise spreadsheet surface with accessible toolbar controls', () => {
    render(
      <PureSheetsShell
        initialDocument={createBlankWorkbook('UI requirements')}
      />,
    )

    expect(
      document.querySelector('[data-testid="univer-spreadsheet-surface"]'),
    ).toBeTruthy()
    expect(document.querySelector('[aria-label="Formula bar"]')).toBeTruthy()
    expect(
      document.querySelector('[aria-label="Formula or value"]'),
    ).toBeTruthy()
    expect(document.querySelector('[role="columnheader"]')?.textContent).toBe(
      'A',
    )
    expect(document.body.textContent).toContain('Sheet 1')

    const toolbarButtons = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[aria-label="Spreadsheet toolbar"] button',
      ),
    )
    const accessibleLabels = toolbarButtons.map(button =>
      button.getAttribute('aria-label'),
    )
    expect(accessibleLabels).not.toContain('Save')
    expect(accessibleLabels).toContain('Undo')
    expect(accessibleLabels).toContain('Redo')
    expect(accessibleLabels).toContain('Freeze first row')
    expect(accessibleLabels).toContain('Freeze first column')
    expect(accessibleLabels).toContain('Clear selected range')
    expect(accessibleLabels).toContain('Insert row')
    expect(accessibleLabels).toContain('Insert column')
    expect(accessibleLabels).toContain('Find next')
    expect(accessibleLabels).toContain('Replace matches in selected range')
    // Dropdowns are custom menus, never native <select> (enterprise rules).
    expect(
      document.querySelector(
        'button[aria-label="Insert formula or operation"]',
      ),
    ).toBeTruthy()
    expect(
      document.querySelector('button[aria-label="Number format"]'),
    ).toBeTruthy()
    expect(
      document.querySelector('button[aria-label="Font size"]'),
    ).toBeTruthy()
    expect(
      document.querySelector('[aria-label="Spreadsheet toolbar"] select'),
    ).toBeNull()
    expect(document.querySelector('label[title="Text color"]')).toBeTruthy()
    expect(document.querySelector('label[title="Fill color"]')).toBeTruthy()

    const mysteryLabels = new Set([
      'S',
      'U',
      'R',
      'FR',
      'FC',
      'W+',
      'W-',
      'H+',
      'H-',
    ])
    const visibleButtonLabels = toolbarButtons.map(
      button => button.textContent?.trim() ?? '',
    )
    expect(visibleButtonLabels.some(label => mysteryLabels.has(label))).toBe(
      false,
    )
    expect(document.head.textContent).toContain('content:attr(aria-label)')
    expect(document.head.textContent).toContain('content:attr(title)')
  })

  it('inserts named formula and operation templates from the toolbar', async () => {
    vi.useFakeTimers()
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)

    render(
      <PureSheetsShell
        initialDocument={createBlankWorkbook('Formula helper')}
      />,
    )

    const formulaTrigger = document.querySelector(
      'button[aria-label="Insert formula or operation"]',
    )
    const rangeInput = document.querySelector(
      'input[aria-label="Selected range"]',
    )
    if (!(formulaTrigger instanceof HTMLButtonElement)) {
      throw new Error('Formula menu trigger not found')
    }
    if (!(rangeInput instanceof HTMLInputElement)) {
      throw new Error('Range input not found')
    }

    await act(async () => {
      formulaTrigger.click()
    })

    const menuItems = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[data-platform-menu-surface] [role="menuitem"]',
      ),
    )
    expect(menuItems.map(item => item.textContent?.trim())).toEqual([
      'SUM',
      'AVERAGE',
      'MIN',
      'MAX',
      'COUNT',
      'COUNTA',
      'IF',
      'ROUND',
      'CONCAT',
      'LEFT',
      'RIGHT',
      'VLOOKUP',
      'COUNTIF',
      'SUMIF',
      'AVERAGEIF',
      'RANK',
      'IFERROR',
      'TODAY',
      'NOW',
      'Add (+)',
      'Subtract (-)',
      'Multiply (*)',
      'Divide (/)',
    ])
    const sumItem = menuItems.find(item => item.textContent?.trim() === 'SUM')
    if (!sumItem) throw new Error('SUM menu item not found')

    await act(async () => {
      sumItem.click()
    })

    expect(
      document.querySelector<HTMLInputElement>(
        'input[aria-label="Formula or value"]',
      )?.value,
    ).toBe('=SUM()')
    expect(document.body.textContent).toContain('Select a range for SUM')

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(rangeInput, 'B2:C3')
      rangeInput.dispatchEvent(new Event('input', { bubbles: true }))
      await vi.advanceTimersByTimeAsync(400)
    })

    expect(
      document.querySelector<HTMLInputElement>(
        'input[aria-label="Formula or value"]',
      )?.value,
    ).toBe('=SUM(B2:C3)')
    expect(document.body.textContent).toContain(
      'Inserted SUM formula from B2:C3',
    )

    await act(async () => {
      triggerSave()
      await flushAsync()
    })

    const saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A1.value).toBe('=SUM(B2:C3)')
    expect(saved.workbook.sheets[0].cells.A1.kind).toBe('formula')
  })

  it('autosaves edited cells after the debounce interval', async () => {
    vi.useFakeTimers()
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)

    render(
      <PureSheetsShell
        initialDocument={createBlankWorkbook('Autosave test')}
      />,
    )

    const formulaInput = document.querySelector(
      'input[aria-label="Formula or value"]',
    )
    if (!(formulaInput instanceof HTMLInputElement)) {
      throw new Error('Formula input not found')
    }

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, '42')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    expect(autosavePlatformDocument).not.toHaveBeenCalled()

    await act(async () => {
      await vi.advanceTimersByTimeAsync(1600)
    })

    expect(createPlatformDraft).toHaveBeenCalledTimes(1)
    expect(autosavePlatformDocument).toHaveBeenCalledTimes(1)
    expect(lastSavedContent()).toContain('"value": "42"')
    expect(writeTextFile).not.toHaveBeenCalled()
  })

  it('applies range formatting and structural toolbar edits to saved workbooks', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)

    render(
      <PureSheetsShell
        initialDocument={createBlankWorkbook('Toolbar operations')}
      />,
    )

    const formulaInput = document.querySelector(
      'input[aria-label="Formula or value"]',
    )
    const rangeInput = document.querySelector(
      'input[aria-label="Selected range"]',
    )
    const borderButton = document.querySelector(
      'button[title="Toggle all borders"]',
    )
    const insertRowButton = document.querySelector('button[title="Insert row"]')
    const insertColumnButton = document.querySelector(
      'button[title="Insert column"]',
    )
    if (!(formulaInput instanceof HTMLInputElement)) {
      throw new Error('Formula input not found')
    }
    if (!(rangeInput instanceof HTMLInputElement)) {
      throw new Error('Range input not found')
    }
    if (!(borderButton instanceof HTMLButtonElement)) {
      throw new Error('Border button not found')
    }
    if (!(insertRowButton instanceof HTMLButtonElement)) {
      throw new Error('Insert row button not found')
    }
    if (!(insertColumnButton instanceof HTMLButtonElement)) {
      throw new Error('Insert column button not found')
    }

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, '42')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(rangeInput, 'A1:B1')
      rangeInput.dispatchEvent(new Event('input', { bubbles: true }))
      rangeInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
      borderButton.click()
      insertRowButton.click()
      insertColumnButton.click()
    })

    await act(async () => {
      triggerSave()
      await flushAsync()
    })

    const saved = JSON.parse(lastSavedContent())
    const sheet = saved.workbook.sheets[0]
    expect(sheet.rowCount).toBe(41)
    expect(sheet.columnCount).toBe(19)
    expect(sheet.cells.B2.value).toBe('42')
    expect(sheet.cells.B2.style.border).toBe(true)
    expect(saved.engine.snapshot.styles).toBeTruthy()
  })

  it('replaces matches and clears selected ranges from the grid keyboard', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)

    render(
      <PureSheetsShell
        initialDocument={createBlankWorkbook('Replace operations')}
      />,
    )

    const formulaInput = document.querySelector(
      'input[aria-label="Formula or value"]',
    )
    const findInput = document.querySelector('input[aria-label="Find cells"]')
    const replaceInput = document.querySelector(
      'input[aria-label="Replace with"]',
    )
    const replaceButton = document.querySelector(
      'button[title="Replace matches in selected range"]',
    )
    const cellA1 = document.querySelector('button[aria-label="A1"]')
    if (!(formulaInput instanceof HTMLInputElement)) {
      throw new Error('Formula input not found')
    }
    if (!(findInput instanceof HTMLInputElement)) {
      throw new Error('Find input not found')
    }
    if (!(replaceInput instanceof HTMLInputElement)) {
      throw new Error('Replace input not found')
    }
    if (!(replaceButton instanceof HTMLButtonElement)) {
      throw new Error('Replace button not found')
    }

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, 'North forecast')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(findInput, 'North')
      findInput.dispatchEvent(new Event('input', { bubbles: true }))
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(replaceInput, 'South')
      replaceInput.dispatchEvent(new Event('input', { bubbles: true }))
      replaceButton.click()
    })

    expect(document.body.textContent).toContain('Replaced 1 match in A1:A1')

    await act(async () => {
      triggerSave()
      await flushAsync()
    })

    let saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A1.value).toBe('South forecast')

    const updatedCellA1 = document.querySelector('button[aria-label="A1"]')
    if (!(updatedCellA1 instanceof HTMLButtonElement)) {
      throw new Error('A1 cell not found')
    }

    await act(async () => {
      updatedCellA1.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Delete',
        }),
      )
    })

    await act(async () => {
      triggerSave()
      await Promise.resolve()
    })

    saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A1).toBeUndefined()
    expect(cellA1).toBeTruthy()
  })

  it('reports autosave write failures without losing dirty-state feedback', async () => {
    vi.useFakeTimers()
    autosavePlatformDocument.mockRejectedValue(new Error('Disk full'))

    render(
      <PureSheetsShell
        initialDocument={createBlankWorkbook('Autosave failure')}
      />,
    )

    const formulaInput = document.querySelector(
      'input[aria-label="Formula or value"]',
    )
    if (!(formulaInput instanceof HTMLInputElement)) {
      throw new Error('Formula input not found')
    }

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, 'Unsaved')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    await act(async () => {
      await vi.advanceTimersByTimeAsync(1600)
    })

    expect(autosavePlatformDocument).toHaveBeenCalledTimes(1)
    expect(writeTextFile).not.toHaveBeenCalled()
  })

  it('undoes and redoes formatting without losing the edited cell value', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)

    render(
      <PureSheetsShell
        initialDocument={createBlankWorkbook('Undo formatting')}
      />,
    )

    const formulaInput = document.querySelector(
      'input[aria-label="Formula or value"]',
    )
    const boldButton = document.querySelector('button[title="Bold"]')
    const undoButton = document.querySelector('button[title="Undo"]')
    const redoButton = document.querySelector('button[title="Redo"]')
    if (!(formulaInput instanceof HTMLInputElement)) {
      throw new Error('Formula input not found')
    }
    if (!(boldButton instanceof HTMLButtonElement)) {
      throw new Error('Bold button not found')
    }
    if (!(undoButton instanceof HTMLButtonElement)) {
      throw new Error('Undo button not found')
    }
    if (!(redoButton instanceof HTMLButtonElement)) {
      throw new Error('Redo button not found')
    }

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, '42')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    await act(async () => {
      boldButton.click()
    })

    await act(async () => {
      undoButton.click()
    })

    await act(async () => {
      triggerSave()
      await Promise.resolve()
    })

    const undoneSave = JSON.parse(lastSavedContent())
    const undoneCell = undoneSave.workbook.sheets[0].cells.A1
    expect(undoneCell.value).toBe('42')
    expect(undoneCell.style?.bold).toBeUndefined()

    await act(async () => {
      redoButton.click()
    })

    await act(async () => {
      triggerSave()
      await Promise.resolve()
    })

    const redoneSave = JSON.parse(lastSavedContent())
    const redoneCell = redoneSave.workbook.sheets[0].cells.A1
    expect(redoneCell.value).toBe('42')
    expect(redoneCell.style.bold).toBe(true)
  })

  it('pastes values-only and formatting-only from the in-app clipboard', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)
    writeClipboard.mockResolvedValue(undefined)
    let workbookDocument = createBlankWorkbook('Paste special')
    workbookDocument = {
      ...workbookDocument,
      workbook: {
        ...workbookDocument.workbook,
        sheets: [
          {
            ...workbookDocument.workbook.sheets[0],
            cells: {
              A1: { value: '2', kind: 'number' },
              A2: { value: '3', kind: 'number' },
              B1: {
                value: '=SUM(A1:A2)',
                kind: 'formula',
                style: { bold: true },
              },
            },
          },
        ],
      },
    }

    render(<PureSheetsShell initialDocument={workbookDocument} />)

    const cellB1 = document.querySelector('button[aria-label="B1"]')
    const cellD1 = document.querySelector('button[aria-label="D1"]')
    const copyButton = document.querySelector(
      'button[title="Copy selected range"]',
    )
    const pasteValuesButton = document.querySelector(
      'button[title="Paste values only"]',
    )
    const pasteFormattingButton = document.querySelector(
      'button[title="Paste formatting only"]',
    )
    if (!(cellB1 instanceof HTMLButtonElement)) {
      throw new Error('B1 cell not found')
    }
    if (!(cellD1 instanceof HTMLButtonElement)) {
      throw new Error('D1 cell not found')
    }
    if (!(copyButton instanceof HTMLButtonElement)) {
      throw new Error('Copy button not found')
    }
    if (!(pasteValuesButton instanceof HTMLButtonElement)) {
      throw new Error('Paste values button not found')
    }
    if (!(pasteFormattingButton instanceof HTMLButtonElement)) {
      throw new Error('Paste formatting button not found')
    }

    // Copy B1 (a bold formula), then paste values-only at D1.
    await act(async () => {
      cellB1.click()
    })
    await act(async () => {
      copyButton.click()
      await Promise.resolve()
    })
    await act(async () => {
      cellD1.click()
    })
    await act(async () => {
      pasteValuesButton.click()
    })

    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    let saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.D1.value).toBe('5')
    expect(saved.workbook.sheets[0].cells.D1.kind).toBe('number')
    expect(saved.workbook.sheets[0].cells.D1.style).toBeUndefined()

    // Paste formatting-only at E1: no value, just the bold style.
    const cellE1 = document.querySelector('button[aria-label="E1"]')
    if (!(cellE1 instanceof HTMLButtonElement)) {
      throw new Error('E1 cell not found')
    }
    await act(async () => {
      cellE1.click()
    })
    await act(async () => {
      pasteFormattingButton.click()
    })
    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.E1.value).toBe('')
    expect(saved.workbook.sheets[0].cells.E1.style).toEqual({ bold: true })
    // The original formula cell is untouched.
    expect(saved.workbook.sheets[0].cells.B1.value).toBe('=SUM(A1:A2)')
  })

  it('offers formula autocomplete, highlights references, and surfaces errors', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)

    render(
      <PureSheetsShell initialDocument={createBlankWorkbook('Formula UX')} />,
    )

    const formulaInput = document.querySelector(
      'input[aria-label="Formula or value"]',
    )
    if (!(formulaInput instanceof HTMLInputElement)) {
      throw new Error('Formula input not found')
    }

    // Autocomplete: typing "=SU" while focused suggests SUM/SUMIF.
    await act(async () => {
      formulaInput.focus()
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, '=SU')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const listbox = document.querySelector(
      '[role="listbox"][aria-label="Formula suggestions"]',
    )
    if (!listbox) throw new Error('Formula suggestion listbox not found')
    const optionLabels = Array.from(
      listbox.querySelectorAll('[role="option"]'),
    ).map(option => option.textContent ?? '')
    expect(optionLabels.some(label => label.includes('SUM(range)'))).toBe(true)
    expect(
      optionLabels.some(label => label.includes('SUMIF(range, criteria')),
    ).toBe(true)

    // Accepting a suggestion completes the function name with its paren.
    const sumOption = Array.from(
      listbox.querySelectorAll<HTMLButtonElement>('[role="option"]'),
    ).find(option => option.textContent?.includes('SUM(range)'))
    if (!sumOption) throw new Error('SUM suggestion not found')
    await act(async () => {
      sumOption.dispatchEvent(
        new MouseEvent('mousedown', { bubbles: true, cancelable: true }),
      )
    })
    expect(formulaInput.value).toBe('=SUM(')

    // Reference highlighting: each unique reference gets a colored token.
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, '=SUM(A1:B2)+C3')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const referenceTokens = Array.from(
      document.querySelectorAll('[data-formula-reference]'),
    ).map(token => token.textContent)
    expect(referenceTokens).toEqual(['A1:B2', 'C3'])

    // Error surfacing: a bad formula commits and shows its error value.
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, '=1/0')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })
    expect(document.body.textContent).toContain('#DIV/0!')
  })

  it('adjusts decimal places and applies table style presets from the toolbar', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)
    let workbookDocument = createBlankWorkbook('Formatting S2')
    workbookDocument = {
      ...workbookDocument,
      workbook: {
        ...workbookDocument.workbook,
        sheets: [
          {
            ...workbookDocument.workbook.sheets[0],
            cells: {
              A1: { value: 'Region', kind: 'text' },
              B1: { value: 'Revenue', kind: 'text' },
              A2: { value: 'North', kind: 'text' },
              B2: { value: '1200.5', kind: 'number' },
            },
          },
        ],
      },
    }

    render(<PureSheetsShell initialDocument={workbookDocument} />)

    const rangeInput = document.querySelector(
      'input[aria-label="Selected range"]',
    )
    const cellB2 = document.querySelector('button[aria-label="B2"]')
    const increaseDecimals = document.querySelector(
      'button[title="Increase decimal places"]',
    )
    if (!(rangeInput instanceof HTMLInputElement)) {
      throw new Error('Range input not found')
    }
    if (!(cellB2 instanceof HTMLButtonElement)) {
      throw new Error('B2 cell not found')
    }
    if (!(increaseDecimals instanceof HTMLButtonElement)) {
      throw new Error('Increase decimals button not found')
    }

    // Decimals: select B2 and add a decimal place (default 2 → 3).
    await act(async () => {
      cellB2.click()
    })
    await act(async () => {
      increaseDecimals.click()
    })
    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    let saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.B2.style).toEqual({
      numberFormat: 'number',
      decimals: 3,
    })
    expect(document.body.textContent).toContain('1200.500')

    // Table preset: header row over A1:B2 via the Table style menu.
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(rangeInput, 'A1:B2')
      rangeInput.dispatchEvent(new Event('input', { bubbles: true }))
      rangeInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })
    const tableMenuTrigger = document.querySelector(
      'button[aria-label="Table style"]',
    )
    if (!(tableMenuTrigger instanceof HTMLButtonElement)) {
      throw new Error('Table style menu trigger not found')
    }
    await act(async () => {
      tableMenuTrigger.click()
    })
    const headerRowItem = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[data-platform-menu-surface] [role="menuitem"]',
      ),
    ).find(item => item.textContent?.trim() === 'Header row')
    if (!headerRowItem) throw new Error('Header row menu item not found')
    await act(async () => {
      headerRowItem.click()
    })
    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A1.style).toEqual({
      bold: true,
      fillColor: '#eef1eb',
      border: true,
    })
    expect(saved.workbook.sheets[0].cells.B1.style.bold).toBe(true)
    expect(saved.workbook.sheets[0].cells.A2.style).toBeUndefined()
  })

  it('fills down with the platform fill shortcut', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)
    let workbookDocument = createBlankWorkbook('Fill shortcut')
    workbookDocument = {
      ...workbookDocument,
      workbook: {
        ...workbookDocument.workbook,
        sheets: [
          {
            ...workbookDocument.workbook.sheets[0],
            cells: {
              A1: { value: '5', kind: 'number' },
            },
          },
        ],
      },
    }

    render(<PureSheetsShell initialDocument={workbookDocument} />)

    const rangeInput = document.querySelector(
      'input[aria-label="Selected range"]',
    )
    if (!(rangeInput instanceof HTMLInputElement)) {
      throw new Error('Range input not found')
    }

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(rangeInput, 'A1:A3')
      rangeInput.dispatchEvent(new Event('input', { bubbles: true }))
      rangeInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    await act(async () => {
      window.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          cancelable: true,
          key: 'd',
          metaKey: true,
        }),
      )
    })

    expect(document.body.textContent).toContain('Filled down A1:A3')

    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    const saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A2.value).toBe('5')
    expect(saved.workbook.sheets[0].cells.A3.value).toBe('5')
  })

  it('saves cell comments from the comment bar and marks the cell', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)

    render(
      <PureSheetsShell initialDocument={createBlankWorkbook('Comments')} />,
    )

    const commentToggle = document.querySelector('button[title="Cell comment"]')
    if (!(commentToggle instanceof HTMLButtonElement)) {
      throw new Error('Cell comment button not found')
    }
    await act(async () => {
      commentToggle.click()
    })

    const commentInput = document.querySelector(
      'input[aria-label="Cell comment"]',
    )
    if (!(commentInput instanceof HTMLInputElement)) {
      throw new Error('Comment input not found')
    }
    // Type first so the field is dirty (the Save button is disabled until the
    // draft differs from the stored comment), then flush before clicking Save.
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(commentInput, 'double-check with finance')
      commentInput.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const saveNoteButton = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[aria-label="Cell comment editor"] button',
      ),
    ).find(button => button.textContent === 'Save comment')
    if (!saveNoteButton) throw new Error('Save comment button not found')

    await act(async () => {
      saveNoteButton.click()
    })

    expect(document.body.textContent).toContain('Saved comment on A1')
    const cellA1 = document.querySelector('button[aria-label="A1"]')
    expect(cellA1?.getAttribute('title')).toContain(
      'Note: double-check with finance',
    )

    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    const saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].comments).toEqual({
      A1: 'double-check with finance',
    })
  })

  it('protects ranges: locked cells refuse edits until unprotected', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)
    let workbookDocument = createBlankWorkbook('Protection')
    workbookDocument = {
      ...workbookDocument,
      workbook: {
        ...workbookDocument.workbook,
        sheets: [
          {
            ...workbookDocument.workbook.sheets[0],
            cells: {
              A1: { value: 'keep me', kind: 'text' },
            },
          },
        ],
      },
    }

    render(<PureSheetsShell initialDocument={workbookDocument} />)

    const lockButton = document.querySelector(
      'button[title="Protect selected range"]',
    )
    const unlockButton = document.querySelector(
      'button[title="Unprotect selected range"]',
    )
    const formulaInput = document.querySelector(
      'input[aria-label="Formula or value"]',
    )
    if (!(lockButton instanceof HTMLButtonElement)) {
      throw new Error('Protect button not found')
    }
    if (!(unlockButton instanceof HTMLButtonElement)) {
      throw new Error('Unprotect button not found')
    }
    if (!(formulaInput instanceof HTMLInputElement)) {
      throw new Error('Formula input not found')
    }

    // Lock A1 (default selection), then try to edit it.
    await act(async () => {
      lockButton.click()
    })
    expect(document.body.textContent).toContain('Protected A1:A1')

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, 'overwrite attempt')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })
    expect(document.body.textContent).toContain(
      'A1 is protected — unlock the range to edit it',
    )

    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    let saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A1.value).toBe('keep me')
    expect(saved.workbook.sheets[0].protection).toEqual({
      lockedRanges: ['A1:A1'],
    })

    // Unlock and the edit goes through.
    await act(async () => {
      unlockButton.click()
    })
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, 'edited freely')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })
    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A1.value).toBe('edited freely')
    expect(saved.workbook.sheets[0].protection).toBeUndefined()
  })

  it('applies validation and conditional-format rules and flags invalid entries', async () => {
    // Validation feedback is computed by the shell from the LIVE workbook
    // (invalidCellFlagsForSheet) and handed to both editing surfaces: the
    // Univer canvas paints the flagged cells, and the fallback grid (the
    // surface under vitest) titles them — the assertions below exercise that
    // shared shell-computed map, plus the commit-time status line.
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)

    render(<PureSheetsShell initialDocument={createBlankWorkbook('Rules')} />)

    // --- validation: number range 0..10 on A1:A1 ---
    const validationToggle = document.querySelector(
      'button[aria-label="Data validation"]',
    )
    if (!(validationToggle instanceof HTMLButtonElement)) {
      throw new Error('Data validation button not found')
    }
    await act(async () => {
      validationToggle.click()
    })
    const numberRangeOption = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[aria-label="Data validation rule"] [role="radio"]',
      ),
    ).find(option => option.textContent === 'Number range')
    if (!numberRangeOption) throw new Error('Number range option not found')
    await act(async () => {
      numberRangeOption.click()
    })
    const maxInput = document.querySelector('input[aria-label="Maximum value"]')
    if (!(maxInput instanceof HTMLInputElement)) {
      throw new Error('Maximum value input not found')
    }
    const applyValidationButton = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[aria-label="Data validation rule"] button',
      ),
    ).find(button => button.textContent === 'Apply rule')
    if (!applyValidationButton) throw new Error('Apply rule button not found')
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(maxInput, '10')
      maxInput.dispatchEvent(new Event('input', { bubbles: true }))
      applyValidationButton.click()
    })
    expect(document.body.textContent).toContain(
      'Added number-range validation for A1:A1',
    )

    // Committing an out-of-range value is FLAGGED but not rejected.
    const formulaInput = document.querySelector(
      'input[aria-label="Formula or value"]',
    )
    if (!(formulaInput instanceof HTMLInputElement)) {
      throw new Error('Formula input not found')
    }
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(formulaInput, '50')
      formulaInput.dispatchEvent(new Event('input', { bubbles: true }))
      formulaInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })
    expect(document.body.textContent).toContain(
      'A1 flagged: Value must be at most 10',
    )
    const cellA1 = document.querySelector('button[aria-label="A1"]')
    expect(cellA1?.getAttribute('title')).toContain('Invalid:')

    // --- conditional formatting: greater than 10 → preset style ---
    const conditionalToggle = document.querySelector(
      'button[aria-label="Conditional formatting"]',
    )
    if (!(conditionalToggle instanceof HTMLButtonElement)) {
      throw new Error('Conditional formatting button not found')
    }
    await act(async () => {
      conditionalToggle.click()
    })
    const conditionValueInput = document.querySelector(
      'input[aria-label="Condition value"]',
    )
    if (!(conditionValueInput instanceof HTMLInputElement)) {
      throw new Error('Condition value input not found')
    }
    const applyConditionalButton = Array.from(
      document.querySelectorAll<HTMLButtonElement>(
        '[aria-label="Conditional formatting rule"] button',
      ),
    ).find(button => button.textContent === 'Apply rule')
    if (!applyConditionalButton) {
      throw new Error('Apply conditional rule button not found')
    }
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(conditionValueInput, '10')
      conditionValueInput.dispatchEvent(new Event('input', { bubbles: true }))
      applyConditionalButton.click()
    })
    expect(document.body.textContent).toContain(
      'Added conditional format for A1:A1',
    )

    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    const saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A1.value).toBe('50')
    expect(saved.workbook.sheets[0].validations).toHaveLength(1)
    expect(saved.workbook.sheets[0].validations[0]).toMatchObject({
      range: 'A1:A1',
      kind: 'number-range',
      max: 10,
    })
    expect(saved.workbook.sheets[0].conditionalFormats).toHaveLength(1)
    expect(saved.workbook.sheets[0].conditionalFormats[0]).toMatchObject({
      range: 'A1:A1',
      condition: { kind: 'greater', value: '10' },
      style: { textColor: '#982d25', bold: true },
    })
  })

  it('shows drawer-agent activity from the workbook agent log', async () => {
    let workbookDocument = createBlankWorkbook('Agent log')
    workbookDocument = {
      ...workbookDocument,
      workbook: {
        ...workbookDocument.workbook,
        agentLog: [
          {
            id: 'agent-log-1',
            at: '2026-08-30T09:30:00.000Z',
            agentName: 'PureSheets Assistant',
            tool: 'applyCellChanges',
            summary: 'Filled the totals row from Q3 data.',
            sheetId: workbookDocument.workbook.sheets[0].id,
            sheetName: 'Sheet 1',
            range: 'A9:D9',
            cells: [{ cell: 'A9', after: 'Total', kind: 'text' as const }],
          },
        ],
      },
    }
    render(<PureSheetsShell initialDocument={workbookDocument} />)

    const activityToggle = document.querySelector(
      'button[title="Agent activity (1) — click to view"]',
    )
    if (!(activityToggle instanceof HTMLButtonElement)) {
      throw new Error('Agent activity button not found')
    }
    expect(document.body.textContent).not.toContain(
      'Filled the totals row from Q3 data.',
    )

    await act(async () => {
      activityToggle.click()
    })

    expect(document.body.textContent).toContain('Agent activity')
    expect(document.body.textContent).toContain(
      'Filled the totals row from Q3 data.',
    )
    expect(document.body.textContent).toContain('PureSheets Assistant')
    expect(document.body.textContent).toContain('A9:D9')

    const closeButton = Array.from(
      document.querySelectorAll<HTMLButtonElement>('button'),
    ).find(button => button.textContent === 'Close')
    if (!closeButton) throw new Error('Close button not found')
    await act(async () => {
      closeButton.click()
    })
    expect(document.body.textContent).not.toContain(
      'Filled the totals row from Q3 data.',
    )
  })

  it('replaces across the whole sheet honoring match case', async () => {
    writeTextFile.mockResolvedValue(undefined)
    updatePureSheetsSettings.mockResolvedValue(undefined)
    let workbookDocument = createBlankWorkbook('Replace all')
    workbookDocument = {
      ...workbookDocument,
      workbook: {
        ...workbookDocument.workbook,
        sheets: [
          {
            ...workbookDocument.workbook.sheets[0],
            cells: {
              A1: { value: 'North', kind: 'text' },
              A2: { value: 'NORTH', kind: 'text' },
              B1: { value: 'north wind', kind: 'text' },
            },
          },
        ],
      },
    }

    render(<PureSheetsShell initialDocument={workbookDocument} />)

    const findInput = document.querySelector('input[aria-label="Find cells"]')
    const replaceInput = document.querySelector(
      'input[aria-label="Replace with"]',
    )
    const matchCaseToggle = document.querySelector(
      'button[title="Match case: off"]',
    )
    const replaceAllButton = document.querySelector(
      'button[title="Replace all in sheet"]',
    )
    if (!(findInput instanceof HTMLInputElement)) {
      throw new Error('Find input not found')
    }
    if (!(replaceInput instanceof HTMLInputElement)) {
      throw new Error('Replace input not found')
    }
    if (!(matchCaseToggle instanceof HTMLButtonElement)) {
      throw new Error('Match case toggle not found')
    }
    if (!(replaceAllButton instanceof HTMLButtonElement)) {
      throw new Error('Replace all button not found')
    }

    // Turn match case ON, replace "North" → "West": only the exact-case
    // occurrences change.
    await act(async () => {
      matchCaseToggle.click()
    })
    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(findInput, 'North')
      findInput.dispatchEvent(new Event('input', { bubbles: true }))
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(replaceInput, 'West')
      replaceInput.dispatchEvent(new Event('input', { bubbles: true }))
      replaceAllButton.click()
    })
    expect(document.body.textContent).toContain('Replaced 1 match in Sheet 1')

    await act(async () => {
      triggerSave()
      await flushAsync()
    })
    const saved = JSON.parse(lastSavedContent())
    expect(saved.workbook.sheets[0].cells.A1.value).toBe('West')
    expect(saved.workbook.sheets[0].cells.A2.value).toBe('NORTH')
    expect(saved.workbook.sheets[0].cells.B1.value).toBe('north wind')
  })

  it('shows range summaries (sum, average, min, max, count) in the status bar', async () => {
    let workbookDocument = createBlankWorkbook('Summaries')
    workbookDocument = {
      ...workbookDocument,
      workbook: {
        ...workbookDocument.workbook,
        sheets: [
          {
            ...workbookDocument.workbook.sheets[0],
            cells: {
              A1: { value: '10', kind: 'number' },
              A2: { value: '20', kind: 'number' },
              A3: { value: '=A1+A2', kind: 'formula' },
              A4: { value: 'label', kind: 'text' },
            },
          },
        ],
      },
    }

    render(<PureSheetsShell initialDocument={workbookDocument} />)

    const rangeInput = document.querySelector(
      'input[aria-label="Selected range"]',
    )
    if (!(rangeInput instanceof HTMLInputElement)) {
      throw new Error('Range input not found')
    }

    // Single cell: no summary shown.
    expect(document.body.textContent).not.toContain('Sum ')

    await act(async () => {
      Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        'value',
      )?.set?.call(rangeInput, 'A1:A4')
      rangeInput.dispatchEvent(new Event('input', { bubbles: true }))
      rangeInput.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
    })

    // Formula A3 computes to 30; the text cell is ignored by the numerics.
    expect(document.body.textContent).toContain(
      'Sum 60 · Avg 20 · Min 10 · Max 30 · Count 3',
    )
  })
})
