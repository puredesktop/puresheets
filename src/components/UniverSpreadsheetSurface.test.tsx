// @vitest-environment happy-dom
import { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { UniverSpreadsheetSurface } from './UniverSpreadsheetSurface'
import type { UniverEditorBridge } from './univerBridgeTypes'
import {
  createUniverSnapshot,
  type PureSheetsUniverSnapshot,
} from '../lib/univerAdapter'
import { createBlankWorkbook } from '../lib/workbookModel'

const runtime = vi.hoisted(() => ({
  snapshot: null as PureSheetsUniverSnapshot | null,
  calculation: null as Promise<void> | null,
}))
const worksheet = {
  getSheetId: () => 'sheet-1',
  activate: vi.fn(),
  getRange: () => ({ setValues: vi.fn() }),
}
const api = {
  getActiveWorkbook: () => ({
    save: () => structuredClone(runtime.snapshot),
    getActiveSheet: () => worksheet,
    insertSheet: () => worksheet,
    getSheetBySheetId: () => worksheet,
  }),
  disposeUnit: vi.fn(),
  getFormula: () => ({
    onCalculationResultApplied: () => {
      const calculation = runtime.calculation
      runtime.calculation = null
      return calculation ?? Promise.resolve()
    },
  }),
}
vi.mock('@univerjs/core', async importOriginal => ({
  ...(await importOriginal<typeof import('@univerjs/core')>()),
  LocaleType: { EN_US: 'en-US' },
  ThemeService: class {},
  UniverInstanceType: { UNIVER_SHEET: 2 },
  Univer: class {
    registerPlugins() {}
    createUnit(_type: unknown, snapshot: PureSheetsUniverSnapshot) {
      runtime.snapshot = snapshot
    }
    dispose() {}
  },
}))
vi.mock('@univerjs/core/facade', () => ({ FUniver: { newAPI: () => api } }))
vi.mock('@univerjs/preset-sheets-core', () => ({
  UniverSheetsCorePreset: () => ({ plugins: [] }),
}))
vi.mock('@univerjs/preset-sheets-conditional-formatting', () => ({
  UniverSheetsConditionalFormattingPreset: () => ({ plugins: [] }),
}))
vi.mock('@univerjs/preset-sheets-core/locales/en-US', () => ({ default: {} }))
vi.mock('@univerjs/preset-sheets-conditional-formatting/locales/en-US', () => ({
  default: {},
}))
let root: Root
let editor: UniverEditorBridge | null
const change = vi.fn(),
  selection = vi.fn()
const ready = (value: UniverEditorBridge | null) => {
  editor = value
}
const first = createUniverSnapshot(createBlankWorkbook('Original'))
const next = createUniverSnapshot(createBlankWorkbook('Replacement'))
function render(snapshot = first, revision = 0) {
  act(() =>
    root.render(
      <UniverSpreadsheetSurface
        snapshot={snapshot}
        revision={revision}
        onEditorReady={ready}
        onSnapshotChange={change}
        onSelectionChange={selection}
      />,
    ),
  )
}
function delayedCalculation() {
  let resolve!: () => void
  runtime.calculation = new Promise<void>(done => {
    resolve = done
  })
  return resolve
}
beforeEach(async () => {
  vi.stubEnv('MODE', 'development') // Exercise the real surface lifecycle with a disposable facade.
  const host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  editor = null
  change.mockClear()
  selection.mockClear()
  runtime.calculation = null
  ;(
    globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }
  ).IS_REACT_ACT_ENVIRONMENT = true
  render()
  await act(async () => {
    await vi.waitFor(() => expect(editor).not.toBeNull())
  })
})
afterEach(() => {
  act(() => root.unmount())
  document.body.innerHTML = ''
  vi.unstubAllEnvs()
})
it('discards a snapshot flush whose calculation finishes after a unit reload', async () => {
  const finish = delayedCalculation(),
    operation = editor!.flushSnapshot()
  render({ ...next, id: first.id }, 1)
  await act(async () => {
    finish()
    expect(await operation).toBeNull()
  })
  expect(change).not.toHaveBeenCalled()
  await act(async () => {
    expect(await editor!.flushSnapshot()).not.toBeNull()
  })
  expect(change).toHaveBeenCalledTimes(1)
})
it('does not export an old workbook bridge edit over a replacement', async () => {
  const finish = delayedCalculation(),
    operation = editor!.addSheet('Pending sheet')
  render({ ...next, id: first.id }, 1)
  await act(async () => {
    finish()
    expect(await operation).toBe(false)
  })
  expect(change).not.toHaveBeenCalled()
  expect(selection).not.toHaveBeenCalled()
})
it('drops a range-edit snapshot after the editor is unmounted', async () => {
  const finish = delayedCalculation(),
    operation = editor!.setRangeValues('A1', [[42]])
  act(() => root.unmount())
  await act(async () => {
    finish()
    expect(await operation).toBe(false)
  })
  expect(change).not.toHaveBeenCalled()
  expect(selection).not.toHaveBeenCalled()
})
