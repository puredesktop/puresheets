import { useEffect, useRef } from 'react'
import type React from 'react'
import styled from 'styled-components'
import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import { CELL_BORDER_COLOR } from '../constants'
import { cellKey, parseCellKey } from '../lib/workbookModel'
import { normalizeConditionalFormats } from '../lib/sheetMetadataNormalize'
import { requestedSheetId } from '../lib/workbookView'
import type { InvalidCellFlag } from '../lib/sheetDataFeatures'
import type { PureSheetsUniverSnapshot } from '../lib/univerAdapter'
import type { CellStyle, ConditionalFormatRule } from '../types'
import type {
  UniverConditionalFormatBuilder,
  UniverEditorBridge,
  UniverFacadeWorkbook,
  UniverFacadeWorksheet,
  UniverPasteValue,
  UniverRangeLike,
  UniverSelectionState,
} from './univerBridgeTypes'

// Per-sheet snapshot metadata that the React model — not Univer — owns (Phase
// S3: cell comments, validations, protection, conditional formats, and the
// filter/sort state). Univer round-trips whatever it was loaded with but never
// edits these, so its copy goes stale the instant the shell changes one without
// reloading the surface. `documentFromUniverSnapshot` treats a *present* key as
// authoritative and a *missing* key as "fall back to the model", so if the
// stale Univer echo still carries these keys it silently reverts them — most
// visibly a just-added cell comment vanishing on the next edit or save (#273).
const REACT_OWNED_SHEET_METADATA = [
  'filterRow',
  'filterColumn',
  'filterQuery',
  'comments',
  'validations',
  'protection',
  'conditionalFormats',
] as const

// Drop the React-owned keys from every exported snapshot so the re-import falls
// back to the live model for them instead of the stale Univer copy (#273).
const stripReactOwnedSheetMetadata = (
  snapshot: PureSheetsUniverSnapshot,
): PureSheetsUniverSnapshot => {
  const sheets = (snapshot as { sheets?: Record<string, unknown> }).sheets
  if (!sheets || typeof sheets !== 'object') return snapshot
  const nextSheets: Record<string, unknown> = {}
  for (const [sheetId, sheet] of Object.entries(sheets)) {
    const sheetRecord = sheet as Record<string, unknown>
    const custom = sheetRecord.custom as Record<string, unknown> | undefined
    const puresheets = custom?.puresheets as Record<string, unknown> | undefined
    if (!puresheets) {
      nextSheets[sheetId] = sheet
      continue
    }
    const nextPuresheets: Record<string, unknown> = { ...puresheets }
    for (const key of REACT_OWNED_SHEET_METADATA) delete nextPuresheets[key]
    nextSheets[sheetId] = {
      ...sheetRecord,
      custom: { ...custom, puresheets: nextPuresheets },
    }
  }
  return { ...snapshot, sheets: nextSheets } as PureSheetsUniverSnapshot
}

// Validation-flag painting (Phase S3): the indicator style and the per-sheet
// cap on generated display rules.
const INVALID_CELL_FILL = '#fdecea'
const INVALID_CELL_TEXT = '#b3261e'
const MAX_VALIDATION_FLAG_RULES = 100

// Univer mutates the snapshot object it is handed by `createUnit` IN PLACE as
// the workbook is edited. The snapshots we pass come from a document's cached
// `engine.snapshot`, which also lives in the undo/redo history — so without a
// defensive copy, a structural edit (insert/delete row/column) after a reload
// silently rewrites the *historical* snapshot stored in an undo entry, and
// undoing that edit reloads the already-corrupted state (#271).
const cloneSnapshot = <T,>(snapshot: T): T =>
  typeof structuredClone === 'function'
    ? structuredClone(snapshot)
    : (JSON.parse(JSON.stringify(snapshot)) as T)

/**
 * UniverSpreadsheetSurface — mounts the Univer sheets editor (via
 * `@univerjs/preset-sheets-core`) inside the PureSheets shell and exposes the
 * `UniverEditorBridge` for shell-level commands.
 *
 * Ownership boundary (Phase S0, MAIL_SHEETS_MASTER_PLAN.md):
 * - Univer owns interactive editing, rendering, and live formula evaluation.
 * - PureDesktop keeps file routing, `.sheets`/`.sheets.html` persistence,
 *   agent context, and app-level actions. The Univer snapshot is the editing
 *   bridge: workbook model → snapshot on mount, snapshot → `.sheets` document
 *   on every edit/save (`onSnapshotChange`).
 *
 * In vitest (`import.meta.env.MODE === 'test'`) the editor does not mount —
 * the shell falls back to the legacy compatibility grid instead.
 */
export function UniverSpreadsheetSurface({
  onEditorReady,
  onSelectionChange,
  snapshot,
  revision,
  onSnapshotChange,
  onEditorCopy,
  invalidCells,
}: {
  onEditorReady: (bridge: UniverEditorBridge | null) => void
  onSelectionChange: (selection: UniverSelectionState) => void
  snapshot: PureSheetsUniverSnapshot | undefined
  revision: number
  onSnapshotChange: (
    snapshot: PureSheetsUniverSnapshot,
    opts?: { record?: boolean },
  ) => void
  /** Fired when the user copies/cuts on the canvas (Ctrl+C / Ctrl+X) so the
   *  shell can mirror the selection into its internal clipboard. */
  onEditorCopy?: () => void
  /** Data-validation feedback (Phase S3), keyed by sheet id: the cells whose
   *  current value fails a validation rule. Painted on the canvas as
   *  display-only conditional-format rules — a red tint, never a rejection. */
  invalidCells?: Record<string, InvalidCellFlag[]>
}): React.ReactElement {
  const hostRef = useRef<HTMLDivElement | null>(null)
  const lastSnapshotRef = useRef(snapshot)
  const invalidCellsRef = useRef<Record<string, InvalidCellFlag[]>>({})
  invalidCellsRef.current = invalidCells ?? {}
  // Repaints the validation flags on the live canvas. Populated at mount.
  const applyValidationFlagsRef = useRef<(() => void) | null>(null)
  const invalidCellsSignatureRef = useRef('')
  // Kept in a ref so a changing copy handler doesn't remount the whole surface.
  const onEditorCopyRef = useRef(onEditorCopy)
  useEffect(() => {
    onEditorCopyRef.current = onEditorCopy
  }, [onEditorCopy])
  // Reloads the workbook data into the *existing* Univer instance without a
  // full teardown/rebuild. Populated once the editor mounts; a `revision` bump
  // drives it instead of remounting the whole surface (#269).
  const reloadUnitRef = useRef<(() => void) | null>(null)

  useEffect(() => {
    lastSnapshotRef.current = snapshot
  }, [snapshot])

  useEffect(() => {
    const host = hostRef.current
    if (!host || !snapshot || import.meta.env.MODE === 'test') return
    let disposed = false
    let disposeUniver: (() => void) | undefined
    let snapshotTimeout: number | undefined
    // A programmatic reload replays the whole workbook into the live Univer
    // instance, which echoes a burst of commands back through onCommandExecuted.
    // While `suppressExport` is set those echoes are absorbed (never exported to
    // the shell); it stays set until the command stream goes quiet, so a delayed
    // async echo (e.g. formula recalc) can't land as a spurious user edit and
    // corrupt the shell's undo/redo stacks (#269).
    let suppressExport = false
    let suppressDrainTimeout: number | undefined
    // Bumped every time a reload takes ownership of the suppression window, so a
    // stale recalc-settle callback from an earlier reload can't lift suppression
    // out from under a newer one (#271).
    let suppressToken = 0
    let readyForChanges = false
    let readyTimeout: number | undefined
    let preserveNextCommandRangeSync = false
    // Set while a bridge method is committing its own edit. Bridge edits export
    // one authoritative snapshot explicitly; without this, the `onCommandExecuted`
    // handler ALSO exported the same edit (debounced), so a single toolbar action
    // (insert row, apply format, …) landed on the undo stack twice — needing two
    // Undo presses / making formatting look un-undoable (#271).
    let suppressCommandExport = false
    // A direct (in-canvas) edit spans its first content mutation through to the
    // moment its formula recalc settles. `editOpen` holds across that span so
    // the recalc write-back folds into the SAME undo step instead of starting a
    // new one; `recordedThisEdit` ensures exactly the first export of the span
    // pushes an undo entry (the rest are content-only updates). `editToken`
    // guards the async close against a newer mutation superseding it (#271).
    let editOpen = false
    let recordedThisEdit = false
    let editToken = 0

    async function mountUniver(): Promise<void> {
      const [
        { LocaleType, Univer, UniverInstanceType },
        { UniverSheetsCorePreset },
        { UniverSheetsConditionalFormattingPreset },
        { FUniver },
        enUS,
        conditionalFormattingEnUS,
      ] = await Promise.all([
        import('@univerjs/core'),
        import('@univerjs/preset-sheets-core'),
        import('@univerjs/preset-sheets-conditional-formatting'),
        import('@univerjs/core/facade'),
        import('@univerjs/preset-sheets-core/locales/en-US'),
        import('@univerjs/preset-sheets-conditional-formatting/locales/en-US'),
      ])
      if (disposed || !hostRef.current || !lastSnapshotRef.current) return

      const univer = new Univer({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: {
            ...enUS.default,
            ...conditionalFormattingEnUS.default,
          },
        },
      })
      const preset = UniverSheetsCorePreset({
        container: hostRef.current,
        header: false,
        toolbar: false,
        formulaBar: false,
        footer: false,
        contextMenu: true,
        disableAutoFocus: true,
      })
      // Conditional-formatting rules live in the PureSheets model
      // (sheet.conditionalFormats) and are seeded into Univer on every unit
      // load — this preset makes Univer evaluate and PAINT them live (#270).
      const conditionalFormattingPreset =
        UniverSheetsConditionalFormattingPreset()
      univer.registerPlugins(
        [...preset.plugins, ...conditionalFormattingPreset.plugins].map(
          plugin => (Array.isArray(plugin) ? plugin : [plugin]),
        ) as Parameters<typeof univer.registerPlugins>[0],
      )
      univer.createUnit(
        UniverInstanceType.UNIVER_SHEET,
        cloneSnapshot(lastSnapshotRef.current),
      )

      // Univer relayouts off the global `resize` event. Firing it on EVERY
      // ResizeObserver callback created a feedback loop: dispatch → Univer
      // relayout → scrollbar show/hide changes the host's content box →
      // observer fires again → … a self-sustaining repaint that pegged the CPU
      // once the grid was tall enough to scroll (#272). Only re-dispatch when
      // the host actually changed size, and coalesce to one dispatch per frame
      // so a burst of sub-pixel churn can't spin.
      let lastResizeWidth = 0
      let lastResizeHeight = 0
      let resizeRaf = 0
      const requestResize = (): void => {
        if (resizeRaf) return
        resizeRaf = window.requestAnimationFrame(() => {
          resizeRaf = 0
          window.dispatchEvent(new Event('resize'))
        })
      }
      requestResize()
      const resizeObserver = new ResizeObserver(entries => {
        const rect = entries[0]?.contentRect
        if (!rect) return
        const width = Math.round(rect.width)
        const height = Math.round(rect.height)
        if (width === lastResizeWidth && height === lastResizeHeight) return
        lastResizeWidth = width
        lastResizeHeight = height
        requestResize()
      })
      if (hostRef.current) resizeObserver.observe(hostRef.current)

      const univerApi = FUniver.newAPI(univer) as unknown as {
        Enum?: {
          BorderStyleTypes?: Record<string, unknown>
          BorderType?: Record<string, unknown>
        }
        Event?: {
          SelectionChanged?: string
          SheetEditEnded?: string
        }
        addEvent?: (
          event: string,
          callback: (params: {
            selections?: Array<{
              startRow: number
              startColumn: number
              endRow: number
              endColumn: number
            }>
            workbook?: {
              getActiveSheet?: () => UniverFacadeWorksheet | null
            }
            worksheet?: UniverFacadeWorksheet
          }) => void,
        ) => { dispose: () => void }
        getFormula?: () => {
          calculationResultApplied?: (callback: () => void) => {
            dispose: () => void
          }
          onCalculationResultApplied?: (timeout?: number) => Promise<void>
        }
        onCommandExecuted?: (
          callback: (command?: { id?: string; type?: number }) => void,
        ) => { dispose: () => void }
        getActiveWorkbook?: () => UniverFacadeWorkbook | null
        disposeUnit?: (unitId: string) => boolean
      }

      // Best-effort canvas enforcement of protected ranges (Phase S5):
      // mirror the model-layer locks into Univer's permission service where
      // the facade exposes it, so typing directly into a protected cell is
      // refused in the canvas too. Every call is optional-chained and the
      // whole block is inert on failure — the MODEL guards remain the
      // authoritative enforcement (`isCellLocked` in every mutation path).
      const applyCanvasPermissions = (): void => {
        try {
          const permissionApi = (
            univerApi as unknown as {
              getPermission?: () => {
                addRangeBaseProtection?: (
                  unitId: string,
                  subUnitId: string,
                  ranges: UniverRangeLike[],
                ) => unknown
                setPermissionDialogVisible?: (visible: boolean) => void
              }
            }
          ).getPermission?.()
          const snapshotForLocks = lastSnapshotRef.current
          if (permissionApi?.addRangeBaseProtection && snapshotForLocks) {
            permissionApi.setPermissionDialogVisible?.(false)
            for (const [sheetId, worksheetSnapshot] of Object.entries(
              snapshotForLocks.sheets ?? {},
            )) {
              const custom = (
                worksheetSnapshot as {
                  custom?: {
                    puresheets?: { protection?: { lockedRanges?: unknown } }
                  }
                }
              ).custom?.puresheets
              const lockedRanges = Array.isArray(
                custom?.protection?.lockedRanges,
              )
                ? (custom.protection.lockedRanges as string[])
                : []
              const ranges = lockedRanges.flatMap(token => {
                const [startKey, endKey] = token.split(':')
                const start = parseCellKey(startKey ?? '')
                const end = parseCellKey(endKey ?? startKey ?? '')
                if (!start || !end) return []
                return [
                  {
                    startRow: start.row,
                    startColumn: start.column,
                    endRow: end.row,
                    endColumn: end.column,
                  },
                ]
              })
              if (ranges.length) {
                permissionApi.addRangeBaseProtection(
                  snapshotForLocks.id,
                  sheetId,
                  ranges,
                )
              }
            }
          }
        } catch {
          // Canvas protection is best effort; the model layer enforces locks.
        }
      }
      applyCanvasPermissions()

      // Seed the model's conditional-format rules into Univer's CF engine so
      // the canvas evaluates and paints them live (#270). Rules are owned by
      // the PureSheets model (sheet.conditionalFormats, carried in the
      // snapshot's custom metadata); Univer is a pure renderer here, so every
      // unit load clears whatever the engine holds and reseeds from scratch.
      // Best-effort like applyCanvasPermissions — inert when the CF facade is
      // unavailable (the CompatibilityGrid path still renders rules in tests).
      const applyConditionalFormats = (): void => {
        try {
          const snapshotForRules = lastSnapshotRef.current
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          if (!snapshotForRules || !workbook) return
          for (const [sheetId, worksheetSnapshot] of Object.entries(
            snapshotForRules.sheets ?? {},
          )) {
            const worksheet = workbook.getSheetBySheetId?.(sheetId)
            if (!worksheet?.addConditionalFormattingRule) continue
            for (const existing of worksheet.getConditionalFormattingRules?.() ??
              []) {
              if (existing.cfId) {
                worksheet.deleteConditionalFormattingRule?.(existing.cfId)
              }
            }
            const custom = (
              worksheetSnapshot as {
                custom?: { puresheets?: { conditionalFormats?: unknown } }
              }
            ).custom?.puresheets
            const rules =
              normalizeConditionalFormats(custom?.conditionalFormats) ?? []
            for (const rule of rules) {
              const built = buildUniverConditionalFormat(
                worksheet.newConditionalFormattingRule?.(),
                rule,
              )
              if (built) worksheet.addConditionalFormattingRule(built)
            }
          }
        } catch {
          // CF painting is cosmetic; the rules stay intact in the model.
        }
      }
      applyConditionalFormats()

      // Paint data-validation feedback (Phase S3) on the live canvas: each
      // currently invalid cell gets a display-only conditional-format rule
      // matching its exact current value — a red tint plus red text, never a
      // rejection. Exact-match rules self-invalidate the moment the value
      // changes; the shell recomputes the flag set from the live workbook on
      // every model update and this repaints. Best-effort like the blocks
      // above — the model flags (and the status line) remain authoritative.
      const validationCfIdsBySheet = new Map<string, string[]>()
      const applyValidationFlags = (): void => {
        try {
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          const snapshotForSheets = lastSnapshotRef.current
          if (!workbook || !snapshotForSheets) return
          for (const sheetId of Object.keys(snapshotForSheets.sheets ?? {})) {
            const worksheet = workbook.getSheetBySheetId?.(sheetId)
            if (!worksheet?.addConditionalFormattingRule) continue
            // Drop the previous pass's flag rules (a unit reload may already
            // have cleared them; deleting a missing id is a no-op).
            for (const cfId of validationCfIdsBySheet.get(sheetId) ?? []) {
              try {
                worksheet.deleteConditionalFormattingRule?.(cfId)
              } catch {
                // already gone
              }
            }
            validationCfIdsBySheet.set(sheetId, [])
            const flags = invalidCellsRef.current[sheetId] ?? []
            if (!flags.length) continue
            // One exact-match rule per distinct invalid value, scoped to
            // exactly the flagged cells. Bounded so a sheet-wide rule over
            // thousands of bad cells cannot flood the CF engine.
            const cellsByValue = new Map<string, UniverRangeLike[]>()
            for (const flag of flags) {
              const position = parseCellKey(flag.cell)
              if (!position) continue
              const ranges = cellsByValue.get(flag.value) ?? []
              ranges.push({
                startRow: position.row,
                startColumn: position.column,
                endRow: position.row,
                endColumn: position.column,
              })
              cellsByValue.set(flag.value, ranges)
            }
            const before = new Set(
              (worksheet.getConditionalFormattingRules?.() ?? [])
                .map(rule => rule.cfId)
                .filter((id): id is string => Boolean(id)),
            )
            let ruleCount = 0
            for (const [value, ranges] of cellsByValue) {
              if (ruleCount >= MAX_VALIDATION_FLAG_RULES) break
              const builder = worksheet.newConditionalFormattingRule?.()
              if (!builder) break
              const numeric = Number(value)
              const conditioned =
                value.trim() !== '' && Number.isFinite(numeric)
                  ? builder.whenNumberEqualTo?.(numeric)
                  : builder.whenTextEqualTo?.(value)
              if (!conditioned) continue
              let styled = conditioned
              styled = styled.setBackground?.(INVALID_CELL_FILL) ?? styled
              styled = styled.setFontColor?.(INVALID_CELL_TEXT) ?? styled
              const built = styled.setRanges?.(ranges)?.build?.()
              if (built) {
                worksheet.addConditionalFormattingRule(built)
                ruleCount += 1
              }
            }
            validationCfIdsBySheet.set(
              sheetId,
              (worksheet.getConditionalFormattingRules?.() ?? [])
                .map(rule => rule.cfId)
                .filter(
                  (id): id is string => id !== undefined && !before.has(id),
                ),
            )
          }
        } catch {
          // Painting is cosmetic; the flags stay intact in the model.
        }
      }
      applyValidationFlags()
      applyValidationFlagsRef.current = applyValidationFlags

      const snapshotWithActiveSheet = (
        snapshot: PureSheetsUniverSnapshot,
        workbook: UniverFacadeWorkbook,
      ): PureSheetsUniverSnapshot => {
        // Strip the React-owned per-sheet metadata first so the shell's model
        // stays authoritative for it (comments etc.) on re-import (#273).
        const cleaned = stripReactOwnedSheetMetadata(snapshot)
        const activeSheetId = workbook.getActiveSheet?.()?.getSheetId?.()
        if (!activeSheetId) return cleaned
        const custom = (cleaned.custom ?? {}) as Record<string, unknown>
        const puresheets = (
          typeof custom.puresheets === 'object' && custom.puresheets !== null
            ? custom.puresheets
            : {}
        ) as Record<string, unknown>
        return {
          ...cleaned,
          custom: {
            ...custom,
            puresheets: {
              ...puresheets,
              activeSheetId,
            },
          },
        }
      }
      const saveFromUniver = (): PureSheetsUniverSnapshot | null => {
        const workbook = univerApi.getActiveWorkbook?.() ?? null
        const snapshot = workbook?.save() ?? null
        return snapshot && workbook
          ? snapshotWithActiveSheet(snapshot, workbook)
          : snapshot
      }
      // Cancel the debounced export the triggering command queued and absorb
      // any follow-up echoes (recalc, selection) for the duration of this
      // bridge edit, so exactly ONE snapshot — the one exported here — reaches
      // the shell and the edit is a single undo step (#271).
      const beginBridgeExport = (): void => {
        if (snapshotTimeout) window.clearTimeout(snapshotTimeout)
        // A bridge edit's own mutation fired through onCommandExecuted just
        // before this (it runs before suppressCommandExport is set), which
        // spuriously OPENED a direct-edit span. Cancel that span here — the
        // bridge owns this edit and records it via its explicit export — so
        // editOpen isn't left stuck true, which would fold the NEXT real edit
        // into a dead span and break its undo (#271).
        editOpen = false
        recordedThisEdit = false
        editToken += 1
        suppressCommandExport = true
      }
      const exportSnapshotAfterBridgeEdit = async (
        worksheet: UniverFacadeWorksheet,
      ): Promise<boolean> => {
        beginBridgeExport()
        try {
          await univerApi
            .getFormula?.()
            .onCalculationResultApplied?.(1000)
            .catch(() => undefined)
          const nextSnapshot = saveFromUniver()
          if (nextSnapshot) onSnapshotChange(nextSnapshot)
          syncSelectionFromUniver(worksheet, undefined, true)
          return Boolean(nextSnapshot)
        } finally {
          suppressCommandExport = false
        }
      }
      const exportWorkbookAfterBridgeEdit = async (
        workbook: UniverFacadeWorkbook,
        preserveRange = true,
        // Pure metadata/navigation edits (rename, sheet switch) change no cell
        // values, so there is nothing to recalculate — skip the up-to-1s
        // `onCalculationResultApplied` wait that otherwise makes every such op
        // feel like the app froze (#273 follow-up).
        waitForRecalc = true,
      ): Promise<boolean> => {
        beginBridgeExport()
        try {
          if (waitForRecalc) {
            await univerApi
              .getFormula?.()
              .onCalculationResultApplied?.(1000)
              .catch(() => undefined)
          }
          const nextSnapshot = snapshotWithActiveSheet(
            workbook.save(),
            workbook,
          )
          if (nextSnapshot) onSnapshotChange(nextSnapshot)
          syncSelectionFromUniver(
            workbook.getActiveSheet?.() ?? null,
            undefined,
            preserveRange,
          )
          return Boolean(nextSnapshot)
        } finally {
          suppressCommandExport = false
        }
      }
      const editActiveWorksheet = async (
        edit: (worksheet: UniverFacadeWorksheet) => boolean,
      ): Promise<boolean> => {
        if (disposed) return false
        const worksheet =
          univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
        preserveNextCommandRangeSync = true
        if (!worksheet || !edit(worksheet)) return false
        return exportSnapshotAfterBridgeEdit(worksheet)
      }
      const parsePasteMatrix = (text: string): UniverPasteValue[][] =>
        text
          .split(/\r?\n/)
          .filter(
            (rowText, index, rows) =>
              rowText.length > 0 || index < rows.length - 1,
          )
          .map(rowText =>
            rowText
              .split('\t')
              .map(value =>
                value.trim().startsWith('=') ? { f: value } : value,
              ),
          )
      const syncSelectionFromUniver = (
        worksheet = univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null,
        selections?: Array<{
          startRow: number
          startColumn: number
          endRow: number
          endColumn: number
        }>,
        preserveRange = false,
      ): void => {
        if (!worksheet) return
        const activeSelection = worksheet.getSelection?.()
        const currentCell = activeSelection?.getCurrentCell?.()
        let selectedRangeFromUniver: UniverRangeLike | undefined
        try {
          selectedRangeFromUniver = activeSelection
            ?.getActiveRange?.()
            ?.getRange?.()
        } catch {
          selectedRangeFromUniver = undefined
        }
        const activeRange = selections?.[0] ?? selectedRangeFromUniver
        const row =
          currentCell?.actualRow ??
          currentCell?.row ??
          activeRange?.startRow ??
          0
        const column =
          currentCell?.actualColumn ??
          currentCell?.column ??
          activeRange?.startColumn ??
          0
        const cell = cellKey(row, column)
        const range = activeRange
          ? `${cellKey(
              activeRange.startRow,
              activeRange.startColumn,
            )}:${cellKey(activeRange.endRow, activeRange.endColumn)}`
          : `${cell}:${cell}`
        let formula = ''
        let rawValue: unknown = ''
        let displayValue = ''
        try {
          const cellRange = worksheet.getRange?.(cell)
          formula = cellRange?.getFormula?.() ?? ''
          rawValue = cellRange?.getRawValue?.() ?? cellRange?.getValue?.() ?? ''
          displayValue = cellRange?.getDisplayValue?.() ?? String(rawValue)
        } catch {
          // A selection can momentarily land out of bounds (e.g. Univer moves it
          // past the grid after hiding a block of rows) — getRange throws there.
          // Skip the value read rather than crash the selection sync (#274).
        }
        onSelectionChange({
          cell,
          range,
          editValue: formula || String(rawValue),
          displayValue,
          preserveRange,
        })
      }
      onEditorReady({
        addSheet: async name => {
          if (disposed) return false
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          if (!workbook?.insertSheet) return false
          preserveNextCommandRangeSync = true
          const sheet = workbook.insertSheet(name)
          sheet.activate?.()
          return exportWorkbookAfterBridgeEdit(workbook, false)
        },
        applyRangeStyle: async (rangeToken, patch) => {
          if (disposed) return false
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          // Commit any open in-cell edit first: applying a style while the user
          // is mid-typing would otherwise export the cell's *committed* (often
          // empty) state and silently discard the uncommitted text.
          if (workbook?.isCellEditing?.()) {
            await workbook.endEditingAsync?.(true).catch(() => undefined)
          }
          const worksheet = workbook?.getActiveSheet?.() ?? null
          const range = worksheet?.getRange?.(rangeToken)
          if (!worksheet || !range) return false
          preserveNextCommandRangeSync = true
          if (patch.bold !== undefined)
            range.setFontWeight?.(patch.bold ? 'bold' : 'normal')
          if (patch.italic !== undefined)
            range.setFontStyle?.(patch.italic ? 'italic' : 'normal')
          if (patch.underline !== undefined)
            range.setFontLine?.(patch.underline ? 'underline' : 'none')
          if (patch.fontSize !== undefined) range.setFontSize?.(patch.fontSize)
          if (patch.textColor !== undefined)
            range.setFontColor?.(patch.textColor)
          if (patch.fillColor !== undefined)
            range.setBackgroundColor?.(patch.fillColor)
          if (patch.align !== undefined)
            range.setHorizontalAlignment?.(
              facadeHorizontalAlignment(patch.align),
            )
          if (patch.numberFormat !== undefined)
            range.setNumberFormat?.(
              numberFormatPatternForUniver(patch.numberFormat, patch.decimals),
            )
          if (patch.border !== undefined) {
            const borderType = univerApi.Enum?.BorderType?.ALL ?? 'all'
            const borderStyle = patch.border
              ? univerApi.Enum?.BorderStyleTypes?.THIN ?? 'thin'
              : univerApi.Enum?.BorderStyleTypes?.NONE ?? 'none'
            range.setBorder?.(borderType, borderStyle, CELL_BORDER_COLOR)
          }
          return exportSnapshotAfterBridgeEdit(worksheet)
        },
        clearRange: async rangeToken => {
          if (disposed) return false
          const worksheet =
            univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
          const range = worksheet?.getRange?.(rangeToken)
          if (!worksheet || !range) return false
          preserveNextCommandRangeSync = true
          range.clearContent?.()
          return exportSnapshotAfterBridgeEdit(worksheet)
        },
        applyRowVisibility: async (hiddenRows, rowCount) => {
          if (disposed) return false
          const worksheet =
            univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
          if (!worksheet) return false
          // Row visibility mirrors the shell's filter — it's display-only, not a
          // model edit, so absorb the echoes (like a reload) instead of letting
          // them export/undo.
          suppressExport = true
          try {
            if (rowCount > 0) worksheet.showRows?.(0, rowCount)
            const sorted = [...new Set(hiddenRows)]
              .filter(row => Number.isInteger(row) && row >= 0)
              .sort((a, b) => a - b)
            for (let i = 0; i < sorted.length; ) {
              const start = sorted[i]
              let count = 1
              while (i + 1 < sorted.length && sorted[i + 1] === sorted[i] + 1) {
                i += 1
                count += 1
              }
              worksheet.hideRows?.(start, count)
              i += 1
            }
            return true
          } finally {
            // Row show/hide echoes a burst of commands; absorb them and lift
            // suppression once the stream settles (pre-existing bug: this
            // called a `scheduleSuppressionDrain` that never existed, so the
            // first filter application threw a ReferenceError at runtime and
            // left exports suppressed).
            liftSuppressionAfterRecalc()
          }
        },
        deleteColumn: column =>
          editActiveWorksheet(worksheet => {
            if (!worksheet.deleteColumns) return false
            worksheet.deleteColumns(column, 1)
            return true
          }),
        deleteRow: row =>
          editActiveWorksheet(worksheet => {
            if (!worksheet.deleteRows) return false
            worksheet.deleteRows(row, 1)
            return true
          }),
        deleteSheet: async sheetId => {
          if (disposed) return false
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          if (!workbook?.deleteSheet) return false
          preserveNextCommandRangeSync = true
          const deleted = workbook.deleteSheet(sheetId)
          if (!deleted) return false
          return exportWorkbookAfterBridgeEdit(workbook, false)
        },
        duplicateActiveSheet: async () => {
          if (disposed) return false
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          if (!workbook?.duplicateActiveSheet) return false
          preserveNextCommandRangeSync = true
          const sheet = workbook.duplicateActiveSheet()
          sheet.activate?.()
          return exportWorkbookAfterBridgeEdit(workbook, false)
        },
        insertColumn: column =>
          editActiveWorksheet(worksheet => {
            if (!worksheet.insertColumns) return false
            worksheet.insertColumns(column, 1)
            return true
          }),
        insertRow: row =>
          editActiveWorksheet(worksheet => {
            if (!worksheet.insertRows) return false
            worksheet.insertRows(row, 1)
            return true
          }),
        pasteTabularData: async (cell, text) => {
          if (disposed) return false
          const matrix = parsePasteMatrix(text)
          const worksheet =
            univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
          const start = parseCellKey(cell)
          if (!worksheet || !start || !matrix.length) return false
          const target =
            matrix.length === 1 && matrix[0].length === 1
              ? cell
              : `${cell}:${cellKey(
                  start.row + matrix.length - 1,
                  start.column + Math.max(...matrix.map(row => row.length)) - 1,
                )}`
          const range = worksheet.getRange?.(target)
          if (!range?.setValues) return false
          preserveNextCommandRangeSync = true
          range.setValues(matrix)
          return exportSnapshotAfterBridgeEdit(worksheet)
        },
        renameSheet: async (sheetId, name) => {
          if (disposed) return false
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          const worksheet = workbook?.getSheetBySheetId?.(sheetId) ?? null
          const trimmed = name.trim()
          if (!workbook || !worksheet?.setName || !trimmed) return false
          preserveNextCommandRangeSync = true
          worksheet.setName(trimmed)
          return exportWorkbookAfterBridgeEdit(workbook, true, false)
        },
        selectSheet: async sheetId => {
          if (disposed) return false
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          const worksheet = workbook?.getSheetBySheetId?.(sheetId) ?? null
          if (!workbook || !worksheet) return false
          preserveNextCommandRangeSync = true
          if (workbook.setActiveSheet) workbook.setActiveSheet(worksheet)
          else worksheet.activate?.()
          return exportWorkbookAfterBridgeEdit(workbook, false, false)
        },
        selectRange: async rangeToken => {
          // Move the active selection to highlight a range (the filter column).
          // Display-only: no snapshot export, so it never lands on the undo
          // stack. Guarded because a selection can momentarily be out of bounds.
          if (disposed) return false
          const worksheet =
            univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
          const range = worksheet?.getRange?.(rangeToken)
          if (!worksheet || !range) return false
          try {
            worksheet.setActiveRange?.(range)
          } catch {
            // out-of-bounds / no active viewport — ignore, it's only a highlight
          }
          return true
        },
        setRangeValues: async (rangeToken, values) => {
          if (disposed) return false
          const worksheet =
            univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
          const range = worksheet?.getRange?.(rangeToken)
          if (!worksheet || !range?.setValues || !values.length) return false
          preserveNextCommandRangeSync = true
          range.setValues(values)
          return exportSnapshotAfterBridgeEdit(worksheet)
        },
        flushSnapshot: async () => {
          if (disposed) return null
          await univerApi
            .getFormula?.()
            .onCalculationResultApplied?.(1000)
            .catch(() => undefined)
          const nextSnapshot = saveFromUniver()
          if (nextSnapshot) onSnapshotChange(nextSnapshot)
          return nextSnapshot
        },
        autoFitRows: async (sheetId, rows) => {
          if (disposed) return false
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          const worksheet = workbook?.getSheetBySheetId?.(sheetId)
          if (!workbook || !worksheet?.autoFitRow) return false
          for (const row of rows) worksheet.autoFitRow(row)
          return exportWorkbookAfterBridgeEdit(workbook, true)
        },
        isCellEditing: () =>
          univerApi.getActiveWorkbook?.()?.isCellEditing?.() ?? false,
        abortEditing: async () =>
          (await univerApi.getActiveWorkbook?.()?.abortEditingAsync?.()) ??
          false,
        commitEditing: async () =>
          (await univerApi.getActiveWorkbook?.()?.endEditingAsync?.(true)) ??
          false,
        setColumnWidth: (column, width) =>
          editActiveWorksheet(worksheet => {
            if (!worksheet.setColumnWidth) return false
            worksheet.setColumnWidth(column, width)
            return true
          }),
        revealCell: async cell => {
          // Highlight a Find match; only scroll if it's off-screen, so an
          // in-view match doesn't yank the whole sheet around (#277).
          // Display-only: no snapshot export, so it never hits the undo stack.
          if (disposed) return false
          const worksheet =
            univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
          if (!worksheet) return false
          const pos = parseCellKey(cell)
          try {
            worksheet.getRange?.(cell)?.activate?.()
            if (pos) {
              const scroll = worksheet.getScrollState?.()
              const el = hostRef.current
              // Approx the visible window from default sizes (24px column
              // header, 46px row header, 28px rows, 96px cols) with a 1-cell
              // margin, so a match sitting right at the edge still scrolls.
              const startRow = scroll?.sheetViewStartRow ?? 0
              const startCol = scroll?.sheetViewStartColumn ?? 0
              const rows = el
                ? Math.max(1, Math.floor((el.clientHeight - 24) / 28) - 1)
                : 0
              const cols = el
                ? Math.max(1, Math.floor((el.clientWidth - 46) / 96) - 1)
                : 0
              const inView =
                scroll != null &&
                el != null &&
                pos.row >= startRow &&
                pos.row < startRow + rows &&
                pos.column >= startCol &&
                pos.column < startCol + cols
              if (!inView) worksheet.scrollToCell?.(pos.row, pos.column)
            }
          } catch {
            // selection can momentarily be out of bounds — ignore
          }
          return true
        },
        setFrozenColumns: columns =>
          editActiveWorksheet(worksheet => {
            if (!worksheet.setFrozenColumns) return false
            worksheet.setFrozenColumns(columns)
            return true
          }),
        setFrozenRows: rows =>
          editActiveWorksheet(worksheet => {
            if (!worksheet.setFrozenRows) return false
            worksheet.setFrozenRows(rows)
            return true
          }),
        setRowHeight: (row, height) =>
          editActiveWorksheet(worksheet => {
            if (!worksheet.setRowHeight) return false
            worksheet.setRowHeight(row, height)
            return true
          }),
        setCellValue: async (cell, value) => {
          if (disposed) return false
          const worksheet =
            univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
          const range = worksheet?.getRange?.(cell)
          if (!worksheet || !range) return false
          preserveNextCommandRangeSync = true
          if (value.trim().startsWith('=')) range.setFormula?.(value)
          else range.setValue?.(value)
          return exportSnapshotAfterBridgeEdit(worksheet)
        },
      })
      readyTimeout = window.setTimeout(() => {
        readyForChanges = true
        syncSelectionFromUniver()
      }, 750)
      const selectionDisposable = univerApi.addEvent?.(
        univerApi.Event?.SelectionChanged ?? 'SelectionChanged',
        params =>
          syncSelectionFromUniver(
            params.worksheet ?? params.workbook?.getActiveSheet?.() ?? null,
            params.selections,
          ),
      )
      const editEndDisposable = univerApi.addEvent?.(
        univerApi.Event?.SheetEditEnded ?? 'SheetEditEnded',
        params => {
          window.setTimeout(
            () =>
              syncSelectionFromUniver(
                params.worksheet ?? params.workbook?.getActiveSheet?.() ?? null,
              ),
            0,
          )
        },
      )
      const formulaDisposable = univerApi
        .getFormula?.()
        .calculationResultApplied?.(() => {
          syncSelectionFromUniver()
        })
      // Lift reload suppression only once the reload's formula recalc has
      // actually settled — not after a guessed quiet interval. The old 350ms
      // timer was shorter than recalc can take (the bridge itself waits up to
      // 1000ms via onCalculationResultApplied), so a slow recalc write-back
      // (`sheet.mutation.set-range-values`) landed after suppression lifted,
      // exported, and pushed a phantom undo entry while clearing redo — the
      // "undo then undo redoes it" bug (#271). `onCalculationResultApplied`
      // resolves when the result is applied, or after its own short watchdog
      // when there is nothing to recalc, so this never hangs.
      const liftSuppressionAfterRecalc = (): void => {
        if (suppressDrainTimeout) {
          window.clearTimeout(suppressDrainTimeout)
          suppressDrainTimeout = undefined
        }
        const token = ++suppressToken
        const settle = univerApi
          .getFormula?.()
          .onCalculationResultApplied?.(1000)
        const lift = (): void => {
          if (token !== suppressToken) return // a newer reload owns suppression
          suppressExport = false
        }
        if (settle) {
          settle.then(lift, lift)
        } else {
          // No formula facade — nothing async to wait on; lift next tick.
          suppressDrainTimeout = window.setTimeout(lift, 0)
        }
      }
      const commandDisposable = univerApi.onCommandExecuted?.(command => {
        if (!readyForChanges) return
        const commandId = command?.id ?? ''
        // A canvas copy/cut (Ctrl+C / Ctrl+X) is handled entirely by Univer and
        // never reaches our Copy button, so the shell's internal clipboard —
        // which the paste-values / paste-formatting buttons read — stays empty.
        // Mirror it: capture the selection into the shell clipboard on copy/cut.
        // Copy changes no data, so stop there; cut also clears cells → let that
        // fall through to the normal export below.
        if (
          commandId === 'univer.command.copy' ||
          commandId === 'univer.command.cut'
        ) {
          onEditorCopyRef.current?.()
          if (commandId === 'univer.command.copy') return
        }
        if (suppressExport) {
          // Echo from a programmatic reload — drop it, so it never reaches the
          // shell as a user edit. Suppression is lifted by the recalc-settle
          // gate (liftSuppressionAfterRecalc), not by these echoes.
          return
        }
        // A bridge method owns this edit and will export exactly one snapshot
        // itself — drop the command echo so the edit isn't committed twice (#271).
        if (suppressCommandExport) {
          return
        }
        // Only committed sheet-model *mutations* are undoable edits — Univer's
        // own undo history is mutation-based, so this is the same set of events
        // it records. Everything else must NOT export:
        //   • the live in-cell editor (`doc.*`) — each keystroke was landing as
        //     its own undo entry, and exporting mid-edit saved the sheet before
        //     the typed value was committed, silently deleting the content;
        //   • formula recalc (`formula.*`);
        //   • pure UI commands/operations — selection, scroll, cursor, tab-move
        //     (`sheet.operation.*`, and `sheet.command.*` navigation like
        //     `scroll-view` / `move-selection-enter-tab`). These are type-0
        //     COMMANDs that change no model state, but were slipping through and
        //     adding a spurious second undo entry per typed value (#271).
        if (!commandId.startsWith('sheet.mutation.')) {
          return
        }
        // Open a new edit span on the first content mutation; later mutations
        // (notably the formula recalc write-back) fold into it.
        if (!editOpen) {
          editOpen = true
          recordedThisEdit = false
        }
        const myToken = ++editToken
        // A short debounce only batches the synchronous mutation burst of one
        // action (was 250ms — long enough for a user Undo to race in and get the
        // edit cancelled by the reload, losing it (#271)). Recording promptly
        // means the undo entry lands before any following user action.
        if (snapshotTimeout) window.clearTimeout(snapshotTimeout)
        const preserveRange = preserveNextCommandRangeSync
        preserveNextCommandRangeSync = false
        snapshotTimeout = window.setTimeout(() => {
          const record = !recordedThisEdit
          recordedThisEdit = true
          const nextSnapshot = saveFromUniver()
          if (nextSnapshot) onSnapshotChange(nextSnapshot, { record })
          syncSelectionFromUniver(undefined, undefined, preserveRange)
          // Close the edit span once recalc has settled, unless a newer
          // mutation has already superseded this one.
          const settle = univerApi
            .getFormula?.()
            .onCalculationResultApplied?.(1000)
          const close = (): void => {
            if (myToken !== editToken) return
            editOpen = false
            recordedThisEdit = false
          }
          settle ? settle.then(close, close) : close()
        }, 40)
      })

      // Put the canvas back where the user left it after a unit reload: same
      // active sheet, same selection, same scroll position. Best-effort —
      // a facade missing one of these methods just skips that part of the
      // restore; the reloaded data is already correct either way.
      const restoreViewState = (
        sheetId: string | null,
        range: UniverRangeLike | null,
        scroll: {
          sheetViewStartRow?: number
          sheetViewStartColumn?: number
        } | null,
      ): void => {
        try {
          const workbook = univerApi.getActiveWorkbook?.() ?? null
          if (!workbook) return
          const worksheet =
            (sheetId ? workbook.getSheetBySheetId?.(sheetId) : null) ??
            workbook.getActiveSheet?.() ??
            null
          if (!worksheet) return
          worksheet.activate?.()
          if (range) {
            const token = `${cellKey(
              range.startRow,
              range.startColumn,
            )}:${cellKey(range.endRow, range.endColumn)}`
            worksheet.getRange?.(token)?.activate?.()
          }
          const scrollRow = scroll?.sheetViewStartRow ?? range?.startRow
          const scrollColumn = scroll?.sheetViewStartColumn ?? range?.startColumn
          if (scrollRow !== undefined && scrollColumn !== undefined) {
            worksheet.scrollToCell?.(scrollRow, scrollColumn)
          }
        } catch {
          // View-state restore is cosmetic; never let it break the reload.
        }
      }

      // Re-sync the canvas with the model without tearing down Univer: replace
      // the workbook unit in place, keeping the instance, plugins, api and
      // listeners alive. This is what a `revision` bump now triggers instead of
      // a full remount, so undo/redo/paste/protect no longer flash the whole
      // surface (#269).
      const reloadUnit = (): void => {
        const nextSnapshot = lastSnapshotRef.current
        if (disposed || !nextSnapshot) return
        // Capture the live view state first: a recreated unit resets to the
        // first sheet with A1:A1 selected and the viewport scrolled to the
        // top-left (the generated snapshot carries no selection and zeroed
        // scroll offsets), which yanked the user's selection away on every
        // model-driven refresh — most visibly right after sorting a range
        // (#275).
        const previousWorksheet =
          univerApi.getActiveWorkbook?.()?.getActiveSheet?.() ?? null
        const previousSheetId = previousWorksheet?.getSheetId?.() ?? null
        let previousRange: UniverRangeLike | null = null
        try {
          previousRange =
            previousWorksheet
              ?.getSelection?.()
              ?.getActiveRange?.()
              ?.getRange?.() ?? null
        } catch {
          previousRange = null
        }
        let previousScroll: {
          sheetViewStartRow?: number
          sheetViewStartColumn?: number
        } | null = null
        try {
          previousScroll = previousWorksheet?.getScrollState?.() ?? null
        } catch {
          previousScroll = null
        }
        // Absorb the command echo that disposeUnit + createUnit emit while
        // loading data (however long the burst takes) so the reload never
        // pushes a spurious snapshot back to the shell (which would clear the
        // redo stack and double the undo stack).
        suppressExport = true
        if (snapshotTimeout) window.clearTimeout(snapshotTimeout)
        try {
          univerApi.disposeUnit?.(nextSnapshot.id)
        } catch {
          // If the unit is already gone, createUnit below re-establishes it.
        }
        univer.createUnit(
          UniverInstanceType.UNIVER_SHEET,
          cloneSnapshot(nextSnapshot),
        )
        const targetSheetId = requestedSheetId(nextSnapshot) ?? previousSheetId
        const sameSheet = targetSheetId === previousSheetId
        restoreViewState(targetSheetId, sameSheet ? previousRange : null, sameSheet ? previousScroll : null)
        window.dispatchEvent(new Event('resize'))
        applyCanvasPermissions()
        applyConditionalFormats()
        applyValidationFlags()
        liftSuppressionAfterRecalc()
      }
      reloadUnitRef.current = reloadUnit
      // Univer ignores our app-owned activeSheetId metadata on initial mount.
      if (snapshot) restoreViewState(requestedSheetId(snapshot), null, null)

      disposeUniver = () => {
        reloadUnitRef.current = null
        applyValidationFlagsRef.current = null
        resizeObserver.disconnect()
        if (resizeRaf) window.cancelAnimationFrame(resizeRaf)
        if (snapshotTimeout) window.clearTimeout(snapshotTimeout)
        if (suppressDrainTimeout) window.clearTimeout(suppressDrainTimeout)
        if (readyTimeout) window.clearTimeout(readyTimeout)
        selectionDisposable?.dispose()
        editEndDisposable?.dispose()
        formulaDisposable?.dispose()
        commandDisposable?.dispose()
        onEditorReady(null)
        window.setTimeout(() => {
          try {
            univer.dispose()
          } catch {
            // Univer owns a nested React tree; HMR/reload can race its internal unmount.
          }
        }, 0)
      }
    }

    void mountUniver()

    return () => {
      disposed = true
      onEditorReady(null)
      disposeUniver?.()
    }
  }, [snapshot?.id, onEditorReady, onSelectionChange, onSnapshotChange])

  // A `revision` bump asks the canvas to re-sync with the model (undo/redo,
  // paste, fill, sort, protect, …). Reload the workbook unit in place instead
  // of remounting the whole Univer surface. No-op until the editor has mounted
  // (the mount itself loads the current snapshot) (#269).
  useEffect(() => {
    reloadUnitRef.current?.()
  }, [revision])

  // Repaint validation flags when the flagged-cell set actually changes (the
  // shell recomputes the prop on every model update; compare by content so a
  // fresh-but-identical object doesn't churn the CF engine).
  useEffect(() => {
    const signature = JSON.stringify(invalidCells ?? {})
    if (signature === invalidCellsSignatureRef.current) return
    invalidCellsSignatureRef.current = signature
    applyValidationFlagsRef.current?.()
  }, [invalidCells])

  return (
    <UniverMountFrame
      ref={hostRef}
      aria-label="Univer spreadsheet editor"
      data-testid="mounted-univer-editor"
    />
  )
}

// Univer's facade spells right alignment 'normal' (and throws on 'right' —
// see transformFacadeHorizontalAlignment in @univerjs/sheets facade), so the
// toolbar's 'right' must be converted before dispatching (#270).
function facadeHorizontalAlignment(
  align: NonNullable<CellStyle['align']>,
): 'left' | 'center' | 'normal' {
  if (align === 'center') return 'center'
  if (align === 'right') return 'normal'
  return 'left'
}

// Translate one PureSheets conditional-format rule into a built Univer CF
// rule via the (duck-typed) highlight-rule builder. Returns null when the
// rule cannot be represented (missing facade methods, non-numeric bounds for
// numeric conditions) — matching the model evaluator's semantics, which also
// treats those as non-matching (`conditionMatches` in sheetDataFeatures).
function buildUniverConditionalFormat(
  builder: UniverConditionalFormatBuilder | undefined,
  rule: ConditionalFormatRule,
): unknown {
  if (!builder) return null
  const [startKey, endKey] = rule.range.split(':')
  const start = parseCellKey(startKey ?? '')
  const end = parseCellKey(endKey ?? startKey ?? '')
  if (!start || !end) return null
  const { kind, value, value2 } = rule.condition
  const numeric = Number(value)
  const numericOk = value.trim() !== '' && Number.isFinite(numeric)
  let conditioned: UniverConditionalFormatBuilder | undefined
  if (kind === 'greater') {
    conditioned = numericOk
      ? builder.whenNumberGreaterThan?.(numeric)
      : undefined
  } else if (kind === 'less') {
    conditioned = numericOk ? builder.whenNumberLessThan?.(numeric) : undefined
  } else if (kind === 'between') {
    const upper = Number(value2)
    conditioned =
      numericOk && Number.isFinite(upper)
        ? builder.whenNumberBetween?.(
            Math.min(numeric, upper),
            Math.max(numeric, upper),
          )
        : undefined
  } else if (kind === 'equal') {
    // Mirrors the model evaluator: numeric equality when the target parses
    // as a number, text equality otherwise.
    conditioned = numericOk
      ? builder.whenNumberEqualTo?.(numeric)
      : builder.whenTextEqualTo?.(value)
  } else if (kind === 'text-contains') {
    conditioned = builder.whenTextContains?.(value)
  }
  if (!conditioned) return null
  let styled = conditioned
  if (rule.style.fillColor)
    styled = styled.setBackground?.(rule.style.fillColor) ?? styled
  if (rule.style.textColor)
    styled = styled.setFontColor?.(rule.style.textColor) ?? styled
  if (rule.style.bold) styled = styled.setBold?.(true) ?? styled
  if (rule.style.italic) styled = styled.setItalic?.(true) ?? styled
  if (rule.style.underline) styled = styled.setUnderline?.(true) ?? styled
  const ranged = styled.setRanges?.([
    {
      startRow: start.row,
      startColumn: start.column,
      endRow: end.row,
      endColumn: end.column,
    },
  ])
  return ranged?.build?.() ?? null
}

function numberFormatPatternForUniver(
  format: NonNullable<CellStyle['numberFormat']>,
  decimals?: number,
): string {
  const zeros =
    decimals !== undefined && decimals > 0
      ? `.${'0'.repeat(Math.min(decimals, 8))}`
      : ''
  if (format === 'currency')
    return decimals === undefined ? '$#,##0.00' : `$#,##0${zeros}`
  if (format === 'percent')
    return decimals === undefined ? '0.00%' : `0${zeros}%`
  if (format === 'date') return 'yyyy-mm-dd'
  if (format === 'number')
    return decimals === undefined ? '#,##0.########' : `#,##0${zeros}`
  return '@'
}

const UniverMountFrame = styled.div`
  // Fill the pane exactly so Univer's canvas owns scrolling. A fixed
  // min-width wider than the pane forced the wrapper to scroll horizontally,
  // adding a second scrollbar alongside Univer's own (#272).
  width: 100%;
  min-width: 0;
  min-height: 0;
  height: 100%;
  border-bottom: 1px solid var(--sheets-line);
  background: var(--sheets-canvas);

  > * {
    min-height: 0;
    height: 100%;
  }

  &:empty::before {
    display: grid;
    place-items: center;
    min-height: 100%;
    color: var(--platform-colors-text-muted, #687064);
    font-size: 12px;
    content: 'Univer spreadsheet editor';
  }
`
