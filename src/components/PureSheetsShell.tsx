import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { useSheetsAgentTools } from '../hooks/useSheetsAgentTools'
import { AppFrame } from '@purescience/platform-bridge/components/AppFrame'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import { useDocumentLifecycle } from '@purescience/platform-ui/bridge/react/useDocumentLifecycle'
import { useDocumentHotkeys } from '@purescience/platform-ui/bridge/react/useDocumentHotkeys'
import {
  DocumentHeaderActions,
  DocumentSwitcher,
} from '@purescience/platform-ui/components/common/documents'
import {
  chooseSaveFilePath,
  openImportFile,
  readBinaryFile,
  readClipboard,
  readTextFile,
  recordOperation,
  updatePureSheetsSettings,
  writeBinaryFile,
  writeClipboard,
  writeTextFile,
} from '../bridge/platformBridge'
import type { PlatformOperationInput } from '../bridge/platformBridge'
import { sheetsSnapshotHtml } from '../lib/sheetsSnapshot'
import {
  readWorkbookFromPath,
  workbookFilesForSave,
} from '../lib/workbookFiles'
import {
  createUniverSnapshot,
  documentFromEditorSnapshot,
  type PureSheetsUniverSnapshot,
} from '../lib/univerAdapter'
import { PURESHEETS_APP_SLUG } from '../constants'
import {
  addSheet as addWorkbookSheet,
  addSheetToWorkbook,
  applyTableStylePreset,
  createBlankWorkbook,
  cellKey,
  cellsInRange,
  clearRange,
  cloneWithUpdatedSheet,
  columnLabel,
  computedCellValue,
  copyRangeCells,
  createChartFromRange,
  deleteColumn,
  deleteRow,
  fillRange,
  findMatches,
  isCellLocked,
  isRangeLocked,
  lockRange,
  pasteCellBlock,
  replaceAllInSheet,
  setCellComment,
  setFilterQuery,
  setFilterHasHeader,
  sheetHasPositionalMetadata,
  unlockRange,
  updateChart,
  deleteSheet as deleteWorkbookSheet,
  displayCellValue,
  duplicateSheet as duplicateWorkbookSheet,
  insertColumn,
  insertRow,
  pasteTabularData,
  isFormulaError,
  normalizeRangeToken,
  parseCellKey,
  rangeToDelimitedText,
  renameSheet as renameWorkbookSheet,
  removeChart,
  replaceFirstOccurrence,
  replaceInRange,
  setColumnWidth,
  setRangeStyle,
  setRowHeight,
  sortRange,
  toggleFilterRow,
  toggleFreezeFirstColumn,
  toggleFreezeFirstRow,
  upsertCell,
  visibleRowsForSheet,
  type CopiedCellBlock,
  type PasteMode,
  type TableStylePreset,
} from '../lib/workbookModel'
import {
  addConditionalFormatRule,
  addValidationRule,
  clearConditionalFormatsInRange,
  clearValidationInRange,
  invalidCellFlagsForSheet,
  validateValue,
  validationRuleForCell,
  type InvalidCellFlag,
} from '../lib/sheetDataFeatures'
import { csvFromSheet, sheetFromCsv } from '../lib/csv'
import {
  base64ToBytes,
  bytesToBase64,
  exportXlsxWorkbook,
  importXlsxWorkbook,
} from '../lib/xlsxIO'
import type {
  CellStyle,
  ConditionalFormatRule,
  PureSheetsDocument,
  WorkbookChart,
  WorkbookSheet,
} from '../types'
import { AgentActivityBar } from './AgentActivityBar'
import {
  CommandBar,
  type FormulaTemplate,
  type ImportExportAction,
} from './CommandBar'
import { CommentBar } from './CommentBar'
import type { ValidationKind } from './DataRuleControls'
import { CompatibilityGrid } from './CompatibilityGrid'
import { ChartPanel, type ChartUpdatePatch } from './ChartPanel'
import { FormulaBar } from './FormulaBar'
import { SheetTabsBar } from './SheetTabsBar'
import { UniverSpreadsheetSurface } from './UniverSpreadsheetSurface'
import type {
  UniverEditorBridge,
  UniverPasteValue,
  UniverSelectionState,
} from './univerBridgeTypes'

interface PureSheetsShellProps {
  initialDocument: PureSheetsDocument
  resource?: { path?: string } | null
  onResourceHandled?: () => void
}

import { completeWorkbook as createWorkbookReceipt } from '../lib/completeWorkbook'

type SaveState = 'saved' | 'dirty' | 'saving' | 'error'

/**
 * PureSheetsShell — the orchestrator.
 *
 * Owns document state, the unified save lifecycle, selection, and undo/redo.
 * Rendering is delegated to focused region components: `CommandBar`,
 * `FormulaBar`, `UniverSpreadsheetSurface` (the Univer editor),
 * `CompatibilityGrid` (legacy fallback grid), `ChartPanel`,
 * `AgentActivityBar` (the agent-log viewer), and `SheetTabsBar`. File
 * import/export flows live in `lib/workbookFiles.ts`.
 */
export function PureSheetsShell({
  initialDocument,
  resource,
  onResourceHandled,
}: PureSheetsShellProps): React.ReactElement {
  const [document, setDocumentState] = useState(initialDocument)
  const documentRef = useRef(document)
  const setDocument = useCallback((update: PureSheetsDocument | ((current: PureSheetsDocument) => PureSheetsDocument)): void => {
    const next = typeof update === 'function' ? update(documentRef.current) : update
    documentRef.current = next
    setDocumentState(next)
  }, [])
  const [filePath, setFilePath] = useState('')
  const [selectedCell, setSelectedCell] = useState('A1')
  const [editValue, setEditValue] = useState('')
  const [saveState, setSaveState] = useState<SaveState>('saved')
  const [statusMessage, setStatusMessage] = useState('Ready')
  // Brief, prominent confirmation (e.g. copy/cut) — easy to miss in the status
  // bar, so surface it as a fading toast too (#277).
  const [toast, setToast] = useState<string | null>(null)
  const toastTimerRef = useRef<number | null>(null)
  function showToast(message: string): void {
    setToast(message)
    if (toastTimerRef.current) window.clearTimeout(toastTimerRef.current)
    toastTimerRef.current = window.setTimeout(() => setToast(null), 1800)
  }
  // The cell the last Find landed on, so a repeat Find advances past it while a
  // fresh Find can still land on the cell you're currently sitting on (#277).
  const lastFindRef = useRef<string | null>(null)
  const [undoStack, setUndoStack] = useState<PureSheetsDocument[]>([])
  const [redoStack, setRedoStack] = useState<PureSheetsDocument[]>([])
  const [findText, setFindText] = useState('')
  const [replaceText, setReplaceText] = useState('')
  const [matchCase, setMatchCase] = useState(false)
  const [commentBarOpen, setCommentBarOpen] = useState(false)
  const [importWarnings, setImportWarnings] = useState<string[] | null>(null)
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null)
  const [renameDraft, setRenameDraft] = useState('')
  const [agentActivityOpen, setAgentActivityOpen] = useState(false)
  const [univerSurfaceRevision, setUniverSurfaceRevision] = useState(0)
  const [selectedRange, setSelectedRange] = useState('A1:A1')
  const [univerSelection, setUniverSelection] =
    useState<UniverSelectionState | null>(null)
  useSheetsAgentTools(true, {
    document,
    setDocument: updater => {
      commitDocument(updater)
      refreshUniverSurface()
    },
    filePath,
    currentFilePath: () => boundPathRef.current ?? '',
    selectedRange,
    autoFitRows: async (sheetId, rows) =>
      await univerEditorRef.current?.autoFitRows(sheetId, rows) ?? false,
    createWorkbook: async title => {
      const previous = documentRef.current
      const previousPath = boundPathRef.current
      if ((filePath.trim() || null) !== previousPath)
        throw new Error('Workbook is switching; retry creation once it is open.')
      // flush is a forced write, not a dirty-only drain. Do not rewrite a clean
      // restored file merely because the agent is creating another workbook.
      if (saveState !== 'saved' || lifecycleRef.current.doc.saving)
        await lifecycleRef.current.flush({ throwOnError: true })
      if (documentRef.current !== previous || boundPathRef.current !== previousPath)
        throw new Error('Workbook changed while saving; retry creation.')
      // Flush succeeded against the old binding. Detach it synchronously before
      // exposing the blank document, so subsequent agent writes cannot save over it.
      lifecycleRef.current.reset()
      boundPathRef.current = null
      lastPathRef.current = null
      savedWorkbookPayloadRef.current = null
      const fresh = createBlankWorkbook(title)
      documentRef.current = fresh
      boundDocRef.current = fresh
      newWorkbook(fresh)
      return { title: fresh.metadata.title, filePath: null }
    },
    completeWorkbook: async () => {
      const snapshot = documentRef.current
      const originalPath = boundPathRef.current
      let savedPath: string | null = null
      return createWorkbookReceipt({
        save: async () => {
          if ((filePath.trim() || null) !== originalPath)
            throw new Error('Workbook is switching; retry completion once it is open.')
          boundDocRef.current = snapshot
          // A new draft is a package before the lifecycle serializes it.
          if (!originalPath) boundIsPackageRef.current = true
          const lifecycle = lifecycleRef.current
          savedPath = await lifecycle.ensureDraft()
          if (!savedPath) throw new Error('Could not create a saved workbook draft.')
          if (documentRef.current !== snapshot)
            throw new Error('Workbook changed while creating its draft.')
          await lifecycle.flush({ throwOnError: true })
          const payload = savedWorkbookPayloadRef.current
          if (!payload) throw new Error('Workbook save did not return its written content.')
          return { path: savedPath, contentPath: payload.name ? `${savedPath}/${payload.name}` : savedPath, content: payload.content }
        },
        read: readTextFile,
        unchanged: () => documentRef.current === snapshot &&
          (boundPathRef.current === originalPath || boundPathRef.current === savedPath),
      })
    },
  })
  const manualRangeRef = useRef(false)
  const univerEditorRef = useRef<UniverEditorBridge | null>(null)
  // In-app clipboard: the copied cell block (formulas + styles) plus the TSV
  // that was written to the system clipboard at copy time. When a paste's
  // clipboard text still matches, the block is pasted instead of the lossy
  // TSV, and it also backs paste-values / paste-formatting.
  const copiedBlockRef = useRef<CopiedCellBlock | null>(null)
  const copiedTextRef = useRef<string | null>(null)
  const formulaRangePickRef = useRef<{
    template: FormulaTemplate
    targetCell: string
  } | null>(null)
  const formulaRangePickTimeoutRef = useRef<number | undefined>(undefined)
  const formulaRangeCommitRef = useRef<{
    formula: string
    range: string
    targetCell: string
    template: FormulaTemplate
  } | null>(null)
  const formulaRangeCommitClearRef = useRef<number | undefined>(undefined)

  // ---- Unified document lifecycle -----------------------------------------
  // One workbook is one `.sheets` (or `.sheets.html`) file. `boundDocRef`
  // holds the workbook belonging to the bound file and only advances while
  // that binding is unchanged, so switching files flushes the OLD workbook
  // to the OLD path first.
  const boundPathRef = useRef<string | null>(null)
  const boundDocRef = useRef(document)
  const lastPathRef = useRef<string | null | undefined>(undefined)
  // New workbooks are `.sheets` package folders; legacy flat `.sheets` /
  // `.sheets.html` files opened from disk stay single files. The package
  // suffix collides with the legacy flat suffix, so package-ness can't be
  // read from the path — it's tracked explicitly (set at bind/open time).
  const boundIsPackageRef = useRef(false)
  const [switcherOpen, setSwitcherOpen] = useState(false)
  const [titleDraft, setTitleDraft] = useState<string | null>(null)
  const savedWorkbookPayloadRef = useRef<{ name: string | null; content: string } | null>(null)
  // A slow restored file must not replace a newer open or an explicit New.
  const openGenerationRef = useRef(0)
  useEffect(() => () => { openGenerationRef.current += 1 }, [])

  const lifecycle = useDocumentLifecycle({
    appSlug: PURESHEETS_APP_SLUG,
    suffix: '.sheets',
    kind: 'package',
    suggestedTitle: document.metadata.title,
    onSaved: files => {
      const file = files.find(entry => entry.name === 'workbook.json') ?? files.find(entry => entry.name === null)
      savedWorkbookPayloadRef.current = file && typeof file.content === 'string'
        ? { name: file.name ?? null, content: file.content } : null
    },
    serialize: () =>
      workbookFilesForSave(boundDocRef.current, {
        boundPath: boundPathRef.current,
        isPackage: boundIsPackageRef.current,
      }),
  })
  const lifecycleRef = useRef(lifecycle)
  useEffect(() => {
    lifecycleRef.current = lifecycle
  }, [lifecycle])

  // Document switch (open, new workbook): flush the old file, then bind.
  const boundTarget = filePath.trim() || null
  useEffect(() => {
    if (lastPathRef.current === boundTarget) return
    lastPathRef.current = boundTarget
    if (boundPathRef.current === boundTarget) return
    const lifecycle = lifecycleRef.current
    void lifecycle.flush().finally(() => {
      boundPathRef.current = boundTarget
      boundDocRef.current = documentRef.current
      if (boundTarget) lifecycle.adopt(boundTarget)
      else lifecycle.reset()
    })
  }, [boundTarget])

  // Lifecycle-owned path changes (lazy draft, promote, rename) → filePath.
  // These paths are always `.sheets` packages (drafts and their promotions).
  useEffect(() => {
    const path = lifecycle.doc.path
    if (!path || boundPathRef.current === path) return
    boundIsPackageRef.current = true
    boundPathRef.current = path
    lastPathRef.current = path
    setFilePath(path)
  }, [lifecycle.doc.path])

  // Every user edit marks the bound workbook dirty — the first one lazily
  // creates the durable draft in PureDrafts.
  useEffect(() => {
    if (saveState !== 'dirty') return
    if ((filePath.trim() || null) !== boundPathRef.current) return
    boundDocRef.current = document
    lifecycleRef.current.markDirty()
  }, [document, filePath, saveState])

  // Reflect completed lifecycle saves in the local save-state latch.
  useEffect(() => {
    if (lifecycle.doc.savedAt && !lifecycle.doc.saving) setSaveState('saved')
  }, [lifecycle.doc.savedAt, lifecycle.doc.saving])
  // ---- end document lifecycle ----------------------------------------------

  const activeSheet = useMemo(() => {
    const selected =
      document.workbook.sheets.find(
        sheet => sheet.id === document.workbook.activeSheetId,
      ) ?? document.workbook.sheets[0]
    if (selected?.agentDraft) {
      return (
        document.workbook.sheets.find(
          sheet => sheet.id === selected.agentDraft?.sourceSheetId,
        ) ??
        document.workbook.sheets.find(sheet => !sheet.agentDraft) ??
        selected
      )
    }
    return selected
  }, [document])
  const selectedPosition = useMemo(
    () => parseCellKey(selectedCell),
    [selectedCell],
  )
  const selectedStyle = activeSheet.cells[selectedCell]?.style
  const gridTemplate = useMemo(() => {
    return Array.from({ length: activeSheet.columnCount }, (_, column) => {
      return `${activeSheet.columnWidths?.[String(column)] ?? 112}px`
    }).join(' ')
  }, [activeSheet])
  const visibleRows = useMemo(
    () => visibleRowsForSheet(activeSheet),
    [activeSheet],
  )
  const univerSnapshot = useMemo(() => {
    return isUniverSnapshot(document.engine.snapshot)
      ? document.engine.snapshot
      : createUniverSnapshot(document)
  }, [document])

  // Shared file loader. Both the platform "open this resource" path and the
  // boot-time self-reopen of the last file route through this so that
  // `document` and `filePath` are always set together — never a real path
  // pointing at unloaded (blank) content.
  const openWorkbookPath = useCallback(
    async (resourcePath: string): Promise<void> => {
      const generation = ++openGenerationRef.current
      let loaded: Awaited<ReturnType<typeof readWorkbookFromPath>>
      try {
        loaded = await readWorkbookFromPath(resourcePath, readTextFile)
      } catch (error) {
        if (generation !== openGenerationRef.current) return
        throw error
      }
      if (generation !== openGenerationRef.current) return
      boundIsPackageRef.current = loaded.isPackage
      setDocument(loaded.document)
      // Same-ID opens do not remount Univer. Explicitly reload its unit so
      // the canvas receives the loaded bytes as well as the React controls.
      setUniverSurfaceRevision(revision => revision + 1)
      setUndoStack([])
      setRedoStack([])
      setFilePath(loaded.path)
      setSelectedCell('A1')
      manualRangeRef.current = false
      setSelectedRange('A1:A1')
      setTitleDraft(null)
      // A recovered document is dirty on purpose: the next save rewrites
      // the corrupt main file with the good backup contents.
      setSaveState(loaded.recoveredFromBackup ? 'dirty' : 'saved')
      setStatusMessage(
        loaded.recoveredFromBackup
          ? `The file was unreadable; recovered the last good save from ${
              loaded.path.split(/[\\/]/).pop() ?? loaded.path
            }.bak`
          : `Opened ${loaded.path.split(/[\\/]/).pop() ?? loaded.path}`,
      )
      await updatePureSheetsSettings({ lastFilePath: loaded.path })
    },
    [],
  )

  useEffect(() => {
    const path = resource?.path?.trim()
    if (!path) return
    let cancelled = false
    void openWorkbookPath(path)
      .catch(error => {
        if (cancelled) return
        setStatusMessage(error instanceof Error ? error.message : String(error))
        setSaveState('error')
      })
      .finally(() => {
        if (cancelled) return
        onResourceHandled?.()
      })
    return () => {
      cancelled = true
    }
  }, [resource?.path, onResourceHandled, openWorkbookPath])

  useEffect(() => {
    setEditValue(activeSheet.cells[selectedCell]?.value ?? '')
  }, [activeSheet, selectedCell])

  useEffect(() => {
    return () => {
      if (formulaRangePickTimeoutRef.current)
        window.clearTimeout(formulaRangePickTimeoutRef.current)
      if (formulaRangeCommitClearRef.current)
        window.clearTimeout(formulaRangeCommitClearRef.current)
    }
  }, [])

  useEffect(() => {
    const pendingPick = formulaRangePickRef.current
    if (!pendingPick) return
    const range = normalizeRangeToken(selectedRange)
    if (!range) return
    const formula = formulaTemplate(
      pendingPick.template,
      range,
      pendingPick.targetCell,
      {
        exactRange: true,
      },
    )
    setEditValue(formula)
    if (formulaRangePickTimeoutRef.current) {
      window.clearTimeout(formulaRangePickTimeoutRef.current)
    }
    formulaRangePickTimeoutRef.current = window.setTimeout(() => {
      formulaRangeCommitRef.current = {
        formula,
        range,
        targetCell: pendingPick.targetCell,
        template: pendingPick.template,
      }
      if (formulaRangeCommitClearRef.current) {
        window.clearTimeout(formulaRangeCommitClearRef.current)
      }
      commitCellValue(pendingPick.targetCell, formula)
      setSelectedCell(pendingPick.targetCell)
      setSelectedRange(range)
      setStatusMessage(`Inserted ${pendingPick.template} formula from ${range}`)
      formulaRangePickRef.current = null
      formulaRangePickTimeoutRef.current = undefined
      formulaRangeCommitClearRef.current = window.setTimeout(() => {
        formulaRangeCommitRef.current = null
        formulaRangeCommitClearRef.current = undefined
      }, 1200)
    }, 350)
    return () => {
      if (formulaRangePickTimeoutRef.current) {
        window.clearTimeout(formulaRangePickTimeoutRef.current)
        formulaRangePickTimeoutRef.current = undefined
      }
    }
  }, [selectedRange])

  // Fire-and-forget append to the platform operations ledger. Never blocks or
  // breaks an edit when the bridge is unavailable (tests, standalone dev).
  function recordLedgerOperation(
    input: Omit<PlatformOperationInput, 'appSlug'>,
  ): void {
    try {
      void recordOperation({ appSlug: PURESHEETS_APP_SLUG, ...input })?.catch?.(
        () => undefined,
      )
    } catch {
      // operations bridge unavailable — the interaction itself must not fail
    }
  }

  function commitDocument(
    updater: (current: PureSheetsDocument) => PureSheetsDocument,
    message?: string,
  ): void {
    const current = documentRef.current
    const next = updater(current)
    if (next === current) return
    setUndoStack(stack => [...stack, current].slice(-50))
    setRedoStack([])
    setDocument(next)
    setSaveState('dirty')
    if (message) setStatusMessage(message)
  }

  function refreshUniverSurface(): void {
    setUniverSurfaceRevision(revision => revision + 1)
  }

  function updateActiveSheet(nextSheet: WorkbookSheet, message?: string): void {
    commitDocument(
      current => cloneWithUpdatedSheet(current, nextSheet),
      message,
    )
  }

  function updateActiveSheetFromCurrent(
    updater: (sheet: WorkbookSheet) => WorkbookSheet,
    message?: string,
  ): void {
    commitDocument(current => {
      const currentSheet =
        current.workbook.sheets.find(
          sheet => sheet.id === current.workbook.activeSheetId,
        ) ?? current.workbook.sheets[0]
      return cloneWithUpdatedSheet(current, updater(currentSheet))
    }, message)
  }

  function commitCellValue(cell: string, value: string): void {
    // Protection (Phase S3): refuse edits to locked cells on every surface —
    // the model would silently no-op, so tell the user why instead.
    if (isCellLocked(activeSheet, cell)) {
      setEditValue(activeSheet.cells[cell]?.value ?? '')
      setStatusMessage(`${cell} is protected — unlock the range to edit it`)
      return
    }
    // Validation (Phase S3): flag, never reject.
    const rule = validationRuleForCell(activeSheet, cell)
    const validation = rule ? validateValue(rule, value) : undefined
    const univerEditor = univerEditorRef.current
    if (univerEditor) {
      setEditValue(value)
      void univerEditor
        .setCellValue(cell, value)
        .then(applied => {
          if (!applied) updateActiveSheet(upsertCell(activeSheet, cell, value))
        })
        .catch(() => {
          updateActiveSheet(upsertCell(activeSheet, cell, value))
        })
    } else {
      updateActiveSheet(upsertCell(activeSheet, cell, value))
    }
    if (validation && !validation.valid) {
      setStatusMessage(
        `${cell} flagged: ${validation.message ?? 'invalid value'}`,
      )
    }
  }

  function commitSelectedCell(value: string): void {
    if (formulaRangePickRef.current && /^=\w+\(\s*\)$/.test(value.trim())) {
      return
    }
    formulaRangePickRef.current = null
    if (formulaRangePickTimeoutRef.current) {
      window.clearTimeout(formulaRangePickTimeoutRef.current)
      formulaRangePickTimeoutRef.current = undefined
    }
    commitCellValue(selectedCell, value)
  }

  function formulaReferenceRange(
    range: string,
    cell = selectedCell,
    exactRange = false,
  ): string {
    const selectedPosition = parseCellKey(cell)
    const normalizedRange = normalizeRangeToken(range) ?? range
    const referencedCells = cellsInRange(normalizedRange)
    if (exactRange) return normalizedRange
    if (
      !selectedPosition ||
      referencedCells.length !== 1 ||
      referencedCells[0] !== cell
    ) {
      return normalizedRange
    }

    const startRow =
      selectedPosition.row < activeSheet.rowCount
        ? selectedPosition.row + 1
        : Math.max(1, selectedPosition.row - 10)
    const endRow =
      selectedPosition.row < activeSheet.rowCount
        ? Math.min(activeSheet.rowCount, selectedPosition.row + 10)
        : selectedPosition.row - 1

    return `${cellKey(startRow, selectedPosition.column)}:${cellKey(
      endRow,
      selectedPosition.column,
    )}`
  }

  function formulaTemplate(
    template: FormulaTemplate,
    range: string,
    cell = selectedCell,
    options: { exactRange?: boolean } = {},
  ): string {
    const referenceRange = formulaReferenceRange(
      range,
      cell,
      options.exactRange,
    )
    if (
      ['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'CONCAT'].includes(
        template,
      )
    ) {
      return `=${template}(${referenceRange})`
    }
    if (template === 'IF') return `=IF(${cell}>0, "Yes", "No")`
    if (template === 'ROUND') return `=ROUND(${cell}, 2)`
    if (template === 'LEFT') return `=LEFT(${cell}, 3)`
    if (template === 'RIGHT') return `=RIGHT(${cell}, 3)`
    if (template === 'VLOOKUP')
      return `=VLOOKUP(${cell}, ${referenceRange}, 2, FALSE)`
    if (template === 'COUNTIF') return `=COUNTIF(${referenceRange}, ">0")`
    if (template === 'SUMIF') return `=SUMIF(${referenceRange}, ">0")`
    if (template === 'AVERAGEIF')
      return `=AVERAGEIF(${referenceRange}, ">0")`
    if (template === 'RANK') return `=RANK(${cell}, ${referenceRange})`
    if (template === 'IFERROR') return `=IFERROR(${cell}, "")`
    if (template === 'TODAY') return '=TODAY()'
    if (template === 'NOW') return '=NOW()'
    if (template === 'ADD') return `=${cell}+0`
    if (template === 'SUBTRACT') return `=${cell}-0`
    if (template === 'MULTIPLY') return `=${cell}*1`
    return `=${cell}/1`
  }

  function insertFormulaTemplate(template: FormulaTemplate): void {
    const range = activeRange()
    if (!range) return
    const referencedCells = cellsInRange(range)
    if (
      ['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'CONCAT'].includes(
        template,
      ) &&
      referencedCells.length === 1 &&
      referencedCells[0] === selectedCell
    ) {
      formulaRangePickRef.current = { template, targetCell: selectedCell }
      setEditValue(`=${template}()`)
      setStatusMessage(`Select a range for ${template}`)
      window.setTimeout(() => {
        globalThis.document
          .querySelector<HTMLInputElement>(
            'input[aria-label="Formula or value"]',
          )
          ?.focus()
      }, 0)
      return
    }
    const formula = formulaTemplate(template, range)
    setEditValue(formula)
    commitSelectedCell(formula)
    setStatusMessage(`Inserted ${template} formula`)
  }

  function selectCell(cell: string): void {
    if (formulaRangePickRef.current) {
      setSelectedRange(`${cell}:${cell}`)
      return
    }
    setUniverSelection(null)
    manualRangeRef.current = false
    setSelectedCell(cell)
    setSelectedRange(`${cell}:${cell}`)
  }

  function changeSelectedRange(
    range: string,
    options: { manual?: boolean } = {},
  ): void {
    if (options.manual) manualRangeRef.current = true
    setSelectedRange(range)
  }

  function changeFormulaValue(value: string): void {
    formulaRangePickRef.current = null
    if (formulaRangePickTimeoutRef.current) {
      window.clearTimeout(formulaRangePickTimeoutRef.current)
      formulaRangePickTimeoutRef.current = undefined
    }
    setEditValue(value)
  }

  function undo(): void {
    // If the user is mid-typing in a cell (edit box open, not yet committed),
    // the first undo cancels that uncommitted edit instead of reaching back
    // into committed history — otherwise the half-typed cell lingers on the
    // canvas while an earlier cell gets cleared, which reads as a broken undo.
    const editor = univerEditorRef.current
    if (editor?.isCellEditing()) {
      void editor.abortEditing()
      setStatusMessage('Discarded unsaved cell edit')
      return
    }
    setUndoStack(stack => {
      const previous = stack.at(-1)
      if (!previous) return stack
      setRedoStack(redo => [document, ...redo].slice(0, 50))
      setDocument(previous)
      // Rebuild the Univer surface from the restored document — without this
      // the grid keeps showing the un-undone state and the next grid command
      // re-exports it, silently re-committing the undone edit.
      refreshUniverSurface()
      setSaveState('dirty')
      setStatusMessage('Undid last edit')
      return stack.slice(0, -1)
    })
  }

  function redo(): void {
    // Mirror undo: resolve any uncommitted in-cell edit before touching the
    // redo history, so redo always operates on committed state.
    const editor = univerEditorRef.current
    if (editor?.isCellEditing()) {
      void editor.abortEditing()
      setStatusMessage('Discarded unsaved cell edit')
      return
    }
    setRedoStack(stack => {
      const next = stack[0]
      if (!next) return stack
      setUndoStack(undoItems => [...undoItems, document].slice(-50))
      setDocument(next)
      refreshUniverSurface()
      setSaveState('dirty')
      setStatusMessage('Redid edit')
      return stack.slice(1)
    })
  }

  // Explicit save (⌘S, header Save): flush pending Univer edits into the
  // document first (the editor's snapshot export is debounced ~250ms, so the
  // freshest keystrokes may not have landed in React state yet), then flush
  // the debounced autosave through the unified lifecycle.
  async function save(): Promise<void> {
    const lifecycle = lifecycleRef.current
    setSaveState('saving')
    // Snapshot bridge: Univer → snapshot → .sheets document. Fall back to the
    // current React document when the surface is unavailable or the snapshot
    // fails to convert — never block an explicit save.
    let docToSave = documentRef.current
    try {
      const flushed = await univerEditorRef.current?.flushSnapshot()
      if (flushed) {
        docToSave = documentFromEditorSnapshot(flushed, documentRef.current)
        setDocument(docToSave)
      }
    } catch {
      docToSave = documentRef.current
    }
    if ((filePath.trim() || null) === boundPathRef.current) {
      boundDocRef.current = docToSave
    }
    const path = await lifecycle.ensureDraft()
    if (!path) {
      setSaveState('error')
      setStatusMessage(lifecycle.doc.error ?? 'Could not save the workbook.')
      return
    }
    await lifecycle.flush()
    setSaveState('saved')
    setStatusMessage(
      lifecycle.doc.status === 'draft'
        ? 'Draft saved — name it to file it.'
        : 'All changes saved.',
    )
    recordLedgerOperation({
      lane: 'user',
      kind: 'document.save',
      summary: `Saved workbook "${docToSave.metadata.title || 'Untitled'}"`,
      refs: { path },
    })
  }

  // New workbook (⌘N, switcher): detach and reset — the next edit lazily
  // creates a fresh draft in PureDrafts.
  function newWorkbook(fresh = createBlankWorkbook()): void {
    openGenerationRef.current += 1
    setSwitcherOpen(false)
    // Detach; the next edit lazily creates a fresh `.sheets` package draft.
    boundIsPackageRef.current = false
    setDocument(fresh)
    refreshUniverSurface()
    setUndoStack([])
    setRedoStack([])
    setSelectedCell('A1')
    manualRangeRef.current = false
    setSelectedRange('A1:A1')
    setTitleDraft(null)
    setSaveState('saved')
    setFilePath('')
    setStatusMessage('New workbook')
  }

  // Commit the workbook title (the naming surface): store it in the document
  // and rename the file when the workbook is already filed.
  function commitWorkbookTitle(value: string): void {
    setTitleDraft(null)
    const trimmed = value.trim() || 'Untitled'
    if (trimmed !== document.metadata.title) {
      commitDocument(current => ({
        ...current,
        metadata: { ...current.metadata, title: trimmed },
      }))
    }
    // Only packages rename via the title. A legacy flat `.sheets` /
    // `.sheets.html` file keeps its on-disk name (renaming would fight the
    // shared/compound suffix); its title still persists into the workbook.
    const lifecycle = lifecycleRef.current
    if (
      boundIsPackageRef.current &&
      lifecycle.doc.status === 'filed' &&
      trimmed !== lifecycle.doc.title
    ) {
      void lifecycle.rename(trimmed)
    }
  }

  useDocumentHotkeys({
    onSave: () => void save(),
    onNew: newWorkbook,
    onOpen: () => setSwitcherOpen(true),
  })


  // ⌘/Ctrl+D fill down, ⌘/Ctrl+R fill right. Skips text-entry targets and
  // events another handler (e.g. the Univer editor's own shortcuts) already
  // claimed via preventDefault.
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent): void => {
      if (!(event.metaKey || event.ctrlKey) || event.altKey || event.shiftKey)
        return
      const key = event.key.toLowerCase()
      if (key !== 'd' && key !== 'r') return
      if (event.defaultPrevented) return
      const target = event.target as HTMLElement | null
      if (
        target &&
        (target.tagName === 'INPUT' ||
          target.tagName === 'TEXTAREA' ||
          target.isContentEditable)
      ) {
        return
      }
      event.preventDefault()
      fillSelectedRange(key === 'd' ? 'down' : 'right')
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  })

  function toggleStyle(style: 'bold' | 'italic' | 'underline'): void {
    const range = activeRange()
    if (!range) return
    const current = Boolean(activeSheet.cells[selectedCell]?.style?.[style])
    applyRangeStyle(range, { [style]: !current }, `Updated ${range} formatting`)
  }

  function applyCellStyle(
    patch: Partial<CellStyle>,
    message = 'Updated formatting',
  ): void {
    const range = activeRange()
    if (!range) return
    applyRangeStyle(range, patch, message)
  }

  // Decimal increase/decrease. Applies to numeric formats; a cell without a
  // numeric format is promoted to 'number' so the control always has effect.
  function adjustDecimals(delta: 1 | -1): void {
    const format =
      selectedStyle?.numberFormat === 'currency' ||
      selectedStyle?.numberFormat === 'percent' ||
      selectedStyle?.numberFormat === 'number'
        ? selectedStyle.numberFormat
        : 'number'
    const current = selectedStyle?.decimals ?? 2
    const next = Math.min(8, Math.max(0, current + delta))
    applyCellStyle(
      { numberFormat: format, decimals: next },
      `Set ${next} decimal place${next === 1 ? '' : 's'}`,
    )
  }

  // Table-style presets are multi-cell style transforms — like paste-special
  // they persist through the document → fresh-snapshot → remount path.
  function applyTableStyle(preset: TableStylePreset): void {
    const range = activeRange()
    if (!range) return
    const label =
      preset === 'header-row'
        ? 'header row style'
        : preset === 'banded-rows'
        ? 'banded rows'
        : preset === 'totals-row'
        ? 'totals row style'
        : 'grid borders'
    updateActiveSheetFromCurrent(
      sheet => applyTableStylePreset(sheet, range, preset),
      `Applied ${label} to ${range}`,
    )
    refreshUniverSurface()
  }

  function applyRangeStyle(
    range: string,
    patch: Partial<CellStyle>,
    message: string,
  ): void {
    const univerEditor = univerEditorRef.current
    if (univerEditor) {
      setStatusMessage(message)
      // When the bridge refuses or throws, the style lands in the model only
      // — without a surface reload it would be invisible on the canvas until
      // the next unrelated refresh (this is exactly how align-right "did
      // nothing" while the facade rejected 'right', #270).
      const modelFallback = (): void => {
        updateActiveSheetFromCurrent(
          sheet => setRangeStyle(sheet, range, patch),
          message,
        )
        refreshUniverSurface()
      }
      void univerEditor
        .applyRangeStyle(range, patch)
        .then(applied => {
          if (!applied) modelFallback()
        })
        .catch(modelFallback)
      return
    }
    updateActiveSheetFromCurrent(
      sheet => setRangeStyle(sheet, range, patch),
      message,
    )
  }

  function addSheet(): void {
    const name = `Sheet ${document.workbook.sheets.length + 1}`
    const message = 'Added sheet'
    const record = () =>
      recordLedgerOperation({
        lane: 'user',
        kind: 'sheets.sheet.add',
        summary: `Added sheet "${name}"`,
      })
    const fallback = () =>
      commitDocument(current => addWorkbookSheet(current, name), message)
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      selectCell('A1')
      record()
      return
    }
    setStatusMessage(message)
    void editor
      .addSheet(name)
      .then(applied => {
        if (!applied) fallback()
        selectCell('A1')
        record()
      })
      .catch(() => {
        fallback()
        selectCell('A1')
        record()
      })
  }

  function selectSheet(sheetId: string): void {
    const fallback = () => {
      setDocument(current => ({
        ...current,
        workbook: {
          ...current.workbook,
          activeSheetId: sheetId,
        },
      }))
    }
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      selectCell('A1')
      return
    }
    void editor
      .selectSheet(sheetId)
      .then(applied => {
        if (!applied) fallback()
        selectCell('A1')
      })
      .catch(() => {
        fallback()
        selectCell('A1')
      })
  }

  function renameCurrentSheet(): void {
    beginRenameSheet(activeSheet.id)
  }

  function beginRenameSheet(sheetId: string): void {
    const sheet = document.workbook.sheets.find(s => s.id === sheetId)
    if (!sheet) return
    setRenamingSheetId(sheetId)
    setRenameDraft(sheet.name)
  }

  function cancelSheetRename(): void {
    setRenamingSheetId(null)
    setRenameDraft('')
  }

  function commitSheetRename(): void {
    if (!renamingSheetId) return
    const nextName = renameDraft.trim()
    const sheetId = renamingSheetId
    cancelSheetRename()
    if (!nextName) return
    const message = 'Renamed sheet'
    const record = () =>
      recordLedgerOperation({
        lane: 'user',
        kind: 'sheets.sheet.rename',
        summary: `Renamed a sheet to "${nextName}"`,
      })
    const fallback = () =>
      commitDocument(
        current => renameWorkbookSheet(current, sheetId, nextName),
        message,
      )
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      record()
      return
    }
    setStatusMessage(message)
    void editor
      .renameSheet(sheetId, nextName)
      .then(applied => {
        if (!applied) fallback()
        record()
      })
      .catch(() => {
        fallback()
        record()
      })
  }

  function duplicateCurrentSheet(): void {
    const message = 'Duplicated sheet'
    const fallback = () =>
      commitDocument(
        current => duplicateWorkbookSheet(current, activeSheet.id),
        message,
      )
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      selectCell('A1')
      return
    }
    setStatusMessage(message)
    void editor
      .duplicateActiveSheet()
      .then(applied => {
        if (!applied) fallback()
        selectCell('A1')
      })
      .catch(() => {
        fallback()
        selectCell('A1')
      })
  }

  function deleteCurrentSheet(): void {
    if (document.workbook.sheets.length <= 1) return
    const message = 'Deleted sheet'
    const deletedName = activeSheet.name
    const record = () =>
      recordLedgerOperation({
        lane: 'user',
        kind: 'sheets.sheet.delete',
        summary: `Deleted sheet "${deletedName}"`,
      })
    const fallback = () =>
      commitDocument(
        current => deleteWorkbookSheet(current, activeSheet.id),
        message,
      )
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      selectCell('A1')
      record()
      return
    }
    setStatusMessage(message)
    void editor
      .deleteSheet(activeSheet.id)
      .then(applied => {
        if (!applied) fallback()
        selectCell('A1')
        record()
      })
      .catch(() => {
        fallback()
        selectCell('A1')
        record()
      })
  }

  function pasteIntoCell(cell: string, text: string): void {
    if (isCellLocked(activeSheet, cell)) {
      setStatusMessage(`${cell} is protected — unlock the range to paste`)
      return
    }
    selectCell(cell)
    const message = 'Pasted tabular data'
    const fallback = () =>
      updateActiveSheetFromCurrent(
        sheet => pasteTabularData(sheet, cell, text),
        message,
      )
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      return
    }
    setStatusMessage(message)
    void editor
      .pasteTabularData(cell, text)
      .then(applied => {
        if (!applied) fallback()
      })
      .catch(fallback)
  }

  function clearSelectedRange(): void {
    const range = activeRange()
    if (!range) return
    // Ranges touching protected cells go through the model (which skips the
    // locked cells) instead of Univer's clear, which knows nothing of locks.
    if (isRangeLocked(activeSheet, range)) {
      updateActiveSheetFromCurrent(
        sheet => clearRange(sheet, range),
        `Cleared unprotected cells in ${range}`,
      )
      refreshUniverSurface()
      return
    }
    const univerEditor = univerEditorRef.current
    if (univerEditor) {
      const message = `Cleared ${range}`
      setStatusMessage(message)
      void univerEditor
        .clearRange(range)
        .then(applied => {
          if (!applied)
            updateActiveSheetFromCurrent(
              sheet => clearRange(sheet, range),
              message,
            )
        })
        .catch(() => {
          updateActiveSheetFromCurrent(
            sheet => clearRange(sheet, range),
            message,
          )
        })
      return
    }
    updateActiveSheetFromCurrent(
      sheet => clearRange(sheet, range),
      `Cleared ${range}`,
    )
  }

  async function copySelectedRange(cut = false): Promise<void> {
    const range = activeRange()
    if (!range) return
    const text = rangeToDelimitedText(activeSheet, range)
    // Capture the rich block BEFORE a cut clears the source cells.
    const block = copyRangeCells(activeSheet, range)
    try {
      await writeClipboard(text)
      copiedBlockRef.current = block
      copiedTextRef.current = text
      setStatusMessage(`${cut ? 'Cut' : 'Copied'} ${range}`)
      showToast(`${cut ? 'Cut' : 'Copied'} ${range}`)
      if (cut) {
        const editor = univerEditorRef.current
        if (editor) {
          void editor
            .clearRange(range)
            .then(applied => {
              if (!applied)
                updateActiveSheetFromCurrent(
                  sheet => clearRange(sheet, range),
                  `Cut ${range}`,
                )
            })
            .catch(() =>
              updateActiveSheetFromCurrent(
                sheet => clearRange(sheet, range),
                `Cut ${range}`,
              ),
            )
        } else {
          updateActiveSheetFromCurrent(
            sheet => clearRange(sheet, range),
            `Cut ${range}`,
          )
        }
      }
    } catch {
      setStatusMessage('Clipboard is not available in this window')
    }
  }

  // A canvas copy/cut (Ctrl+C / Ctrl+X) is handled by Univer and never runs our
  // Copy button, so mirror the selection into the internal clipboard here — that
  // is what the paste-values / paste-formatting buttons read (they can't parse
  // styles back out of the plain-text system clipboard).
  function captureEditorClipboard(): void {
    const range = normalizeRangeToken(selectedRange)
    if (!range) return
    copiedBlockRef.current = copyRangeCells(activeSheet, range)
    copiedTextRef.current = rangeToDelimitedText(activeSheet, range)
  }

  async function pasteFromClipboard(cell = selectedCell): Promise<void> {
    try {
      const text = await readClipboard()
      if (!text) {
        setStatusMessage('Clipboard is empty')
        return
      }
      // In-app paste: when the system clipboard still holds the TSV written
      // by our last copy, paste the rich block instead — formulas shift
      // relative to the new anchor and styles travel with the cells. TSV
      // alone would flatten formulas into their display values.
      if (copiedBlockRef.current && copiedTextRef.current === text) {
        pasteBlockAt(cell, copiedBlockRef.current, 'all', `Pasted into ${cell}`)
        return
      }
      pasteIntoCell(cell, text)
    } catch {
      setStatusMessage('Clipboard is not available in this window')
    }
  }

  // Paste a copied block through the document model, then rebuild the Univer
  // surface from the updated document. Styles cannot travel through the
  // value-only bridge calls, so the snapshot bridge (document → fresh
  // snapshot → remount) is the single persistence path here — the same one
  // undo/redo uses.
  function pasteBlockAt(
    cell: string,
    block: CopiedCellBlock,
    mode: PasteMode,
    message: string,
  ): void {
    selectCell(cell)
    updateActiveSheetFromCurrent(
      sheet => pasteCellBlock(sheet, cell, block, mode),
      message,
    )
    refreshUniverSurface()
  }

  function pasteSpecial(mode: Exclude<PasteMode, 'all'>): void {
    const block = copiedBlockRef.current
    if (!block) {
      setStatusMessage('Copy a range first, then paste special')
      return
    }
    pasteBlockAt(
      selectedCell,
      block,
      mode,
      mode === 'values'
        ? `Pasted values into ${selectedCell}`
        : `Pasted formatting into ${selectedCell}`,
    )
  }

  // Fill down/right: extend the first row/column of the selection across the
  // rest of it (numeric series continue; formulas shift relatively).
  function fillSelectedRange(direction: 'down' | 'right'): void {
    const range = activeRange()
    if (!range) return
    const [startToken, endToken] = range.split(':')
    const start = parseCellKey(startToken ?? '')
    const end = parseCellKey(endToken ?? startToken ?? '')
    if (!start || !end) return
    if (direction === 'down' && end.row <= start.row) {
      setStatusMessage('Select at least two rows to fill down')
      return
    }
    if (direction === 'right' && end.column <= start.column) {
      setStatusMessage('Select at least two columns to fill right')
      return
    }
    const sourceToken =
      direction === 'down'
        ? `${cellKey(start.row, start.column)}:${cellKey(
            start.row,
            end.column,
          )}`
        : `${cellKey(start.row, start.column)}:${cellKey(
            end.row,
            start.column,
          )}`
    const targetToken =
      direction === 'down'
        ? `${cellKey(start.row + 1, start.column)}:${cellKey(
            end.row,
            end.column,
          )}`
        : `${cellKey(start.row, start.column + 1)}:${cellKey(
            end.row,
            end.column,
          )}`
    const message = `Filled ${direction} ${range}`
    updateActiveSheetFromCurrent(
      sheet => fillRange(sheet, sourceToken, targetToken),
      message,
    )
    refreshUniverSurface()
  }

  function insertSelectedRow(): void {
    if (!selectedPosition) return
    const row = selectedPosition.row
    const message = `Inserted row ${row + 1}`
    const fallback = () =>
      updateActiveSheetFromCurrent(sheet => insertRow(sheet, row), message)
    const editor = univerEditorRef.current
    // Sheets with positional metadata (comments, validation, protection,
    // conditional formats) run structural edits model-first: only the model
    // shifts that metadata with the cells (S3 gap closure).
    if (!editor || sheetHasPositionalMetadata(activeSheet)) {
      fallback()
      if (editor) refreshUniverSurface()
      return
    }
    setStatusMessage(message)
    void editor
      .insertRow(row)
      .then(applied => {
        if (!applied) fallback()
      })
      .catch(fallback)
  }

  function deleteSelectedRow(): void {
    if (!selectedPosition) return
    const row = selectedPosition.row
    const nextRow = Math.min(row, Math.max(0, activeSheet.rowCount - 2))
    const nextCell = cellKey(nextRow, selectedPosition.column)
    const message = `Deleted row ${row + 1}`
    const fallback = () => {
      updateActiveSheetFromCurrent(sheet => deleteRow(sheet, row), message)
      selectCell(nextCell)
    }
    const editor = univerEditorRef.current
    if (!editor || sheetHasPositionalMetadata(activeSheet)) {
      fallback()
      if (editor) refreshUniverSurface()
      return
    }
    setStatusMessage(message)
    void editor
      .deleteRow(row)
      .then(applied => {
        if (!applied) fallback()
        else selectCell(nextCell)
      })
      .catch(fallback)
  }

  function insertSelectedColumn(): void {
    if (!selectedPosition) return
    const column = selectedPosition.column
    const message = `Inserted column ${columnLabel(column)}`
    const fallback = () =>
      updateActiveSheetFromCurrent(
        sheet => insertColumn(sheet, column),
        message,
      )
    const editor = univerEditorRef.current
    if (!editor || sheetHasPositionalMetadata(activeSheet)) {
      fallback()
      if (editor) refreshUniverSurface()
      return
    }
    setStatusMessage(message)
    void editor
      .insertColumn(column)
      .then(applied => {
        if (!applied) fallback()
      })
      .catch(fallback)
  }

  function deleteSelectedColumn(): void {
    if (!selectedPosition) return
    const column = selectedPosition.column
    const nextColumn = Math.min(
      column,
      Math.max(0, activeSheet.columnCount - 2),
    )
    const nextCell = cellKey(selectedPosition.row, nextColumn)
    const message = `Deleted column ${columnLabel(column)}`
    const fallback = () => {
      updateActiveSheetFromCurrent(
        sheet => deleteColumn(sheet, column),
        message,
      )
      selectCell(nextCell)
    }
    const editor = univerEditorRef.current
    if (!editor || sheetHasPositionalMetadata(activeSheet)) {
      fallback()
      if (editor) refreshUniverSurface()
      return
    }
    setStatusMessage(message)
    void editor
      .deleteColumn(column)
      .then(applied => {
        if (!applied) fallback()
        else selectCell(nextCell)
      })
      .catch(fallback)
  }

  function toggleFirstRowFreeze(): void {
    const nextFrozenRows = activeSheet.frozenRows ? 0 : 1
    const message = 'Toggled frozen row'
    const fallback = () =>
      updateActiveSheetFromCurrent(toggleFreezeFirstRow, message)
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      return
    }
    setStatusMessage(message)
    void editor
      .setFrozenRows(nextFrozenRows)
      .then(applied => {
        if (!applied) fallback()
      })
      .catch(fallback)
  }

  function toggleFirstColumnFreeze(): void {
    const nextFrozenColumns = activeSheet.frozenColumns ? 0 : 1
    const message = 'Toggled frozen column'
    const fallback = () =>
      updateActiveSheetFromCurrent(toggleFreezeFirstColumn, message)
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      return
    }
    setStatusMessage(message)
    void editor
      .setFrozenColumns(nextFrozenColumns)
      .then(applied => {
        if (!applied) fallback()
      })
      .catch(fallback)
  }

  function activeRange(): string | null {
    const normalized = normalizeRangeToken(selectedRange)
    if (!normalized) setStatusMessage(`Invalid range ${selectedRange}`)
    return normalized
  }

  function sheetRangeValues(
    sheet: WorkbookSheet,
    range: string,
  ): UniverPasteValue[][] {
    const keys = cellsInRange(range)
    const rows: UniverPasteValue[][] = []
    let currentRow = -1
    for (const key of keys) {
      const position = parseCellKey(key)
      if (!position) continue
      if (position.row !== currentRow) {
        rows.push([])
        currentRow = position.row
      }
      const value = sheet.cells[key]?.value ?? ''
      rows[rows.length - 1].push(
        value.trim().startsWith('=') ? { f: value } : value,
      )
    }
    return rows
  }

  // The sorted block as a typed Univer value matrix: numbers stay numbers,
  // formulas ride as `{ f }`, and blanks are `null` — NOT `''` — so a blank
  // landing on a previously filled cell truly clears it in the canvas (#275).
  function sortedRangeCellValues(
    sheet: WorkbookSheet,
    range: string,
  ): UniverPasteValue[][] {
    const keys = cellsInRange(range)
    const rows: UniverPasteValue[][] = []
    let currentRow = -1
    for (const key of keys) {
      const position = parseCellKey(key)
      if (!position) continue
      if (position.row !== currentRow) {
        rows.push([])
        currentRow = position.row
      }
      const cell = sheet.cells[key]
      let value: UniverPasteValue = null
      if (cell && cell.value !== '') {
        if (cell.kind === 'formula') value = { f: cell.value }
        else if (cell.kind === 'number' && Number.isFinite(Number(cell.value)))
          value = Number(cell.value)
        else value = cell.value
      }
      rows[rows.length - 1].push(value)
    }
    return rows
  }

  async function sortSelectedRange(direction: 'asc' | 'desc'): Promise<void> {
    const range = activeRange()
    if (!range) return
    if (isRangeLocked(activeSheet, range)) {
      setStatusMessage(`${range} contains protected cells — sort refused`)
      return
    }
    const message = `Sorted ${range}`
    // No canvas (tests, mount failure) or the bridge refused the write: sort
    // the model and rebuild the surface from it.
    const modelFallback = (): void => {
      updateActiveSheetFromCurrent(
        sheet => sortRange(sheet, range, direction),
        message,
      )
      refreshUniverSurface()
    }
    const editor = univerEditorRef.current
    if (!editor) {
      modelFallback()
      return
    }
    // Canvas path: flush the live canvas into a document first so the sort
    // reads exactly what the user sees (the debounced model sync can trail
    // recent edits), then write the sorted block back through the bridge —
    // the same path replace and paste use. The canvas never reloads, so the
    // selection, scroll position and rendered formatting stay put; the old
    // full-surface reload reset the selection to A1 and flashed the grid on
    // every sort (#275).
    const flushed = await editor.flushSnapshot().catch(() => null)
    const sourceDocument = flushed
      ? documentFromEditorSnapshot(flushed, documentRef.current)
      : documentRef.current
    const sourceSheet =
      sourceDocument.workbook.sheets.find(
        sheet => sheet.id === activeSheet.id,
      ) ?? activeSheet
    const nextSheet = sortRange(sourceSheet, range, direction)
    if (nextSheet === sourceSheet) return
    void editor
      .setRangeValues(range, sortedRangeCellValues(nextSheet, range))
      .then(applied => {
        if (applied) setStatusMessage(message)
        else modelFallback()
      })
      .catch(modelFallback)
  }

  function createChart(type: WorkbookChart['type']): void {
    const range = activeRange()
    if (!range) return
    const nextSheet = createChartFromRange(activeSheet, range, type)
    if (nextSheet === activeSheet) {
      setStatusMessage(`No numeric chart data found in ${range}`)
      return
    }
    updateActiveSheet(nextSheet, `Created ${type} chart from ${range}`)
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.chart.create',
      summary: `Created a ${type} chart from ${range} on ${activeSheet.name}`,
    })
  }

  function deleteChart(chartId: string): void {
    updateActiveSheet(removeChart(activeSheet, chartId), 'Removed chart')
  }

  function updateChartConfig(chartId: string, patch: ChartUpdatePatch): void {
    const next = updateChart(activeSheet, chartId, patch)
    if (next === activeSheet) {
      setStatusMessage('No plottable data in that range')
      return
    }
    updateActiveSheet(next, 'Updated chart')
  }

  // ---- Import / export (Phase S4) -------------------------------------------

  function uniqueSheetName(base: string): string {
    const names = new Set(document.workbook.sheets.map(sheet => sheet.name))
    if (!names.has(base)) return base
    let suffix = 2
    while (names.has(`${base} ${suffix}`)) suffix += 1
    return `${base} ${suffix}`
  }

  function handleImportExport(action: ImportExportAction): void {
    const run =
      action === 'import-csv'
        ? importCsvFile
        : action === 'export-csv'
        ? exportCsvFile
        : action === 'import-xlsx'
        ? importXlsxFile
        : exportXlsxFile
    void run().catch(error => {
      setStatusMessage(error instanceof Error ? error.message : String(error))
    })
  }

  async function importCsvFile(): Promise<void> {
    const path = await openImportFile()
    if (!path) return
    const text = await readTextFile(path)
    const base =
      fileNameFromPath(path)
        .replace(/\.[^.]+$/, '')
        .trim() || 'Imported'
    const sheet = sheetFromCsv(
      text,
      `sheet-csv-${Date.now()}`,
      uniqueSheetName(base),
    )
    commitDocument(
      current => addSheetToWorkbook(current, sheet),
      `Imported ${fileNameFromPath(path)} as sheet "${sheet.name}"`,
    )
    refreshUniverSurface()
    selectCell('A1')
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.import.csv',
      summary: `Imported CSV ${fileNameFromPath(path)}`,
      refs: { path },
    })
  }

  async function exportCsvFile(): Promise<void> {
    const path = await chooseSaveFilePath({
      defaultName: `${activeSheet.name || 'Sheet'}.csv`,
      filters: [{ name: 'CSV', extensions: ['csv'] }],
    })
    if (!path) return
    await writeTextFile(path, csvFromSheet(activeSheet))
    setStatusMessage(
      `Exported ${activeSheet.name} to ${fileNameFromPath(path)}`,
    )
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.export.csv',
      summary: `Exported sheet ${activeSheet.name} as CSV`,
      refs: { path },
    })
  }

  async function importXlsxFile(): Promise<void> {
    const path = await openImportFile()
    if (!path) return
    if (!/\.xlsx$/i.test(path)) {
      setStatusMessage('Choose an .xlsx file to import')
      return
    }
    setStatusMessage(`Importing ${fileNameFromPath(path)}…`)
    const binary = await readBinaryFile(path, 30_000_000)
    if (binary.truncated || !binary.base64) {
      setStatusMessage(
        'That workbook is too large to import (30 MB limit) or unreadable',
      )
      return
    }
    const imported = await importXlsxWorkbook(
      base64ToBytes(binary.base64),
      fileNameFromPath(path).replace(/\.xlsx$/i, '') || 'Imported workbook',
    )
    // Imported content becomes a NEW workbook: detach from the current file
    // so the next save lazily creates a fresh `.sheets` package draft.
    boundIsPackageRef.current = false
    setDocument(imported.document)
    setUndoStack([])
    setRedoStack([])
    setFilePath('')
    setSelectedCell('A1')
    manualRangeRef.current = false
    setSelectedRange('A1:A1')
    setTitleDraft(null)
    setSaveState('dirty')
    setImportWarnings(imported.warnings)
    setStatusMessage(`Imported ${fileNameFromPath(path)}`)
    refreshUniverSurface()
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.import.xlsx',
      summary: `Imported XLSX workbook ${fileNameFromPath(path)}`,
      refs: { path },
    })
  }

  async function exportXlsxFile(): Promise<void> {
    const path = await chooseSaveFilePath({
      defaultName: `${document.metadata.title || 'Workbook'}.xlsx`,
      filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }],
    })
    if (!path) return
    // Flush pending Univer edits so the export matches what is on screen.
    const flushed = await univerEditorRef.current
      ?.flushSnapshot()
      .catch(() => null)
    const docToExport = flushed
      ? documentFromEditorSnapshot(flushed, documentRef.current)
      : documentRef.current
    const bytes = await exportXlsxWorkbook(docToExport)
    await writeBinaryFile(path, bytesToBase64(bytes))
    setStatusMessage(`Exported workbook to ${fileNameFromPath(path)}`)
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.export.xlsx',
      summary: `Exported workbook "${docToExport.metadata.title}" as XLSX`,
      refs: { path },
    })
  }

  function resizeSelectedColumn(width: number): void {
    if (!selectedPosition) return
    const column = selectedPosition.column
    const message = `Set ${columnLabel(column)} width`
    const fallback = () =>
      updateActiveSheet(setColumnWidth(activeSheet, column, width), message)
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      return
    }
    setStatusMessage(message)
    void editor
      .setColumnWidth(column, width)
      .then(applied => {
        if (!applied) fallback()
      })
      .catch(fallback)
  }

  function resizeSelectedRow(height: number): void {
    if (!selectedPosition) return
    const row = selectedPosition.row
    const message = `Set row ${row + 1} height`
    const fallback = () =>
      updateActiveSheet(setRowHeight(activeSheet, row, height), message)
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      return
    }
    setStatusMessage(message)
    void editor
      .setRowHeight(row, height)
      .then(applied => {
        if (!applied) fallback()
      })
      .catch(fallback)
  }

  function selectedColumnWidth(): number {
    if (!selectedPosition) return 112
    return activeSheet.columnWidths?.[String(selectedPosition.column)] ?? 112
  }

  function selectedRowHeight(): number {
    if (!selectedPosition) return 32
    return activeSheet.rowHeights?.[String(selectedPosition.row)] ?? 32
  }

  // Commit any open in-cell edit and pull Univer's freshest data, so the cell
  // you're currently typing in counts in Find/Replace (#277). Returns the
  // up-to-date active sheet; the edit's snapshot export is otherwise debounced,
  // so the just-typed value wouldn't be in the model yet.
  async function committedActiveSheet(): Promise<WorkbookSheet> {
    const editor = univerEditorRef.current
    if (!editor) return activeSheet
    if (editor.isCellEditing?.()) {
      await editor.commitEditing?.().catch(() => undefined)
    }
    const flushed = await editor.flushSnapshot?.().catch(() => null)
    if (!flushed) return activeSheet
    try {
      const doc = documentFromEditorSnapshot(flushed, documentRef.current)
      setDocument(doc)
      setSaveState('dirty')
      return (
        doc.workbook.sheets.find(sheet => sheet.id === activeSheet.id) ??
        activeSheet
      )
    } catch {
      return activeSheet
    }
  }

  async function findNext(): Promise<void> {
    const query = findText.trim()
    if (!query) return
    const sheet = await committedActiveSheet()
    const matches = findMatches(sheet, query, { matchCase })
    if (!matches.length) {
      setStatusMessage(`No matches for ${findText}`)
      return
    }
    const currentIndex = matches.indexOf(selectedCell)
    // Land on the cell you're sitting on if it matches and Find didn't just
    // report it — so the current cell is counted, not skipped (#277). A repeat
    // Find advances to the next match.
    const landOnCurrent =
      currentIndex >= 0 && selectedCell !== lastFindRef.current
    const nextIndex = landOnCurrent
      ? currentIndex
      : (currentIndex + 1) % matches.length
    const next = matches[nextIndex]
    lastFindRef.current = next
    selectCell(next)
    // Highlight the match on the live grid and scroll it into view — selecting
    // in React alone doesn't move the Univer canvas (#277).
    void univerEditorRef.current?.revealCell?.(next)
    setStatusMessage(`Found ${next} (${nextIndex + 1}/${matches.length})`)
  }

  // Replace ONE occurrence at a time: the first match in the current cell, then
  // (once that cell has no matches left) the next matching cell (#277).
  async function replaceNext(): Promise<void> {
    const query = findText.trim()
    if (!query) return
    const sheet = await committedActiveSheet()
    const matches = findMatches(sheet, query, { matchCase })
    if (!matches.length) {
      setStatusMessage(`No matches for ${findText}`)
      return
    }
    const currentIndex = matches.indexOf(selectedCell)
    // Replace in the currently selected match first; otherwise the next one.
    const target =
      currentIndex >= 0
        ? matches[currentIndex]
        : matches[(currentIndex + 1) % matches.length]
    if (isCellLocked(sheet, target)) {
      selectCell(target)
      setStatusMessage(`${target} is protected — skipped`)
      return
    }
    const preview = replaceFirstOccurrence(sheet, target, query, replaceText, {
      matchCase,
    })
    if (!preview.replaced) {
      setStatusMessage(`No replaceable match in ${target}`)
      return
    }
    const nextValue = preview.sheet.cells[target]?.value ?? ''
    selectCell(target)
    commitCellValue(target, nextValue)
    // Reported this cell; a following Find should move past it.
    lastFindRef.current = target
    setStatusMessage(`Replaced one match in ${target}`)
  }

  async function replaceMatches(): Promise<void> {
    const query = findText.trim()
    const range = activeRange()
    if (!query || !range) return
    const sheet = await committedActiveSheet()
    const preview = replaceInRange(sheet, range, query, replaceText, {
      matchCase,
    })
    if (!preview.count) {
      setStatusMessage(`No matches for ${findText} in ${range}`)
      return
    }
    const message = `Replaced ${preview.count} match${
      preview.count === 1 ? '' : 'es'
    } in ${range}`
    setStatusMessage(message)
    const editor = univerEditorRef.current
    if (editor) {
      const values = sheetRangeValues(preview.sheet, range)
      void editor
        .setRangeValues(range, values)
        .then(applied => {
          if (!applied) {
            updateActiveSheetFromCurrent(
              sheet =>
                replaceInRange(sheet, range, query, replaceText, { matchCase })
                  .sheet,
              message,
            )
          }
        })
        .catch(() => {
          updateActiveSheetFromCurrent(
            sheet =>
              replaceInRange(sheet, range, query, replaceText, { matchCase })
                .sheet,
            message,
          )
        })
      return
    }
    updateActiveSheetFromCurrent(
      sheet =>
        replaceInRange(sheet, range, query, replaceText, { matchCase }).sheet,
      message,
    )
  }

  async function replaceAllMatches(): Promise<void> {
    const query = findText.trim()
    if (!query) return
    const sheet = await committedActiveSheet()
    const preview = replaceAllInSheet(sheet, query, replaceText, {
      matchCase,
    })
    if (!preview.count) {
      setStatusMessage(`No matches for ${findText}`)
      return
    }
    const message = `Replaced ${preview.count} match${
      preview.count === 1 ? '' : 'es'
    } in ${sheet.name}`
    updateActiveSheet(preview.sheet, message)
    refreshUniverSurface()
  }

  // ---- Phase S3 handlers ----------------------------------------------------

  function saveCellComment(text: string): void {
    // Apply to every cell in the selection: a single cell → just that cell,
    // a range → the whole range (each cell carries the same comment).
    const range = normalizeRangeToken(selectedRange)
    const cells = range ? cellsInRange(range) : [selectedCell]
    const label = cells.length > 1 ? selectedRange : selectedCell
    const trimmed = text.trim()
    updateActiveSheetFromCurrent(
      sheet =>
        cells.reduce((next, cell) => setCellComment(next, cell, text), sheet),
      trimmed ? `Saved comment on ${label}` : `Removed comment from ${label}`,
    )
  }

  function lockSelectedRange(): void {
    const range = activeRange()
    if (!range) return
    updateActiveSheetFromCurrent(
      sheet => lockRange(sheet, range),
      `Protected ${range}`,
    )
    // Remount the Univer surface so canvas-level range protection (best
    // effort, S5) picks up the new locks immediately.
    refreshUniverSurface()
  }

  function unlockSelectedRange(): void {
    const range = activeRange()
    if (!range) return
    updateActiveSheetFromCurrent(
      sheet => unlockRange(sheet, range),
      `Unprotected ${range}`,
    )
    refreshUniverSurface()
  }

  function commitFilterQuery(query: string): void {
    updateActiveSheetFromCurrent(
      sheet => setFilterQuery(sheet, query),
      query.trim()
        ? `Filtering ${columnLabel(
            activeSheet.filterColumn ?? 0,
          )} by "${query.trim()}"`
        : 'Cleared filter query',
    )
    applyFilterToEditor(setFilterQuery(activeSheet, query))
  }

  // Reflect the sheet's filter on the live Univer editor by hiding the rows
  // that `visibleRowsForSheet` excludes. The React filter state is the source
  // of truth; this just mirrors it onto the canvas so 'contains' etc. actually
  // hide rows (they previously only affected the legacy grid) (#274).
  function applyFilterToEditor(sheet: WorkbookSheet): void {
    const editor = univerEditorRef.current
    if (!editor) return
    // Only hide rows within the DATA extent — the last row that actually holds
    // content. Collapsing the ~1000 trailing empty rows is pointless and, worse,
    // pushes Univer's selection past the grid (row 1000) and crashes the
    // selection sync (#274).
    let lastDataRow = 0
    for (const key of Object.keys(sheet.cells)) {
      const pos = parseCellKey(key)
      if (pos && pos.row > lastDataRow) lastDataRow = pos.row
    }
    const visible = new Set(visibleRowsForSheet(sheet))
    const hidden: number[] = []
    for (let row = 0; row <= lastDataRow; row += 1) {
      if (!visible.has(row)) hidden.push(row)
    }
    void editor.applyRowVisibility(hidden, sheet.rowCount)
  }

  function toggleFilterHasHeader(): void {
    const next = setFilterHasHeader(
      activeSheet,
      activeSheet.filterHasHeader === false,
    )
    updateActiveSheetFromCurrent(
      sheet => setFilterHasHeader(sheet, sheet.filterHasHeader === false),
      next.filterHasHeader === false
        ? 'Filtering without a header row'
        : 'Keeping the header row visible',
    )
    applyFilterToEditor(next)
  }

  function applyValidation(
    kind: ValidationKind,
    options: { min?: number; max?: number },
  ): void {
    const range = activeRange()
    if (!range) return
    let rule: Parameters<typeof addValidationRule>[1]
    if (kind === 'list') {
      const listOptions = Array.from(
        new Set(
          cellsInRange(range)
            .map(key => computedCellValue(activeSheet, key).trim())
            .filter(Boolean),
        ),
      )
      if (!listOptions.length) {
        setStatusMessage('Select cells with values to build the allowed list')
        return
      }
      rule = { range, kind: 'list', options: listOptions }
    } else if (kind === 'number-range') {
      rule = { range, kind, min: options.min, max: options.max }
    } else {
      rule = { range, kind: 'date' }
    }
    updateActiveSheetFromCurrent(
      sheet => addValidationRule(sheet, rule),
      `Added ${kind} validation for ${range}`,
    )
  }

  function clearValidation(): void {
    const range = activeRange()
    if (!range) return
    updateActiveSheetFromCurrent(
      sheet => clearValidationInRange(sheet, range),
      `Cleared validation rules in ${range}`,
    )
  }

  function applyConditionalFormat(
    condition: ConditionalFormatRule['condition'],
    style: CellStyle,
  ): void {
    const range = activeRange()
    if (!range) return
    updateActiveSheetFromCurrent(
      sheet => addConditionalFormatRule(sheet, { range, condition, style }),
      `Added conditional format for ${range}`,
    )
    // Rules render on the Univer canvas by being seeded into its CF engine
    // on unit load — reload so the new rule paints immediately (#270).
    refreshUniverSurface()
  }

  function clearConditionalFormats(): void {
    const range = activeRange()
    if (!range) return
    updateActiveSheetFromCurrent(
      sheet => clearConditionalFormatsInRange(sheet, range),
      `Cleared conditional formats in ${range}`,
    )
    refreshUniverSurface()
  }

  const selectedDisplay =
    univerSelection?.cell === selectedCell
      ? univerSelection.displayValue
      : displayCellValue(activeSheet, selectedCell)
  const selectedFormulaError = isFormulaError(selectedDisplay)

  const importUniverSnapshot = useCallback(
    (snapshot: PureSheetsUniverSnapshot, opts?: { record?: boolean }): void => {
      // `record` (default true) means "this export is the first commit of a new
      // edit — push an undo entry." The surface passes record=false for the
      // follow-up exports of the same edit (formula recalc write-back), so a
      // single action folds into one undo step rather than several (#271).
      const record = opts?.record ?? true
      setDocument(current => {
        const next = documentFromEditorSnapshot(snapshot, current)
        // Univer fires this export for selection, copy, scroll and navigation
        // too — none of which change cell content or formatting. Content and
        // formatting live entirely in workbook.sheets (cursor/selection is
        // tracked separately and never lands here; metadata.updatedAt and the
        // cached engine snapshot are volatile), so compare only the sheets.
        const contentChanged =
          JSON.stringify(next.workbook.sheets) !==
          JSON.stringify(current.workbook.sheets)
        if (!contentChanged) {
          // No content/formatting edit → never an undoable step. Follow the
          // active sheet if it moved, otherwise leave the document untouched so
          // pure selection/copy/navigation don't churn state or the stacks.
          if (next.workbook.activeSheetId === current.workbook.activeSheetId) {
            return current
          }
          return {
            ...current,
            workbook: {
              ...current.workbook,
              activeSheetId: next.workbook.activeSheetId,
            },
          }
        }
        if (record) {
          setUndoStack(stack => [...stack, current].slice(-50))
          setRedoStack([])
        }
        setSaveState('dirty')
        return next
      })
    },
    [],
  )

  const syncUniverSelection = useCallback(
    (selection: UniverSelectionState): void => {
      setUniverSelection(selection)
      if (formulaRangePickRef.current) {
        manualRangeRef.current = false
        setSelectedRange(selection.range)
        setStatusMessage(
          `Using ${selection.range} for ${formulaRangePickRef.current.template}`,
        )
        return
      }
      if (formulaRangeCommitRef.current) {
        manualRangeRef.current = false
        setSelectedCell(formulaRangeCommitRef.current.targetCell)
        setSelectedRange(formulaRangeCommitRef.current.range)
        setEditValue(formulaRangeCommitRef.current.formula)
        setStatusMessage(
          `Inserted ${formulaRangeCommitRef.current.template} formula from ${formulaRangeCommitRef.current.range}`,
        )
        return
      }
      setSelectedCell(selection.cell)
      if (!selection.preserveRange && !manualRangeRef.current) {
        setSelectedRange(selection.range)
      }
      setEditValue(selection.editValue)
    },
    [],
  )

  const resetManualRangeTracking = useCallback((): void => {
    manualRangeRef.current = false
  }, [])

  const setUniverEditorBridge = useCallback(
    (bridge: UniverEditorBridge | null): void => {
      univerEditorRef.current = bridge
    },
    [],
  )

  const normalizedRange = normalizeRangeToken(selectedRange)
  const showCompatibilityGrid = import.meta.env.MODE === 'test'
  const visibleSheets = document.workbook.sheets.filter(
    sheet => !sheet.agentDraft,
  )
  // Data-validation feedback (Phase S3), computed from the LIVE workbook for
  // every sheet: the Univer surface paints these cells and the fallback grid
  // titles them. Flag, never reject — commits still land.
  const invalidFlagsBySheet = useMemo(() => {
    const flagged: Record<string, InvalidCellFlag[]> = {}
    for (const sheet of document.workbook.sheets) {
      const flags = invalidCellFlagsForSheet(sheet)
      if (flags.length) flagged[sheet.id] = flags
    }
    return flagged
  }, [document])
  const activeSheetInvalidCells = useMemo(
    () =>
      new Map(
        (invalidFlagsBySheet[activeSheet.id] ?? []).map(flag => [
          flag.cell,
          flag.message,
        ]),
      ),
    [activeSheet.id, invalidFlagsBySheet],
  )
  // Status-bar summaries (Phase S4): count/sum/average/min/max of the
  // numeric cells in the selected range. Skipped for single cells and for
  // very large selections (evaluating every formula would stall the UI).
  const rangeSummary = useMemo(() => {
    if (!normalizedRange) return null
    const keys = cellsInRange(normalizedRange)
    if (keys.length < 2 || keys.length > 20000) return null
    const numbers: number[] = []
    for (const key of keys) {
      const value = computedCellValue(activeSheet, key)
      if (value.trim() === '') continue
      const numeric = Number(value)
      if (Number.isFinite(numeric)) numbers.push(numeric)
    }
    if (!numbers.length) return null
    const sum = numbers.reduce((total, value) => total + value, 0)
    const compact = (value: number): string =>
      Number.isInteger(value) ? String(value) : String(Number(value.toFixed(4)))
    return `Sum ${compact(sum)} · Avg ${compact(
      sum / numbers.length,
    )} · Min ${compact(Math.min(...numbers))} · Max ${compact(
      Math.max(...numbers),
    )} · Count ${numbers.length}`
  }, [activeSheet, normalizedRange])

  const statusText = `Range ${normalizedRange ?? selectedCell}. ${
    rangeSummary ? `${rangeSummary}. ` : ''
  }${activeSheet.frozenRows ? 'Row frozen. ' : ''}${
    activeSheet.frozenColumns ? 'Column frozen. ' : ''
  }${
    activeSheet.filterRow === 0
      ? `Filter ${columnLabel(activeSheet.filterColumn ?? 0)} on. `
      : ''
  }${statusMessage}`

  // Switcher snapshots: the active sheet's top-left cells as a real grid.
  const loadWorkbookPreview = useCallback(
    async (item: {
      path: string
      kind: 'package' | 'file'
    }): Promise<{ kind: 'html'; html: string; title?: string } | null> => {
      try {
        const loaded = await readWorkbookFromPath(item.path, readTextFile)
        const html = sheetsSnapshotHtml(loaded.document)
        return html
          ? { kind: 'html', html, title: loaded.document.metadata.title }
          : null
      } catch {
        return null
      }
    },
    [],
  )

  return (
    <AppFrame
      headerDocumentName={
        document.metadata.title?.trim() || (lifecycle.doc.path ? lifecycle.doc.path.split(/[\\/]/).pop() : undefined)
      }
      headerActions={
        <DocumentHeaderActions
          lifecycle={lifecycle}
          title={document.metadata.title || 'Untitled'}
          onOpenSwitcher={() => setSwitcherOpen(true)}
        />
      }
    >
      {toast && <Toast role="status">{toast}</Toast>}
      <Shell data-app="sheets">
        <DocumentSwitcher
          appSlug={PURESHEETS_APP_SLUG}
          suffixes={['.sheets', '.sheets.html']}
          variant="modal"
          loadPreview={loadWorkbookPreview}
          open={switcherOpen}
          onClose={() => setSwitcherOpen(false)}
          onOpenDocument={path => {
            setSwitcherOpen(false)
            void openWorkbookPath(path).catch(error => {
              setStatusMessage(
                error instanceof Error ? error.message : String(error),
              )
              setSaveState('error')
            })
          }}
          onCreateNew={() => newWorkbook()}
          newLabel="New workbook"
          title="Open a workbook"
          itemNoun="workbook"
          previewStyle="grid"
        />
        <TopBar>
          <TitleBlock>
            <FileNameInput
              aria-label="Workbook title"
              value={titleDraft ?? document.metadata.title}
              placeholder="Untitled"
              onFocus={() => setTitleDraft(document.metadata.title)}
              onChange={event => setTitleDraft(event.currentTarget.value)}
              onBlur={event => commitWorkbookTitle(event.currentTarget.value)}
              onKeyDown={event => {
                if (event.key === 'Enter') event.currentTarget.blur()
              }}
            />
            <FileMeta>Spreadsheet workbook</FileMeta>
          </TitleBlock>
          <CommandBar
            selectedStyle={selectedStyle}
            canUndo={undoStack.length > 0}
            canRedo={redoStack.length > 0}
            findText={findText}
            replaceText={replaceText}
            onUndo={undo}
            onRedo={redo}
            onCopySelectedRange={copySelectedRange}
            onPasteFromClipboard={() => pasteFromClipboard()}
            onToggleStyle={toggleStyle}
            onApplyCellStyle={applyCellStyle}
            onClearSelectedRange={clearSelectedRange}
            onSortSelectedRange={sortSelectedRange}
            onToggleFilter={() => {
              const next = toggleFilterRow(
                activeSheet,
                selectedPosition?.column ?? 0,
              )
              updateActiveSheet(next, 'Toggled filter row')
              applyFilterToEditor(next)
              // Highlight the column being filtered by selecting it (#274).
              if (next.filterRow === 0 && next.filterColumn !== undefined) {
                const col = columnLabel(next.filterColumn)
                void univerEditorRef.current?.selectRange?.(
                  `${col}1:${col}${next.rowCount}`,
                )
              }
            }}
            onToggleFirstRowFreeze={toggleFirstRowFreeze}
            onToggleFirstColumnFreeze={toggleFirstColumnFreeze}
            firstRowFrozen={(activeSheet.frozenRows ?? 0) > 0}
            firstColumnFrozen={(activeSheet.frozenColumns ?? 0) > 0}
            onCreateChart={createChart}
            onInsertRow={insertSelectedRow}
            onDeleteRow={deleteSelectedRow}
            onInsertColumn={insertSelectedColumn}
            onDeleteColumn={deleteSelectedColumn}
            onResizeColumn={delta =>
              resizeSelectedColumn(selectedColumnWidth() + delta)
            }
            onResizeRow={delta =>
              resizeSelectedRow(selectedRowHeight() + delta)
            }
            onPasteValuesOnly={() => pasteSpecial('values')}
            onPasteFormattingOnly={() => pasteSpecial('formats')}
            onFillDown={() => fillSelectedRange('down')}
            onFillRight={() => fillSelectedRange('right')}
            onAdjustDecimals={adjustDecimals}
            onApplyTableStyle={applyTableStyle}
            onInsertFormulaTemplate={insertFormulaTemplate}
            onFindTextChange={setFindText}
            onReplaceTextChange={setReplaceText}
            onFindNext={findNext}
            onReplaceMatches={replaceMatches}
            matchCase={matchCase}
            onToggleMatchCase={() => setMatchCase(current => !current)}
            onReplaceNext={replaceNext}
            onReplaceAllInSheet={replaceAllMatches}
            onLockSelectedRange={lockSelectedRange}
            onUnlockSelectedRange={unlockSelectedRange}
            onToggleCommentBar={() => setCommentBarOpen(open => !open)}
            commentCount={
              Object.values(activeSheet.comments ?? {}).filter(
                text => text.trim().length > 0,
              ).length
            }
            filterActive={activeSheet.filterRow === 0}
            filterQuery={activeSheet.filterQuery ?? ''}
            filterColumnLabel={columnLabel(activeSheet.filterColumn ?? 0)}
            filterHasHeader={activeSheet.filterHasHeader !== false}
            onToggleFilterHeader={toggleFilterHasHeader}
            onCommitFilterQuery={commitFilterQuery}
            selectedRangeLabel={normalizedRange ?? selectedCell}
            onApplyValidation={applyValidation}
            onClearValidation={clearValidation}
            onApplyConditionalFormat={applyConditionalFormat}
            onClearConditionalFormats={clearConditionalFormats}
            onImportExport={handleImportExport}
            onToggleAgentActivity={() => setAgentActivityOpen(open => !open)}
            agentLogCount={(document.workbook.agentLog ?? []).length}
          />
        </TopBar>

        {importWarnings && importWarnings.length > 0 && (
          <ImportWarningsBanner aria-label="Import warnings">
            <ImportWarningsTitle>Imported with limitations</ImportWarningsTitle>
            <ImportWarningsList>
              {importWarnings.map(warning => (
                <li key={warning}>{warning}</li>
              ))}
            </ImportWarningsList>
            <ImportWarningsDismiss
              type="button"
              onClick={() => setImportWarnings(null)}
            >
              Dismiss
            </ImportWarningsDismiss>
          </ImportWarningsBanner>
        )}

        <SheetSurface>
          <FormulaBar
            selectedCell={selectedCell}
            selectedRange={selectedRange}
            editValue={editValue}
            selectedDisplay={selectedDisplay}
            selectedFormulaError={selectedFormulaError}
            onRangeChange={range =>
              changeSelectedRange(range, { manual: true })
            }
            onRangeInvalid={setStatusMessage}
            onEditValueChange={changeFormulaValue}
            onCommitSelectedCell={commitSelectedCell}
          />

          {agentActivityOpen && (
            <AgentActivityBar
              entries={document.workbook.agentLog ?? []}
              onClose={() => setAgentActivityOpen(false)}
            />
          )}

          {commentBarOpen && (
            <CommentBar
              selectedCell={selectedCell}
              comment={activeSheet.comments?.[selectedCell] ?? ''}
              rangeLabel={
                selectedRange !== `${selectedCell}:${selectedCell}`
                  ? selectedRange
                  : undefined
              }
              comments={activeSheet.comments ?? {}}
              onSaveComment={saveCellComment}
              onSelectCell={selectCell}
              onClose={() => setCommentBarOpen(false)}
            />
          )}

          <MainArea>
            <SheetColumn>
              <GridScroller
                data-testid="univer-spreadsheet-surface"
                onPointerDownCapture={resetManualRangeTracking}
              >
                <UniverSpreadsheetSurface
                  snapshot={univerSnapshot}
                  revision={univerSurfaceRevision}
                  onEditorReady={setUniverEditorBridge}
                  onSelectionChange={syncUniverSelection}
                  onSnapshotChange={importUniverSnapshot}
                  onEditorCopy={captureEditorClipboard}
                  invalidCells={invalidFlagsBySheet}
                />
                {/*
                  The legacy CompatibilityGrid is the editing surface ONLY in
                  vitest (Univer doesn't mount there); in the real app Univer
                  owns the grid. Mounting it hidden still rendered every
                  rowCount×columnCount cell in React on each update — cheap at
                  40×18, but once the grid grew to a full 1000×26 sheet (#272)
                  that hidden tree was ~26k cells and janked scrolling. Only
                  mount it when it's actually the surface.
                */}
                {showCompatibilityGrid && (
                  <CompatibilityGrid
                    sheet={activeSheet}
                    visible={showCompatibilityGrid}
                    gridTemplate={gridTemplate}
                    visibleRows={visibleRows}
                    selectedCell={selectedCell}
                    normalizedRange={normalizedRange}
                    invalidCells={activeSheetInvalidCells}
                    onPasteIntoCell={pasteIntoCell}
                    onPasteFromClipboard={pasteFromClipboard}
                    onClearSelectedRange={clearSelectedRange}
                    onCopySelectedRange={copySelectedRange}
                    onSelectCell={selectCell}
                    onSelectRange={range => changeSelectedRange(range)}
                  />
                )}
              </GridScroller>

              <ChartPanel
                charts={activeSheet.charts ?? []}
                onDeleteChart={deleteChart}
                onUpdateChart={updateChartConfig}
              />
            </SheetColumn>
          </MainArea>

          <SheetTabsBar
            sheets={visibleSheets}
            activeSheetId={activeSheet.id}
            renamingSheetId={renamingSheetId}
            renameDraft={renameDraft}
            statusText={statusText}
            onAddSheet={addSheet}
            onSelectSheet={selectSheet}
            onRenameSheet={renameCurrentSheet}
            onBeginRenameSheet={beginRenameSheet}
            onDuplicateSheet={duplicateCurrentSheet}
            onDeleteSheet={deleteCurrentSheet}
            onRenameDraftChange={setRenameDraft}
            onCommitSheetRename={commitSheetRename}
            onCancelSheetRename={cancelSheetRename}
          />
        </SheetSurface>
      </Shell>
    </AppFrame>
  )
}

function isUniverSnapshot(
  snapshot: unknown,
): snapshot is PureSheetsUniverSnapshot {
  return (
    typeof snapshot === 'object' && snapshot !== null && 'sheets' in snapshot
  )
}

function fileNameFromPath(path: string): string {
  return path.split(/[\\/]/).pop() || path
}

// Brief confirmation toast (copy/cut). Fixed near the bottom-center, above the
// grid, non-interactive, and fades itself out (#277).
const Toast = styled.div`
  position: fixed;
  left: 50%;
  bottom: 56px;
  z-index: 1200;
  transform: translateX(-50%);
  padding: 8px 16px;
  border-radius: 0;
  background: var(--platform-colors-popover);
  color: var(--platform-colors-popover-text);
  font-size: var(--pure-chrome-ui-size);
  font-weight: 600;
  box-shadow: var(--platform-shadow-lg, var(--platform-shadow-md));
  pointer-events: none;
  animation: sheets-toast-in 120ms ease-out;

  @keyframes sheets-toast-in {
    from {
      opacity: 0;
      transform: translate(-50%, 6px);
    }
    to {
      opacity: 1;
      transform: translate(-50%, 0);
    }
  }
`

const Shell = styled.div`
  display: grid;
  grid-template-rows: auto 1fr;
  height: 100%;
  /* One chrome for every app: the --sheets-* names are aliases of the
     platform's --pure-chrome-* tokens so the bars, fields and controls
     around the grid match every other app; the grid itself is Univer's. */
  --sheets-mono: var(--platform-typography-font-family-mono);
  --sheets-canvas: var(--platform-colors-bg);
  --sheets-surface: var(--pure-chrome-surface);
  --sheets-line: var(--pure-chrome-line);
  --sheets-ink: var(--platform-colors-text);
  --sheets-ink-soft: var(--pure-chrome-soft);
  --sheets-ink-faint: var(--pure-chrome-muted);
  /* Office-family identity: Sheets wears the shell's app color in the same
     chrome places its siblings use theirs. Everything accent-colored routes
     through these three tokens. */
  --sheets-accent: var(--pure-chrome-accent);
  --sheets-accent-bg: var(--app-bg, var(--pure-chrome-selection));
  --sheets-accent-text: var(--app-text, var(--pure-chrome-accent));
  background: var(--sheets-canvas);
  color: var(--sheets-ink);
  font-family: var(--platform-typography-font-family);
  font-size: var(--pure-chrome-ui-size);
`

// The title + command bar. It wraps to several rows of 28px toolbar
// controls, so it keeps its own grid rather than the single 36px toolbar;
// colours and hairline are the toolbar's.
const TopBar = styled.header`
  display: grid;
  grid-template-columns: minmax(160px, 1fr) minmax(0, auto) auto;
  gap: var(--pure-chrome-inset);
  align-items: start;
  padding: 6px var(--pure-chrome-inset) 8px;
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-bar);

  @media (max-width: 720px) {
    grid-template-columns: minmax(0, 1fr) auto;
    gap: 8px;
  }
`

const TitleBlock = styled.div`
  min-width: 0;
  padding-top: 2px;
`

const FileNameInput = styled.input`
  width: 100%;
  min-width: 0;
  overflow: hidden;
  margin: 0;
  padding: 0;
  border: 0;
  border-bottom: 1px solid transparent;
  background: transparent;
  color: var(--sheets-ink);
  font: inherit;
  font-size: var(--pure-chrome-ui-size);
  font-weight: 600;
  text-overflow: ellipsis;
  white-space: nowrap;

  &:hover,
  &:focus {
    border-bottom-color: var(--sheets-line);
    outline: 0;
  }
`

const FileMeta = styled(MetaText)`
  display: block;
  overflow: hidden;
  text-overflow: ellipsis;
`

const SheetSurface = styled.div`
  display: flex;
  flex-direction: column;
  min-height: 0;
  overflow: hidden;
`

const ImportWarningsBanner = styled.aside`
  display: grid;
  grid-template-columns: auto minmax(0, 1fr) auto;
  gap: 12px;
  align-items: start;
  padding: 8px 16px;
  border-bottom: 1px solid
    color-mix(
      in srgb,
      var(--platform-colors-warning, var(--platform-colors-semantic-orange))
        34%,
      var(--pure-chrome-line)
    );
  background: color-mix(
    in srgb,
    var(--platform-colors-warning, var(--platform-colors-semantic-orange))
      10%,
    var(--pure-chrome-surface)
  );
  color: var(--platform-colors-semantic-orange-text);
  font-size: var(--pure-chrome-ui-size);
`

const ImportWarningsTitle = styled.strong`
  font-size: var(--pure-chrome-ui-size);
  font-weight: 700;
  white-space: nowrap;
`

const ImportWarningsList = styled.ul`
  margin: 0;
  padding-left: 16px;

  > li {
    margin: 1px 0;
  }
`

const ImportWarningsDismiss = styled.button`
  padding: 4px 10px;
  border: 1px solid var(--platform-colors-semantic-orange-border);
  border-radius: 0;
  background: var(--pure-chrome-surface);
  color: var(--platform-colors-semantic-orange-text);
  font-size: 12px;
  font-weight: 600;
`

const MainArea = styled.div`
  display: flex;
  flex: 1 1 0;
  min-height: 0;
  overflow: hidden;
`

const SheetColumn = styled.div`
  display: flex;
  flex-direction: column;
  flex: 1 1 0;
  min-width: 0;
  min-height: 0;
  overflow: hidden;
`

const GridScroller = styled.div`
  display: grid;
  grid-template-rows: minmax(0, 1fr);
  flex: 1 1 0;
  min-height: 0;
  min-width: 0;
  // Univer renders and scrolls its own grid (its canvas provides both
  // scrollbars). Letting this wrapper ALSO scroll produced a second, nested
  // scrollbar and — with a tall grid — a resize/repaint feedback loop (#272).
  // Clip here and let Univer own scrolling. The legacy CompatibilityGrid
  // (test-only) doesn't rely on real scrolling.
  overflow: hidden;
  background: var(--sheets-canvas);
`
