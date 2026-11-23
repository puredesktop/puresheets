import { linkedTableHtml, tableSnapshotHtml, type LinkedTableReference } from '@purescience/platform-editor/linkedTable'
import { workbookLinkedSnapshot } from '../lib/linkedTables'
import { publishSheetLiveData, sendSheetToApp } from '../lib/liveDataPublisher'
import { WorkspaceThreadStrip } from '@purescience/platform-ui/suite'
import { bridge } from '@purescience/platform-ui/bridge/client'
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { useLiveJson } from '@purescience/platform-ui/editing/live/useLiveJson'
import type React from 'react'
import styled from 'styled-components'
import { createWorkbookTransitions } from '../lib/workbookTransitions'
import { useWorkbookResourceOpen } from '../hooks/useWorkbookResourceOpen'
import { useWorkspaceDocumentBinding } from '../hooks/useWorkspaceDocumentBinding'
import { useSheetsAgentTools } from '../hooks/useSheetsAgentTools'
import { useCompactInRoom } from '../hooks/useCompactInRoom'
import { AppFrame } from '@purescience/platform-bridge/components/AppFrame'
import { useDocumentLifecycle } from '@purescience/platform-ui/bridge/react/useDocumentLifecycle'
import { useDocumentHotkeys } from '@purescience/platform-ui/bridge/react/useDocumentHotkeys'
import {
  DocumentHeaderActions,
  DocumentSwitcher,
} from '@purescience/platform-ui/components/common/documents'
import {
  chooseSaveFilePath,
  fileExists,
  isStandaloneDevMode,
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
import { validateDocument } from '../lib/sheetsDocument'
import { sheetsSnapshotHtml } from '../lib/sheetsSnapshot'
import {
  readWorkbookFromPath,
  validateSafeFilePath,
  workbookFilesForSave,
} from '../lib/workbookFiles'
import {
  createUniverSnapshot,
  createWorkbookUnitId,
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
  clearRangeStyle,
  cloneWithUpdatedSheet,
  columnLabel,
  computedCellValue,
  copyRangeCells,
  createChartFromRange,
  deleteColumn,
  deleteRow,
  fillRange,
  findMatches,
  hasMergeInRange,
  isCellLocked,
  isColumnDeletionBlockedByProtection,
  isRangeLocked,
  isRowDeletionBlockedByProtection,
  lockRange,
  mergeRange,
  pasteCellBlock,
  replaceAllInSheet,
  setCellComment,
  setFilterQuery,
  setFilterHasHeader,
  setFreezePanes,
  sheetHasPositionalMetadata,
  unlockRange,
  unmergeRange,
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
  moveConditionalFormatRule,
  removeConditionalFormatRule,
  removeValidationRule,
  validateValue,
  validationRuleForCell,
  type InvalidCellFlag,
} from '../lib/sheetDataFeatures'
import {
  activeFilterColumnIndices,
  clearAllColumnFilters,
  hideColumnsInSheet,
  hideRowsInSheet,
  moveSheetInWorkbook,
  removeDuplicatesInRange,
  setCellHyperlink,
  setColumnFilter,
  setSheetHidden,
  setSheetTabColor,
  sortRangeMulti,
  splitTextToColumns,
  toggleCheckboxInRange,
  toggleSheetGridlines,
  transposeCopiedBlock,
  trimWhitespaceInRange,
  unhideAllColumnsInSheet,
  unhideAllRowsInSheet,
  type SortKey,
} from '../lib/dataTransforms'
import { recomputeSheetCharts } from '../lib/chartEngine'
import { deleteNamedRange, upsertNamedRange } from '../lib/formulaCatalog'
import { computePivotSheet, inferPivotFields, isPivotStale } from '../lib/pivotEngine'
import { csvFromSheet, isDelimitedTextPath, jsonFromSheet, sheetFromCsv, workbookFromCsv } from '../lib/csv'
import {
  base64ToBytes,
  bytesToBase64,
  exportXlsxWorkbook,
  importXlsxWorkbook,
} from '../lib/xlsxIO'
import type {
  AgentLogEntry,
  CellStyle,
  ColumnFilterRule,
  ConditionalFormatRule,
  NamedRange,
  PivotDefinition,
  PureSheetsDocument,
  ValidationRule,
  WorkbookChart,
  WorkbookSheet,
} from '../types'
import { ToolTextButton } from './controls'
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
import { analysisOptions, WorkbookAnalysisDialog, type AnalysisMode } from './WorkbookAnalysisDialog'
import { ToolbarMenu } from './ToolbarMenu'
import { FormulaBar } from './FormulaBar'
import { SheetTabsBar, type StatPill } from './SheetTabsBar'
import { UniverSpreadsheetSurface } from './UniverSpreadsheetSurface'
import { CommandPalette, type PaletteCommand } from './CommandPalette'
import { InspectorRail, type InspectorTab } from './InspectorRail'
import type { BorderPreset } from './StylePopovers'
import { workbookAgentRefs } from '../agentRefs'
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
  const compact = useCompactInRoom()
  const [document, setDocumentState] = useState<PureSheetsDocument>(() => {
    const existingSnap = initialDocument.engine.snapshot
    const hasId =
      typeof existingSnap === 'object' &&
      existingSnap !== null &&
      typeof (existingSnap as { id?: unknown }).id === 'string' &&
      Boolean((existingSnap as { id: string }).id.trim())
    if (hasId) return initialDocument
    return {
      ...initialDocument,
      engine: {
        ...initialDocument.engine,
        snapshot: { id: createWorkbookUnitId() },
      },
    }
  })
  const documentRef = useRef(document)
  const setDocument = useCallback(
    (
      update:
        | PureSheetsDocument
        | ((current: PureSheetsDocument) => PureSheetsDocument),
    ): void => {
      const next =
        typeof update === 'function' ? update(documentRef.current) : update
      documentRef.current = next
      setDocumentState(next)
    },
    [],
  )
  const [analysis, setAnalysis] = useState<{
    mode: AnalysisMode; source: PureSheetsDocument; sheet: WorkbookSheet; range: string; cell: string; version: number
  } | null>(null)
  const [filePath, setFilePath] = useState('')
  const [selectedCell, setSelectedCell] = useState('A1')
  const [editValue, setEditValue] = useState('')
  const [saveState, setSaveStateState] = useState<SaveState>('saved')
  const saveStateRef = useRef<SaveState>('saved')
  const setSaveState = useCallback((state: SaveState) => {
    saveStateRef.current = state
    setSaveStateState(state)
  }, [])
  const [switching, setSwitching] = useState(false)
  const [statusMessage, setStatusMessage] = useState('Ready')
  // Brief, prominent confirmation (e.g. copy/cut) — easy to miss in the status
  // bar, so surface it as a fading toast too (#277).
  const [toast, setToast] = useState<string | null>(null)
  const toastTimerRef = useRef<number | null>(null)
  useEffect(
    () => () => {
      if (toastTimerRef.current) window.clearTimeout(toastTimerRef.current)
    },
    [],
  )
  function showToast(message: string): void {
    setToast(message)
    if (toastTimerRef.current) window.clearTimeout(toastTimerRef.current)
    toastTimerRef.current = window.setTimeout(() => setToast(null), 1800)
  }
  // The cell the last Find landed on, so a repeat Find advances past it while a
  // fresh Find can still land on the cell you're currently sitting on (#277).
  const lastFindRef = useRef<string | null>(null)
  const [undoStack, setUndoStackState] = useState<PureSheetsDocument[]>([])
  const undoStackRef = useRef<PureSheetsDocument[]>([])
  const [redoStack, setRedoStackState] = useState<PureSheetsDocument[]>([])
  const redoStackRef = useRef<PureSheetsDocument[]>([])
  const setUndoStack = useCallback(
    (
      update:
        | PureSheetsDocument[]
        | ((current: PureSheetsDocument[]) => PureSheetsDocument[]),
    ): void => {
      const next =
        typeof update === 'function' ? update(undoStackRef.current) : update
      undoStackRef.current = next
      setUndoStackState(next)
    },
    [],
  )
  const setRedoStack = useCallback(
    (
      update:
        | PureSheetsDocument[]
        | ((current: PureSheetsDocument[]) => PureSheetsDocument[]),
    ): void => {
      const next =
        typeof update === 'function' ? update(redoStackRef.current) : update
      redoStackRef.current = next
      setRedoStackState(next)
    },
    [],
  )
  const [findOpen, setFindOpen] = useState(false)
  const [findFocusRequest, setFindFocusRequest] = useState(0)
  const [findText, setFindText] = useState('')
  const [replaceText, setReplaceText] = useState('')
  const [matchCase, setMatchCase] = useState(false)
  const [sortHasHeader, setSortHasHeader] = useState(false)
  const skipTitleBlurCommitRef = useRef(false)
  const [commentBarOpen, setCommentBarOpen] = useState(false)
  const [importWarnings, setImportWarnings] = useState<string[] | null>(null)
  const [warningsTitle, setWarningsTitle] = useState('Imported with limitations')
  const importWarningsRef = useRef<string[]>([])
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null)
  const [renameDraft, setRenameDraft] = useState('')
  const [agentActivityOpen, setAgentActivityOpen] = useState(false)
  const [paletteOpen, setPaletteOpen] = useState(false)
  const [inspectorOpen, setInspectorOpen] = useState(false)
  const [inspectorTab, setInspectorTab] = useState<InspectorTab>('rules')
  const [formatPainterStyle, setFormatPainterStyle] = useState<CellStyle | null>(null)
  const [showFormulas, setShowFormulas] = useState(false)
  const [zoom, setZoom] = useState(1)
  const [univerSurfaceRevision, setUniverSurfaceRevision] = useState(0)
  const [selectedRange, setSelectedRange] = useState('A1:A1')
  const [univerSelection, setUniverSelection] =
    useState<UniverSelectionState | null>(null)
  useSheetsAgentTools(
    !isStandaloneDevMode(),
    {
      document,
      setDocument: updater => {
        if (workbookVersion !== transitions.version)
          throw Error(
            'The workbook changed while the tool was running. Read its current context again.',
          )
        if (commitDocument(updater)) {
          refreshUniverSurface()
        }
      },
      filePath,
      currentFilePath: () => boundPathRef.current ?? '',
      selectedRange,
      autoFitRows: async (sheetId, rows) =>
        (await univerEditorRef.current?.autoFitRows(sheetId, rows)) ?? false,
      openWorkbook: async path => {
        const safePath = validateSafeFilePath(path)
        if (!await openWorkbookPath(safePath)) throw Error('A newer workbook request superseded this open. Try again.')
        return { title: documentRef.current.metadata.title, sourcePath: safePath, filePath: boundPathRef.current ?? null, warnings: importWarningsRef.current }
      },
      saveAsXlsx: async (path, overwrite) => {
        const safePath = validateSafeFilePath(path, { requireAbsolute: true })
        if (!/\.xlsx$/i.test(safePath)) throw Error('The Excel path must end in .xlsx.')
        if (!overwrite && (await fileExists(safePath)))
          throw Error(`${fileNameFromPath(safePath)} already exists. Choose another path, or pass overwrite: true if the user asked to replace it.`)
        return exportXlsxToPath(safePath)
      },
      createWorkbook: async title => {
        const fresh = createBlankWorkbook(title)
        if (!(await newWorkbook(fresh)))
          throw Error('A newer workbook request superseded creation.')
        return { title: fresh.metadata.title, filePath: null }
      },
      completeWorkbook: async () => {
        if (transitions.busy)
          throw Error('Wait for the workbook to finish opening before saving.')
        await captureEditorWorkbook()
        const snapshot = documentRef.current
        const originalPath = boundPathRef.current
        const completionVersion = transitions.version
        let savedPath: string | null = null
        return createWorkbookReceipt({
          save: async () => {
            savedPath = await persistWorkbook(snapshot)
            const payload = savedWorkbookPayloadRef.current
            if (!payload)
              throw new Error(
                'Workbook save did not return its written content.',
              )
            return {
              path: savedPath,
              contentPath: payload.name
                ? `${savedPath}/${payload.name}`
                : savedPath,
              content: payload.content,
            }
          },
          read: readTextFile,
          unchanged: () =>
            documentRef.current === snapshot &&
            transitions.version === completionVersion &&
            (boundPathRef.current === originalPath ||
              boundPathRef.current === savedPath),
        })
      },
    },
    () => !transitions.busy,
  )
  const formulaDraftRef = useRef<{sheetId: string; cell: string; value: string} | null>(null)
  const cellCommitRef = useRef<Promise<void>>(Promise.resolve())
  const manualRangeRef = useRef(false)
  const univerEditorRef = useRef<UniverEditorBridge | null>(null)
  // In-app clipboard: the copied cell block (formulas + styles) plus the TSV
  // that was written to the system clipboard at copy time. When a paste's
  // clipboard text still matches, the block is pasted instead of the lossy
  // TSV, and it also backs paste-values / paste-formatting.
  const copiedBlockRef = useRef<CopiedCellBlock | null>(null)
  const copyRequestRef = useRef(0)
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
  // New workbooks are `.sheets` package folders; legacy flat `.sheets` /
  // `.sheets.html` files opened from disk stay single files. The package
  // suffix collides with the legacy flat suffix, so package-ness can't be
  // read from the path — it's tracked explicitly (set at bind/open time).
  const boundIsPackageRef = useRef(true)
  const [switcherOpen, setSwitcherOpen] = useState(false)
  const [titleDraft, setTitleDraft] = useState<string | null>(null)
  const savedWorkbookPayloadRef = useRef<{
    name: string | null
    content: string
  } | null>(null)
  const savingDocumentRef = useRef<PureSheetsDocument | null>(null)
  const [transitions] = useState(() =>
    createWorkbookTransitions(flushBeforeReplace, setSwitching),
  )
  const workbookVersion = transitions.version
  useEffect(() => () => transitions.dispose(), [transitions])

  const lifecycle = useDocumentLifecycle({
    appSlug: PURESHEETS_APP_SLUG,
    suffix: '.sheets',
    kind: 'package',
    suggestedTitle: () => documentRef.current.metadata.title,
    onSaved: files => {
      if (documentRef.current === savingDocumentRef.current)
        setSaveState('saved')
      const file =
        files.find(entry => entry.name === 'workbook.json') ??
        files.find(entry => entry.name === null)
      savedWorkbookPayloadRef.current =
        file && typeof file.content === 'string'
          ? { name: file.name ?? null, content: file.content }
          : null
    },
    serialize: () => {
      savingDocumentRef.current = boundDocRef.current
      return workbookFilesForSave(boundDocRef.current, {
        boundPath: boundPathRef.current,
        isPackage: boundIsPackageRef.current,
      })
    },
  })
  const lifecycleRef = useRef(lifecycle)
  lifecycleRef.current = lifecycle

  // Lifecycle-owned path changes (lazy draft, promote, rename) → filePath.
  // These paths are always `.sheets` packages (drafts and their promotions).
  useEffect(() => {
    const path = lifecycle.doc.path
    if (!path || boundPathRef.current === path) return
    boundIsPackageRef.current = true
    boundPathRef.current = path
    setFilePath(path)
  }, [lifecycle.doc.path])

  // The workbook on disk is this tab's document, as far as the desktop knows: bound once it has a path (an opened
  // file, or the draft a first edit creates), so Share on the header offers it, the desk brings this tab forward
  // instead of opening it twice, and a restart reopens it.
  const workspaceBinding = useWorkspaceDocumentBinding(filePath.trim() || null)
  useEffect(() => {
    if (workspaceBinding.error) setStatusMessage(workspaceBinding.error)
  }, [workspaceBinding.error])

  // A workbook the organisation shares, edited live (useLiveJson): everyone's cells merged cell by cell; the live
  // copy writes the workbook, so this tab stops saving it itself while live.
  const livePath = filePath && boundIsPackageRef.current && boundPathRef.current === filePath ? `${filePath}/workbook.json` : null
  const lastRemoteDocRef = useRef<PureSheetsDocument | null>(null)
  const live = useLiveJson(livePath, {
    loadSaved: async path => JSON.parse(await readTextFile(path)) as Record<string, unknown>,
    onRemote: next => {
      // Someone else's change: validate untrusted remote JSON before applying
      // so malformed payloads never corrupt local state, and record the
      // reference so the dirty-save effect never echoes it back.
      try {
        const doc = validateDocument(next)
        lastRemoteDocRef.current = doc
        setDocument(doc)
        boundDocRef.current = doc
        setSaveState('saved')
        refreshUniverSurface()
      } catch (error) {
        setStatusMessage(
          `Ignored invalid remote update: ${
            error instanceof Error ? error.message : String(error)
          }`,
        )
      }
    },
    serialize: () => ({
      resourcePath: boundPathRef.current ?? filePath,
      files: workbookFilesForSave(boundDocRef.current, { boundPath: boundPathRef.current, isPackage: boundIsPackageRef.current }).map(file => ({ name: file.name ?? 'workbook.json', content: String(file.content) })),
    }),
    // Let go by the live copy: the workbook on screen is saved as it always was, so a cell typed just before is kept.
    onEnded: () => lifecycleRef.current.markDirty(),
  })
  const liveRef = useRef(live)
  liveRef.current = live

  // Every user edit marks the bound workbook dirty — the first one lazily
  // creates the durable draft in PureDrafts.
  useEffect(() => {
    if (saveState !== 'dirty') return
    if ((filePath.trim() || null) !== boundPathRef.current) return
    boundDocRef.current = document
    // Live: the change goes to everyone; the live copy writes the workbook when it is time.
    if (liveRef.current.status === 'on') {
      if (document !== lastRemoteDocRef.current) {
        liveRef.current.write(
          JSON.parse(JSON.stringify(document)) as Record<string, unknown>,
        )
      }
      setSaveState('saved')
      return
    }
    lifecycleRef.current.markDirty()
  }, [document, filePath, saveState, setSaveState])

  // Never let a previous savedAt turn a failed write into a success indication.
  useEffect(() => {
    if (lifecycle.doc.error) {
      setSaveState('error')
      setStatusMessage(lifecycle.doc.error)
    }
  }, [lifecycle.doc.error, setSaveState])
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
    () => visibleRowsForSheet(activeSheet, document.workbook),
    [activeSheet, document.workbook],
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
    async (resourcePath: string): Promise<boolean> => {
      const safePath = validateSafeFilePath(resourcePath)
      if (/\.xlsx$/i.test(safePath)) return importXlsxPath(safePath)
      if (isDelimitedTextPath(safePath)) return importCsvPath(safePath)
      setStatusMessage('Opening workbook…')
      const opened = await transitions.run(
        () => readWorkbookFromPath(safePath, readTextFile),
        loaded => {
          importWarningsRef.current = []
          setImportWarnings(null)
          // The old source and file format have been saved before either changes.
          lifecycleRef.current.adopt(loaded.path, {
            title: loaded.document.metadata.title,
          })
          boundPathRef.current = loaded.path
          boundDocRef.current = loaded.document
          boundIsPackageRef.current = loaded.isPackage
          savedWorkbookPayloadRef.current = null
          setDocument(loaded.document)
          setFilePath(loaded.path)
          resetWorkbookView()
          setSaveState(loaded.recoveredFromBackup ? 'dirty' : 'saved')
          setStatusMessage(
            loaded.recoveredFromBackup
              ? `The file was unreadable; recovered the last good save from ${fileNameFromPath(
                  loaded.path,
                )}.bak`
              : `Opened ${fileNameFromPath(loaded.path)}`,
          )
          const version = transitions.version
          void (async () =>
            updatePureSheetsSettings({ lastFilePath: loaded.path }))().catch(
            error => {
              if (transitions.version === version)
                showToast(
                  `Workbook opened; could not remember its location: ${String(
                    error,
                  )}`,
                )
            },
          )
        },
      )
      return opened
    },
    [transitions, setDocument, setSaveState],
  )

  const resourceOpen = useWorkbookResourceOpen(
    true,
    resource?.path?.trim(),
    async path => {
      try {
        return await openWorkbookPath(path)
      } catch (error) {
        reportFailure(error)
        return false
      }
    },
    onResourceHandled ?? (() => {}),
  )

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
  ): boolean {
    if (transitions.busy)
      throw Error('Wait for the workbook to finish opening before editing.')
    const current = documentRef.current
    const next = updater(current)
    if (next === current) return false
    setUndoStack(stack => [...stack, current].slice(-50))
    setRedoStack([])
    setDocument(next)
    boundDocRef.current = next
    setSaveState('dirty')
    if (message) setStatusMessage(message)
    return true
  }

  function refreshUniverSurface(): void {
    setUniverSurfaceRevision(revision => revision + 1)
  }

  function updateActiveSheet(nextSheet: WorkbookSheet, message?: string): void {
    if (transitions.busy || workbookVersion !== transitions.version) return
    commitDocument(
      current =>
        cloneWithUpdatedSheet(
          current,
          recomputeSheetCharts(nextSheet, current.workbook),
        ),
      message,
    )
  }

  function updateActiveSheetFromCurrent(
    updater: (sheet: WorkbookSheet) => WorkbookSheet,
    message?: string,
  ): void {
    if (transitions.busy || workbookVersion !== transitions.version) return
    commitDocument(current => {
      const currentSheet =
        current.workbook.sheets.find(
          sheet => sheet.id === current.workbook.activeSheetId,
        ) ?? current.workbook.sheets[0]
      const updated = updater(currentSheet)
      return cloneWithUpdatedSheet(
        current,
        recomputeSheetCharts(updated, current.workbook),
      )
    }, message)
  }

  function commitCellValue(cell: string, value: string): void {
    if (transitions.busy || workbookVersion !== transitions.version) return
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
      cellCommitRef.current = univerEditor
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
    formulaDraftRef.current = null
    commitCellValue(selectedCell, value)
  }

  function formulaReferenceRange(
    range: string,
    cell = selectedCell,
    exactRange = false,
  ): string {
    const selectedPosition = parseCellKey(cell)
    const normalizedRange = normalizeRangeToken(range) ?? range
    const referencedCells = cellsInRange(normalizedRange, 2)
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
    if (template === 'XLOOKUP')
      return `=XLOOKUP(${cell}, ${referenceRange}, ${referenceRange}, "")`
    if (template === 'COUNTIF') return `=COUNTIF(${referenceRange}, ">0")`
    if (template === 'SUMIF') return `=SUMIF(${referenceRange}, ">0")`
    if (template === 'SUMIFS')
      return `=SUMIFS(${referenceRange}, ${referenceRange}, ">0")`
    if (template === 'AVERAGEIF') return `=AVERAGEIF(${referenceRange}, ">0")`
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
    const referencedCells = cellsInRange(range, 2)
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
    formulaDraftRef.current = {sheetId: activeSheet.id, cell: selectedCell, value}
    formulaRangePickRef.current = null
    if (formulaRangePickTimeoutRef.current) {
      window.clearTimeout(formulaRangePickTimeoutRef.current)
      formulaRangePickTimeoutRef.current = undefined
    }
    setEditValue(value)
  }

  function cancelFormulaEdit(): void {
    formulaDraftRef.current = null
    formulaRangePickRef.current = null
    if (formulaRangePickTimeoutRef.current) {
      window.clearTimeout(formulaRangePickTimeoutRef.current)
      formulaRangePickTimeoutRef.current = undefined
    }
    setEditValue(activeSheet.cells[selectedCell]?.value ?? '')
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
    const stack = undoStackRef.current
    const previous = stack.at(-1)
    if (!previous) return
    const current = documentRef.current
    setUndoStack(stack.slice(0, -1))
    setRedoStack([current, ...redoStackRef.current].slice(0, 50))
    setDocument(previous)
    boundDocRef.current = previous
    // Rebuild the Univer surface from the restored document — without this
    // the grid keeps showing the un-undone state and the next grid command
    // re-exports it, silently re-committing the undone edit.
    refreshUniverSurface()
    setSaveState('dirty')
    setStatusMessage('Undid last edit')
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
    const stack = redoStackRef.current
    const next = stack[0]
    if (!next) return
    const current = documentRef.current
    setRedoStack(stack.slice(1))
    setUndoStack([...undoStackRef.current, current].slice(-50))
    setDocument(next)
    boundDocRef.current = next
    refreshUniverSurface()
    setSaveState('dirty')
    setStatusMessage('Redid edit')
  }

  function reportFailure(error: unknown): void {
    const message = error instanceof Error ? error.message : String(error)
    setStatusMessage(message)
    showToast(message)
  }

  /** Commit the actual editor, including a cell that has not received Enter yet. */
  async function captureEditorWorkbook(): Promise<void> {
    const version = transitions.version
    await cellCommitRef.current
    if (version !== transitions.version)
      throw Error('The workbook changed while committing its cell.')
    const editor = univerEditorRef.current
    if (editor?.isCellEditing() && !(await editor.commitEditing()))
      throw Error(
        'The current cell could not be committed. Finish editing it before saving.',
      )
    const draft = formulaDraftRef.current
    if (draft) {
      const current = documentRef.current
      const sheet = current.workbook.sheets.find(
        sheet => sheet.id === draft.sheetId,
      )
      if (!sheet || current.workbook.activeSheetId !== draft.sheetId)
        throw Error(
          'The sheet changed while editing the formula. Return to its cell before saving.',
        )
      if (isCellLocked(sheet, draft.cell))
        throw Error(
          `${draft.cell} is protected — unlock it before saving this edit.`,
        )
      if (editor) {
        if (!(await editor.setCellValue(draft.cell, draft.value)))
          throw Error(
            'The formula bar edit could not be committed. Try saving again.',
          )
      } else {
        setDocument(
          cloneWithUpdatedSheet(
            current,
            upsertCell(sheet, draft.cell, draft.value),
          ),
        )
        setSaveState('dirty')
      }
      if (formulaDraftRef.current === draft) formulaDraftRef.current = null
    }
    if (!editor) return
    const snapshot = await editor.flushSnapshot()
    if (version !== transitions.version || editor !== univerEditorRef.current)
      throw Error('The workbook changed while its editor was being saved.')
    if (!snapshot)
      throw Error(
        'The editor could not provide the current workbook. Try saving again.',
      )
    importUniverSnapshot(snapshot)
  }

  async function flushBeforeReplace(): Promise<void> {
    await captureEditorWorkbook()
    boundDocRef.current = documentRef.current
    // A clean restored file needs no forced write just to open another workbook.
    if (saveStateRef.current !== 'saved' || lifecycleRef.current.doc.saving)
      await persistWorkbook(documentRef.current)
  }

  async function persistWorkbook(
    snapshot: PureSheetsDocument,
  ): Promise<string> {
    const version = transitions.version
    const lifecycle = lifecycleRef.current
    boundDocRef.current = snapshot
    if (!boundPathRef.current) boundIsPackageRef.current = true
    const path = await lifecycle.ensureDraft()
    if (!path)
      throw Error(lifecycle.doc.error ?? 'Could not create a workbook draft.')
    if (documentRef.current !== snapshot || transitions.version !== version)
      throw Error(
        'The workbook changed while creating its draft. Save it again.',
      )
    boundPathRef.current = path
    setFilePath(path)
    await lifecycle.flush({ throwOnError: true })
    if (documentRef.current !== snapshot || transitions.version !== version)
      throw Error('New changes arrived while saving. Save the workbook again.')
    return path
  }

  async function save(): Promise<void> {
    if (transitions.busy) return
    setSaveState('saving')
    try {
      await captureEditorWorkbook()
      const snapshot = documentRef.current
      const path = await persistWorkbook(snapshot)
      setSaveState('saved')
      setStatusMessage(
        lifecycleRef.current.doc.status === 'draft'
          ? 'Draft saved — name it to file it.'
          : 'All changes saved.',
      )
      recordLedgerOperation({
        lane: 'user',
        kind: 'document.save',
        summary: `Saved workbook "${snapshot.metadata.title || 'Untitled'}"`,
        refs: { path },
      })
    } catch (error) {
      setSaveState('error')
      reportFailure(error)
    }
  }

  function resetWorkbookView(): void {
    setSwitcherOpen(false)
    formulaDraftRef.current = null
    cellCommitRef.current = Promise.resolve()
    refreshUniverSurface()
    setUndoStack([])
    setRedoStack([])
    setSelectedCell('A1')
    manualRangeRef.current = false
    setSelectedRange('A1:A1')
    setTitleDraft(null)
    setUniverSelection(null)
    setImportWarnings(null)
    copiedBlockRef.current = null
    copiedTextRef.current = null
    formulaRangePickRef.current = null
    formulaRangeCommitRef.current = null
    if (formulaRangePickTimeoutRef.current)
      window.clearTimeout(formulaRangePickTimeoutRef.current)
    if (formulaRangeCommitClearRef.current)
      window.clearTimeout(formulaRangeCommitClearRef.current)
  }

  function installNewWorkbook(next: PureSheetsDocument): void {
    importWarningsRef.current = []
    setImportWarnings(null)
    lifecycleRef.current.reset()
    boundPathRef.current = null
    boundIsPackageRef.current = true
    boundDocRef.current = next
    savedWorkbookPayloadRef.current = null
    setDocument(next)
    setFilePath('')
    resetWorkbookView()
    setSaveState('saved')
    setStatusMessage('New workbook')
  }

  async function newWorkbook(fresh = createBlankWorkbook()): Promise<boolean> {
    return transitions.run(async () => fresh, installNewWorkbook)
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
    onNew: () => void newWorkbook().catch(reportFailure),
    // The document list is hidden beside the Executive room; Expand brings it back.
    onOpen: () => {
      if (!compact) setSwitcherOpen(true)
    },
  })

  // Single stable ref-backed shortcut handler (`[]` deps):
  // ⌘/Ctrl+K (Palette), ⌘/Ctrl+F/H (Find/Replace), Ctrl+` (Show Formulas),
  // ⌘/Ctrl+Z (Undo), ⌘/Ctrl+Shift+Z / Ctrl+Y (Redo), ⌘/Ctrl+Shift+V (Paste Values),
  // ⌘/Ctrl+Shift+L (Toggle Filter), ⌘/Ctrl+B/I/U (Bold/Italic/Underline), ⌘/Ctrl+D/R (Fill).
  const shortcutHandlersRef = useRef({
    switcherOpen: false,
    togglePalette: () => {},
    openFind: () => {},
    toggleShowFormulas: () => {},
    undo: () => {},
    redo: () => {},
    pasteValues: () => {},
    toggleFilter: () => {},
    toggleStyle: (_style: 'bold' | 'italic' | 'underline') => {},
    fillSelectedRange: (_dir: 'down' | 'right') => {},
  })
  shortcutHandlersRef.current = {
    switcherOpen,
    togglePalette: () => setPaletteOpen(open => !open),
    openFind: () => {
      setFindOpen(true)
      setFindFocusRequest(count => count + 1)
    },
    toggleShowFormulas: () => setShowFormulas(prev => !prev),
    undo,
    redo,
    pasteValues: () => pasteSpecial('values'),
    toggleFilter: () => {
      const next = toggleFilterRow(activeSheet, selectedPosition?.column ?? 0)
      updateActiveSheet(next, 'Toggled filter row')
      applyFilterToEditor(next)
    },
    toggleStyle,
    fillSelectedRange,
  }
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.defaultPrevented) return
      const handlers = shortcutHandlersRef.current
      if (handlers.switcherOpen) return
      const target = event.target as HTMLElement | null
      if (typeof target?.closest === 'function' && target.closest('[role="dialog"]')) return
      const key = event.key.toLowerCase()
      const mod = event.metaKey || event.ctrlKey
      if (event.ctrlKey && !event.metaKey && !event.altKey && key === '`') {
        event.preventDefault()
        handlers.toggleShowFormulas()
        return
      }
      if (!mod || event.altKey) return
      const inTextTarget = Boolean(
        target &&
          (target.tagName === 'INPUT' ||
            target.tagName === 'TEXTAREA' ||
            target.isContentEditable),
      )
      if (!event.shiftKey && key === 'k') {
        event.preventDefault()
        handlers.togglePalette()
        return
      }
      if (!event.shiftKey && key === 'f') {
        event.preventDefault()
        handlers.openFind()
        return
      }
      if (inTextTarget || univerEditorRef.current?.isCellEditing()) return
      if (!event.shiftKey && key === 'h') {
        event.preventDefault()
        handlers.openFind()
        return
      }
      if (key === 'z') {
        event.preventDefault()
        if (event.shiftKey) handlers.redo()
        else handlers.undo()
        return
      }
      if (!event.shiftKey && key === 'y') {
        event.preventDefault()
        handlers.redo()
        return
      }
      if (event.shiftKey && key === 'v') {
        event.preventDefault()
        handlers.pasteValues()
        return
      }
      if (event.shiftKey && key === 'l') {
        event.preventDefault()
        handlers.toggleFilter()
        return
      }
      if (event.shiftKey) return
      if (key === 'b') {
        event.preventDefault()
        handlers.toggleStyle('bold')
      } else if (key === 'i') {
        event.preventDefault()
        handlers.toggleStyle('italic')
      } else if (key === 'u') {
        event.preventDefault()
        handlers.toggleStyle('underline')
      } else if (key === 'd') {
        event.preventDefault()
        handlers.fillSelectedRange('down')
      } else if (key === 'r') {
        event.preventDefault()
        handlers.fillSelectedRange('right')
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])

  function toggleStyle(
    style: 'bold' | 'italic' | 'underline' | 'strikethrough',
  ): void {
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

  function clearSelectedFormatting(): void {
    const range = activeRange()
    if (!range) return
    updateActiveSheetFromCurrent(
      sheet => clearRangeStyle(sheet, range),
      `Cleared formatting in ${range}`,
    )
    refreshUniverSurface()
  }

  function toggleMergeSelectedRange(): void {
    const range = activeRange()
    if (!range) return
    if (isRangeLocked(activeSheet, range)) {
      setStatusMessage(`${range} contains protected cells — merge refused`)
      return
    }
    const merged = hasMergeInRange(activeSheet, range)
    const editor = univerEditorRef.current
    if (merged) {
      if (editor?.unmergeRange) {
        void editor.unmergeRange(range).then(ok => {
          if (!ok) {
            updateActiveSheetFromCurrent(sheet => unmergeRange(sheet, range), `Unmerged ${range}`)
            refreshUniverSurface()
          }
        })
      } else {
        updateActiveSheetFromCurrent(sheet => unmergeRange(sheet, range), `Unmerged ${range}`)
        refreshUniverSurface()
      }
      setStatusMessage(`Unmerged ${range}`)
      return
    }
    const [startKey, endKey] = range.split(':')
    if (startKey === endKey) {
      setStatusMessage('Select multiple cells to merge')
      return
    }
    if (editor?.mergeRange) {
      void editor.mergeRange(range).then(ok => {
        if (!ok) {
          updateActiveSheetFromCurrent(sheet => mergeRange(sheet, range), `Merged ${range}`)
          refreshUniverSurface()
        }
      })
    } else {
      updateActiveSheetFromCurrent(sheet => mergeRange(sheet, range), `Merged ${range}`)
      refreshUniverSurface()
    }
    setStatusMessage(`Merged ${range}`)
  }

  function addSheet(): void {
    const name = uniqueSheetName(`Sheet ${documentRef.current.workbook.sheets.length + 1}`)
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
    const currentSheets = documentRef.current.workbook.sheets
    const targetSheet = currentSheets.find(s => s.id === sheetId)
    if (!targetSheet || targetSheet.name === nextName) return
    if (
      currentSheets.some(
        sheet =>
          sheet.id !== sheetId &&
          sheet.name.toLowerCase() === nextName.toLowerCase(),
      )
    ) {
      setStatusMessage(`A sheet named "${nextName}" already exists`)
      return
    }
    const message = 'Renamed sheet'
    const record = () =>
      recordLedgerOperation({
        lane: 'user',
        kind: 'sheets.sheet.rename',
        summary: `Renamed a sheet to "${nextName}"`,
      })
    const fallback = () => {
      if (
        commitDocument(
          current => renameWorkbookSheet(current, sheetId, nextName),
          message,
        )
      ) {
        refreshUniverSurface()
      }
    }
    const hasCrossSheetFormulas = currentSheets.some(sheet =>
      Object.values(sheet.cells).some(
        cell => cell.kind === 'formula' && cell.value.includes('!'),
      ),
    )
    const editor = univerEditorRef.current
    if (!editor || hasCrossSheetFormulas) {
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
    const nonDraftSheets = documentRef.current.workbook.sheets.filter(
      sheet => !sheet.agentDraft,
    )
    if (nonDraftSheets.length <= 1) return
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
    const fallback = () => {
      updateActiveSheetFromCurrent(
        sheet => pasteTabularData(sheet, cell, text),
        message,
      )
      refreshUniverSurface()
    }
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
    if (cut && isRangeLocked(activeSheet, range)) {
      setStatusMessage(`${range} contains protected cells — cut refused`)
      return
    }
    let text = rangeToDelimitedText(activeSheet, range, document.workbook)
    // Capture the rich block BEFORE a cut clears the source cells.
    let block = copyRangeCells(activeSheet, range, document.workbook)
    const copyRequest = ++copyRequestRef.current
    const version = transitions.version
    try {
      let html: string | undefined
      let live = false
      if (!cut) {
        // Publish a durable source before putting a link on the clipboard.
        // Skip redundant disk writes when already bound & cleanly saved.
        try {
          await captureEditorWorkbook()
          if (copyRequest !== copyRequestRef.current || version !== transitions.version) return
          const snapshot = documentRef.current
          const sourceSheet = snapshot.workbook.sheets.find(sheet => sheet.id === activeSheet.id)
          if (!sourceSheet) return
          text = rangeToDelimitedText(sourceSheet, range, snapshot.workbook)
          block = copyRangeCells(sourceSheet, range, snapshot.workbook)
          const path =
            boundPathRef.current &&
            saveStateRef.current === 'saved' &&
            !lifecycleRef.current.doc.saving
              ? boundPathRef.current
              : await persistWorkbook(snapshot)
          if (copyRequest !== copyRequestRef.current || version !== transitions.version) return
          const reference: LinkedTableReference = { version: 1, path,
            file: boundIsPackageRef.current ? 'workbook.json' : null, tableId: activeSheet.id, range }
          html = linkedTableHtml(reference, workbookLinkedSnapshot(snapshot, reference))
          live = true
        } catch {
          if (copyRequest !== copyRequestRef.current || version !== transitions.version) return
          // A missing desktop bridge/save permission must not break ordinary copy.
          try {
            html = tableSnapshotHtml(workbookLinkedSnapshot(documentRef.current, {
              version: 1, path: '/unsaved.sheets', file: null, tableId: activeSheet.id, range,
            }))
          } catch { /* Oversized ranges keep their ordinary TSV copy. */ }
        }
      }
      if (copyRequest !== copyRequestRef.current || version !== transitions.version) return
      if (html) await writeClipboard(text, html)
      else await writeClipboard(text)
      if (copyRequest !== copyRequestRef.current || version !== transitions.version) return
      copiedBlockRef.current = block
      copiedTextRef.current = text
      const message = `${cut ? 'Cut' : live ? 'Copied live range' : 'Copied static range (source could not be linked)'} ${range}`
      setStatusMessage(message)
      showToast(message)
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
  function captureEditorClipboard(cut = false): void {
    const range = normalizeRangeToken(selectedRange)
    if (!range) return
    copiedBlockRef.current = copyRangeCells(activeSheet, range, document.workbook)
    copiedTextRef.current = rangeToDelimitedText(activeSheet, range, document.workbook)
    if (!cut) void copySelectedRange()
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
    if (isRowDeletionBlockedByProtection(activeSheet, row)) {
      setStatusMessage(
        `Row ${row + 1} contains protected cells — delete refused`,
      )
      return
    }
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
    if (isColumnDeletionBlockedByProtection(activeSheet, column)) {
      setStatusMessage(
        `Column ${columnLabel(column)} contains protected cells — delete refused`,
      )
      return
    }
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

  function autoFitSelectedRows(): void {
    const range = activeRange()
    if (!range) return
    const [startKey, endKey] = range.split(':')
    const start = parseCellKey(startKey ?? '')
    const end = parseCellKey(endKey ?? startKey ?? '')
    if (!start || !end) return
    if (activeSheet.protection?.lockedRanges?.length) {
      setStatusMessage('Auto-fit row height is refused on protected sheets')
      return
    }
    const rowSpan = end.row - start.row + 1
    if (rowSpan > 500) {
      setStatusMessage('Select at most 500 rows to auto-fit at once')
      return
    }
    const rows: number[] = []
    for (let r = start.row; r <= end.row; r += 1) rows.push(r)
    const fallback = (): void => {
      updateActiveSheetFromCurrent(sheet => {
        const maxLinesByRow = new Map<number, number>()
        for (const k of Object.keys(sheet.cells)) {
          const pos = parseCellKey(k)
          if (!pos || pos.row < start.row || pos.row > end.row) continue
          const text = displayCellValue(sheet, k, documentRef.current.workbook)
          const colWidth = sheet.columnWidths?.[String(pos.column)] ?? 112
          const charsPerLine = Math.max(8, Math.floor((colWidth - 16) / 7))
          const lines = text
            .split('\n')
            .reduce(
              (sum, part) =>
                sum + Math.max(1, Math.ceil(part.length / charsPerLine)),
              0,
            )
          const prev = maxLinesByRow.get(pos.row) ?? 1
          if (lines > prev) maxLinesByRow.set(pos.row, lines)
        }
        const nextRowHeights: Record<string, number> = { ...sheet.rowHeights }
        for (const r of rows) {
          const lines = maxLinesByRow.get(r) ?? 1
          nextRowHeights[String(r)] = Math.min(160, Math.max(24, Math.max(28, lines * 20 + 10)))
        }
        return { ...sheet, rowHeights: nextRowHeights }
      }, `Auto-fitted row height for ${range}`)
      refreshUniverSurface()
    }
    const editor = univerEditorRef.current
    if (!editor) {
      fallback()
      return
    }
    void editor
      .autoFitRows(activeSheet.id, rows)
      .then(applied => {
        if (applied) setStatusMessage(`Auto-fitted row height for ${range}`)
        else fallback()
      })
      .catch(fallback)
  }

  function autoFitSelectedColumns(): void {
    const range = activeRange()
    if (!range) return
    const [startKey, endKey] = range.split(':')
    const start = parseCellKey(startKey ?? '')
    const end = parseCellKey(endKey ?? startKey ?? '')
    if (!start || !end) return
    if (activeSheet.protection?.lockedRanges?.length) {
      setStatusMessage('Auto-fit column width is refused on protected sheets')
      return
    }
    const colSpan = end.column - start.column + 1
    if (colSpan > 500) {
      setStatusMessage('Select at most 500 columns to auto-fit at once')
      return
    }
    const columns: number[] = []
    for (let c = start.column; c <= end.column; c += 1) columns.push(c)
    const fallback = (): void => {
      updateActiveSheetFromCurrent(sheet => {
        const maxCharsByCol = new Map<number, number>()
        for (const key of Object.keys(sheet.cells)) {
          const pos = parseCellKey(key)
          if (!pos || pos.column < start.column || pos.column > end.column) continue
          const text = displayCellValue(
            sheet,
            key,
            documentRef.current.workbook,
          )
          const longestLine = text
            .split('\n')
            .reduce((m, line) => Math.max(m, line.length), 0)
          const prev = maxCharsByCol.get(pos.column) ?? 6
          if (longestLine > prev) maxCharsByCol.set(pos.column, longestLine)
        }
        const nextColumnWidths: Record<string, number> = { ...sheet.columnWidths }
        for (const c of columns) {
          const maxChars = maxCharsByCol.get(c) ?? 6
          nextColumnWidths[String(c)] = Math.min(360, Math.max(72, maxChars * 8 + 24))
        }
        return { ...sheet, columnWidths: nextColumnWidths }
      }, `Auto-fitted column width for ${range}`)
      refreshUniverSurface()
    }
    const editor = univerEditorRef.current
    if (!editor?.autoFitColumns) {
      fallback()
      return
    }
    void editor
      .autoFitColumns(activeSheet.id, columns)
      .then(applied => {
        if (applied) setStatusMessage(`Auto-fitted column width for ${range}`)
        else fallback()
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
    return sortedRangeCellValues(sheet, range)
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
        else if (
          cell.kind === 'number' &&
          Number.isFinite(Number(cell.value)) &&
          (!Number.isInteger(Number(cell.value)) ||
            Number.isSafeInteger(Number(cell.value)))
        ) {
          value = Number(cell.value)
        } else {
          value = cell.value
        }
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
    const sortOpts = { hasHeader: sortHasHeader }
    // No canvas (tests, mount failure) or the bridge refused the write: sort
    // the model and rebuild the surface from it.
    const modelFallback = (): void => {
      updateActiveSheetFromCurrent(
        sheet =>
          sortRange(
            sheet,
            range,
            direction,
            documentRef.current.workbook,
            sortOpts,
          ),
        message,
      )
      refreshUniverSurface()
    }
    const editor = univerEditorRef.current
    if (!editor) {
      modelFallback()
      return
    }
    const flushed = await editor.flushSnapshot().catch(() => null)
    const sourceDocument = flushed
      ? documentFromEditorSnapshot(flushed, documentRef.current)
      : documentRef.current
    const sourceSheet =
      sourceDocument.workbook.sheets.find(
        sheet => sheet.id === activeSheet.id,
      ) ?? activeSheet
    const nextSheet = sortRange(
      sourceSheet,
      range,
      direction,
      sourceDocument.workbook,
      sortOpts,
    )
    if (nextSheet === sourceSheet) return
    // Value-only bridge writes drop moved cell styles and comments; refresh
    // via model when sorted cells carry styles or comments.
    const rangeKeys = cellsInRange(range, 50_000)
    const hasStylesOrComments = rangeKeys.some(
      key =>
        Boolean(sourceSheet.cells[key]?.style) ||
        Boolean(sourceSheet.comments?.[key]),
    )
    if (hasStylesOrComments) {
      updateActiveSheet(nextSheet, message)
      refreshUniverSurface()
      return
    }
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
    const nextSheet = createChartFromRange(
      activeSheet,
      range,
      type,
      document.workbook,
    )
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
    const next = updateChart(activeSheet, chartId, patch, document.workbook)
    if (next === activeSheet) {
      setStatusMessage('No plottable data in that range')
      return
    }
    updateActiveSheet(next, 'Updated chart')
  }

  // ---- Import / export (Phase S4) -------------------------------------------

  function uniqueSheetName(base: string): string {
    const names = new Set(
      documentRef.current.workbook.sheets.map(sheet => sheet.name.toLowerCase()),
    )
    if (!names.has(base.toLowerCase())) return base
    let suffix = 2
    while (names.has(`${base} ${suffix}`.toLowerCase())) suffix += 1
    return `${base} ${suffix}`
  }

  function handleImportExport(action: ImportExportAction): void {
    if (action === 'print-pdf') {
      printActiveSheet()
      return
    }
    const run =
      action === 'import-csv'
        ? importCsvFile
        : action === 'export-csv'
        ? exportCsvFile
        : action === 'export-selection-csv'
        ? () => exportSelectionFile('csv')
        : action === 'export-selection-tsv'
        ? () => exportSelectionFile('tsv')
        : action === 'export-selection-json'
        ? () => exportSelectionFile('json')
        : action === 'import-xlsx'
        ? importXlsxFile
        : exportXlsxFile
    void run().catch(error => {
      setStatusMessage(error instanceof Error ? error.message : String(error))
    })
  }

  async function exportSelectionFile(
    format: 'csv' | 'tsv' | 'json',
  ): Promise<void> {
    const range = activeRange()
    if (!range) return
    const version = transitions.version
    const sheetId = activeSheet.id
    const ext = format
    const label = format.toUpperCase()
    const path = await chooseSaveFilePath({
      defaultName: `${activeSheet.name || 'Sheet'}-${range.replace(':', '-')}.${ext}`,
      filters: [{ name: label, extensions: [ext] }],
    })
    if (!path) return
    await captureEditorWorkbook()
    if (version !== transitions.version)
      throw Error('The workbook changed before export. Export it again.')
    const sheet = documentRef.current.workbook.sheets.find(
      candidate => candidate.id === sheetId,
    )
    if (!sheet) throw Error('The sheet was removed before export.')
    const content =
      format === 'json'
        ? jsonFromSheet(sheet, range, documentRef.current.workbook)
        : csvFromSheet(
            sheet,
            range,
            documentRef.current.workbook,
            format === 'tsv' ? '\t' : ',',
          )
    await writeTextFile(path, content)
    setStatusMessage(
      `Exported ${range} from ${sheet.name} as ${label} to ${fileNameFromPath(path)}`,
    )
    recordLedgerOperation({
      lane: 'user',
      kind: `sheets.export.${format}`,
      summary: `Exported ${range} from ${sheet.name} as ${label}`,
      refs: { path },
    })
  }

  async function importCsvFile(): Promise<void> {
    const version = transitions.version
    const path = await openImportFile()
    if (!path) return
    const text = await readTextFile(path)
    if (version !== transitions.version)
      throw Error(
        'The workbook changed while the CSV was loading. Import it again.',
      )
    const base =
      fileNameFromPath(path)
        .replace(/\.[^.]+$/, '')
        .trim() || 'Imported'
    const sheet = sheetFromCsv(
      text,
      `sheet-csv-${Date.now()}`,
      uniqueSheetName(base),
      path.toLowerCase().endsWith('.tsv') ? { delimiter: '\t' } : {},
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
    const version = transitions.version
    const sheetId = activeSheet.id
    const path = await chooseSaveFilePath({
      defaultName: `${activeSheet.name || 'Sheet'}.csv`,
      filters: [{ name: 'CSV', extensions: ['csv'] }],
    })
    if (!path) return
    await captureEditorWorkbook()
    if (version !== transitions.version)
      throw Error('The workbook changed before export. Export it again.')
    const sheet = documentRef.current.workbook.sheets.find(
      sheet => sheet.id === sheetId,
    )
    if (!sheet) throw Error('The sheet was removed before export.')
    await writeTextFile(path, csvFromSheet(sheet, undefined, documentRef.current.workbook))
    setStatusMessage(
      `Exported ${sheet.name} to ${fileNameFromPath(path)}`,
    )
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.export.csv',
      summary: `Exported sheet ${sheet.name} as CSV`,
      refs: { path },
    })
  }

  async function importXlsxFile(): Promise<void> {
    const version = transitions.version
    const path = await openImportFile()
    if (!path) return
    if (version !== transitions.version) throw Error('The workbook changed while choosing an import. Try again.')
    await importXlsxPath(path)
  }

  async function importXlsxPath(path: string): Promise<boolean> {
    if (!/\.xlsx$/i.test(path)) throw Error('Choose an .xlsx file to import.')
    const applied = await transitions.run(
      async () => {
        setStatusMessage(`Importing ${fileNameFromPath(path)}…`)
        const binary = await readBinaryFile(path, 30_000_000)
        if (binary.truncated || !binary.base64)
          throw Error(
            'That workbook is too large to import (30 MB limit) or unreadable',
          )
        return importXlsxWorkbook(
          base64ToBytes(binary.base64),
          fileNameFromPath(path).replace(/\.xlsx$/i, '') || 'Imported workbook',
        )
      },
      imported => {
        installNewWorkbook(imported.document)
        setSaveState('dirty')
        importWarningsRef.current = imported.warnings
        setWarningsTitle('Imported with limitations')
        setImportWarnings(imported.warnings)
        setStatusMessage(`Imported ${fileNameFromPath(path)}`)
      },
    )
    if (!applied) return false
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.import.xlsx',
      summary: `Imported XLSX workbook ${fileNameFromPath(path)}`,
      refs: { path },
    })
    return true
  }

  /**
   * A CSV (or TSV) opened as a document, from the desk, a shared file or an agent's openWorkbook: a new workbook of
   * one sheet named after the file. The file itself is left as it is, as an Excel import leaves its file; it was
   * read as workbook JSON before, and refused with a parse error.
   */
  async function importCsvPath(path: string): Promise<boolean> {
    const applied = await transitions.run(
      async () => {
        setStatusMessage(`Importing ${fileNameFromPath(path)}…`)
        return workbookFromCsv(await readTextFile(path), fileNameFromPath(path))
      },
      imported => {
        installNewWorkbook(imported)
        setSaveState('dirty')
        setStatusMessage(`Imported ${fileNameFromPath(path)}`)
      },
    )
    if (!applied) return false
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.import.csv',
      summary: `Imported CSV ${fileNameFromPath(path)}`,
      refs: { path },
    })
    return true
  }

  async function exportXlsxFile(): Promise<void> {
    const path = await chooseSaveFilePath({
      defaultName: `${document.metadata.title || 'Workbook'}.xlsx`,
      filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }],
    })
    if (!path) return
    await exportXlsxToPath(path)
  }

  /**
   * Save as Excel: the menu and the agent's saveWorkbookAsXlsx both land here. The workbook stays bound to its own
   * file; the .xlsx is a copy, and what Excel will not get is listed in the notice and returned.
   */
  async function exportXlsxToPath(path: string): Promise<{ path: string; sheets: string[]; warnings: string[] }> {
    const version = transitions.version
    await captureEditorWorkbook()
    if (version !== transitions.version)
      throw Error('The workbook changed before saving as Excel. Save it again.')
    const docToExport = documentRef.current
    const { bytes, warnings } = await exportXlsxWorkbook(docToExport)
    await writeBinaryFile(path, bytesToBase64(bytes))
    setStatusMessage(`Saved a copy as ${fileNameFromPath(path)}`)
    if (warnings.length) {
      setWarningsTitle('Saved as Excel with limitations')
      setImportWarnings(warnings)
    }
    recordLedgerOperation({
      lane: 'user',
      kind: 'sheets.export.xlsx',
      summary: `Saved workbook "${docToExport.metadata.title}" as Excel`,
      refs: { path },
    })
    return {
      path,
      sheets: docToExport.workbook.sheets.filter(sheet => !sheet.agentDraft).map(sheet => sheet.name),
      warnings,
    }
  }

  async function openAnalysis(mode: AnalysisMode): Promise<void> {
    const version = transitions.version
    await captureEditorWorkbook()
    if (version !== transitions.version) throw new Error('The workbook changed. Open Analyze again.')
    const source = documentRef.current
    const sheet = source.workbook.sheets.find(sheet => sheet.id === source.workbook.activeSheetId)
    if (!sheet) return
    setAnalysis({ mode, source, sheet, range: normalizeRangeToken(selectedRange) ?? `${selectedCell}:${selectedCell}`, cell: selectedCell, version })
  }

  function navigateAnalysis(sheetId: string, range: string): void {
    setAnalysis(null)
    const cell = range.split(':')[0]
    const editor = univerEditorRef.current
    if (!editor) {
      selectSheet(sheetId)
      selectCell(cell)
      changeSelectedRange(normalizeRangeToken(range) ?? range, { manual: true })
      return
    }
    void (async () => {
      if (sheetId !== activeSheet.id) await editor.selectSheet(sheetId)
      await editor.revealCell(cell)
      await editor.selectRange(normalizeRangeToken(range) ?? range)
    })().catch(reportFailure)
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
  // up-to-date active sheet; never marks the workbook dirty when nothing changed.
  async function committedActiveSheet(): Promise<WorkbookSheet> {
    const editor = univerEditorRef.current
    if (!editor) return activeSheet
    if (editor.isCellEditing?.()) {
      await editor.commitEditing?.().catch(() => undefined)
    }
    const flushed = await editor.flushSnapshot?.().catch(() => null)
    if (!flushed) return activeSheet
    try {
      const current = documentRef.current
      const doc = documentFromEditorSnapshot(flushed, current)
      const contentChanged =
        JSON.stringify(doc.workbook.sheets) !==
        JSON.stringify(current.workbook.sheets)
      if (contentChanged) {
        setDocument(doc)
        setSaveState('dirty')
      }
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
    const matches = findMatches(
      sheet,
      query,
      { matchCase },
      documentRef.current.workbook,
    )
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

  async function findPrevious(): Promise<void> {
    const query = findText.trim()
    if (!query) return
    const sheet = await committedActiveSheet()
    const matches = findMatches(
      sheet,
      query,
      { matchCase },
      documentRef.current.workbook,
    )
    if (!matches.length) {
      setStatusMessage(`No matches for ${findText}`)
      return
    }
    const currentIndex = matches.indexOf(selectedCell)
    const prevIndex =
      currentIndex <= 0 ? matches.length - 1 : currentIndex - 1
    const prev = matches[prevIndex]
    lastFindRef.current = prev
    selectCell(prev)
    void univerEditorRef.current?.revealCell?.(prev)
    setStatusMessage(`Found ${prev} (${prevIndex + 1}/${matches.length})`)
  }

  // Replace ONE occurrence at a time: the first match in the current cell, then
  // (once that cell has no matches left) the next matching cell (#277).
  async function replaceNext(): Promise<void> {
    const query = findText.trim()
    if (!query) return
    const sheet = await committedActiveSheet()
    const matches = findMatches(
      sheet,
      query,
      { matchCase },
      documentRef.current.workbook,
    )
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
    const visible = new Set(
      visibleRowsForSheet(sheet, documentRef.current.workbook),
    )
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
            .map(key => computedCellValue(activeSheet, key, document.workbook).trim())
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
      : displayCellValue(activeSheet, selectedCell, document.workbook)
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
        const refreshedSheets = next.workbook.sheets.map(s =>
          s.charts?.length ? recomputeSheetCharts(s, next.workbook) : s,
        )
        const anyChartUpdated = refreshedSheets.some(
          (s, idx) => s !== next.workbook.sheets[idx],
        )
        return anyChartUpdated
          ? { ...next, workbook: { ...next.workbook, sheets: refreshedSheets } }
          : next
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
    sheet => !sheet.agentDraft && !sheet.hidden,
  )
  const hiddenSheets = document.workbook.sheets.filter(
    sheet => !sheet.agentDraft && Boolean(sheet.hidden),
  )

  // Format painter: when active, clicking/selecting a different cell or range applies the painted style once.
  const prevPainterTargetRef = useRef(`${selectedCell}|${selectedRange}`)
  useEffect(() => {
    const signature = `${selectedCell}|${selectedRange}`
    if (formatPainterStyle && signature !== prevPainterTargetRef.current) {
      const target = normalizeRangeToken(selectedRange) ?? `${selectedCell}:${selectedCell}`
      applyRangeStyle(target, formatPainterStyle, `Painted formatting to ${target}`)
      setFormatPainterStyle(null)
    }
    prevPainterTargetRef.current = signature
  }, [selectedCell, selectedRange, formatPainterStyle])

  // --- Wave 1–5 Parity Action Handlers ---
  function applyBorderPreset(preset: BorderPreset, color?: string): void {
    const range = activeRange()
    if (!range) return
    if (preset === 'none') {
      applyRangeStyle(
        range,
        { border: false, borderSides: undefined, borderStyle: undefined, borderColor: undefined },
        'Cleared borders',
      )
      return
    }
    if (preset === 'all' || preset === 'outer') {
      applyRangeStyle(
        range,
        { border: true, borderSides: undefined, borderStyle: 'thin', borderColor: color },
        'Applied borders',
      )
      return
    }
    if (preset === 'double-bottom') {
      applyRangeStyle(
        range,
        { border: false, borderSides: ['bottom'], borderStyle: 'double', borderColor: color },
        'Applied double bottom border',
      )
      return
    }
    const sidesMap: Record<
      Exclude<BorderPreset, 'all' | 'outer' | 'none' | 'double-bottom'>,
      Array<'top' | 'right' | 'bottom' | 'left'>
    > = {
      bottom: ['bottom'],
      top: ['top'],
      'top-bottom': ['top', 'bottom'],
      left: ['left'],
      right: ['right'],
    }
    applyRangeStyle(
      range,
      { border: false, borderSides: sidesMap[preset], borderStyle: 'thin', borderColor: color },
      `Applied ${preset} border`,
    )
  }

  function toggleFormatPainter(): void {
    if (formatPainterStyle) {
      setFormatPainterStyle(null)
      setStatusMessage('Cancelled format painter')
      return
    }
    const style = activeSheet.cells[selectedCell]?.style
    if (!style) {
      setStatusMessage('Selected cell has no formatting to copy')
      return
    }
    setFormatPainterStyle({ ...style })
    setStatusMessage('Format painter active — select target cells')
  }

  function applyCustomFreeze(rows: number, columns: number): void {
    updateActiveSheetFromCurrent(
      sheet => setFreezePanes(sheet, rows, columns),
      rows || columns
        ? `Frozen ${rows} row(s) and ${columns} column(s)`
        : 'Unfroze all panes',
    )
    const editor = univerEditorRef.current
    if (editor) {
      void editor.setFrozenRows(rows)
      void editor.setFrozenColumns(columns)
    }
  }

  function applyMultiSort(keys: SortKey[], hasHeader: boolean): void {
    const range = activeRange()
    if (!range) return
    updateActiveSheetFromCurrent(
      sheet => sortRangeMulti(sheet, range, keys, hasHeader, documentRef.current.workbook),
      `Sorted ${range} (${keys.length} level${keys.length > 1 ? 's' : ''})`,
    )
    refreshUniverSurface()
  }

  function applySmartColumnFilter(
    column: number,
    rule: ColumnFilterRule | null,
  ): void {
    const next = setColumnFilter(activeSheet, column, rule)
    updateActiveSheet(
      next,
      rule ? `Filtered column ${columnLabel(column)}` : `Cleared filter on column ${columnLabel(column)}`,
    )
    applyFilterToEditor(next)
  }

  function clearAllSmartFilters(): void {
    const next = clearAllColumnFilters(activeSheet)
    updateActiveSheet(next, 'Cleared all filters')
    applyFilterToEditor(next)
  }

  function runRemoveDuplicates(hasHeader: boolean): void {
    const range = activeRange()
    if (!range) return
    const result = removeDuplicatesInRange(
      activeSheet,
      range,
      undefined,
      hasHeader,
      document.workbook,
    )
    if (result.removedCount === 0) {
      setStatusMessage('No duplicate rows found')
      showToast('No duplicate rows found')
      return
    }
    updateActiveSheet(
      result.sheet,
      `Removed ${result.removedCount} duplicate row${result.removedCount === 1 ? '' : 's'}`,
    )
    refreshUniverSurface()
    showToast(`Removed ${result.removedCount} duplicate row${result.removedCount === 1 ? '' : 's'}`)
  }

  function runSplitTextToColumns(delimiter: string): void {
    const range = activeRange()
    if (!range) return
    const result = splitTextToColumns(activeSheet, range, delimiter)
    if (result.rowsSplit === 0) {
      setStatusMessage(`No cells contained "${delimiter}"`)
      return
    }
    updateActiveSheet(
      result.sheet,
      `Split ${result.rowsSplit} row${result.rowsSplit === 1 ? '' : 's'} into up to ${result.maxParts} columns`,
    )
    refreshUniverSurface()
  }

  function runTrimWhitespace(): void {
    const range = activeRange()
    if (!range) return
    const result = trimWhitespaceInRange(activeSheet, range)
    if (result.trimmedCount === 0) {
      setStatusMessage('No extra whitespace found')
      return
    }
    updateActiveSheet(
      result.sheet,
      `Trimmed whitespace in ${result.trimmedCount} cell${result.trimmedCount === 1 ? '' : 's'}`,
    )
    refreshUniverSurface()
  }

  function runToggleCheckboxes(): void {
    const range = activeRange()
    if (!range) return
    updateActiveSheetFromCurrent(
      sheet => toggleCheckboxInRange(sheet, range),
      `Toggled checkboxes in ${range}`,
    )
    refreshUniverSurface()
  }

  function runSetHyperlink(url: string, label?: string): void {
    updateActiveSheetFromCurrent(
      sheet => setCellHyperlink(sheet, selectedCell, url, label),
      url.trim() ? `Linked ${selectedCell}` : `Removed link on ${selectedCell}`,
    )
    refreshUniverSurface()
  }

  function pasteTransposed(): void {
    const copied = copiedBlockRef.current
    if (!copied) {
      setStatusMessage('Copy a range first before Paste Transpose')
      return
    }
    const transposed = transposeCopiedBlock(copied)
    updateActiveSheetFromCurrent(
      sheet => pasteCellBlock(sheet, selectedCell, transposed, 'all'),
      `Pasted transposed at ${selectedCell}`,
    )
    refreshUniverSurface()
  }

  function hideSelectedRows(): void {
    const range = activeRange()
    const bounds = range ? range.split(':').map(parseCellKey) : []
    const startRow = bounds[0]?.row ?? selectedPosition?.row ?? 0
    const endRow = bounds[1]?.row ?? startRow
    const rows = Array.from(
      { length: Math.abs(endRow - startRow) + 1 },
      (_, i) => Math.min(startRow, endRow) + i,
    )
    const next = hideRowsInSheet(activeSheet, rows)
    updateActiveSheet(next, `Hid ${rows.length} row${rows.length === 1 ? '' : 's'}`)
    applyFilterToEditor(next)
  }

  function unhideAllRows(): void {
    const next = unhideAllRowsInSheet(activeSheet)
    updateActiveSheet(next, 'Unhid all rows')
    applyFilterToEditor(next)
  }

  function hideSelectedColumns(): void {
    const range = activeRange()
    const bounds = range ? range.split(':').map(parseCellKey) : []
    const startCol = bounds[0]?.column ?? selectedPosition?.column ?? 0
    const endCol = bounds[1]?.column ?? startCol
    const cols = Array.from(
      { length: Math.abs(endCol - startCol) + 1 },
      (_, i) => Math.min(startCol, endCol) + i,
    )
    updateActiveSheetFromCurrent(
      sheet => hideColumnsInSheet(sheet, cols),
      `Hid ${cols.length} column${cols.length === 1 ? '' : 's'}`,
    )
    void univerEditorRef.current?.hideColumns?.(
      Math.min(startCol, endCol),
      cols.length,
    )
  }

  function unhideAllColumns(): void {
    updateActiveSheetFromCurrent(unhideAllColumnsInSheet, 'Unhid all columns')
    void univerEditorRef.current?.showColumns?.(0, activeSheet.columnCount)
  }

  function toggleGridlines(): void {
    const nextVisible = activeSheet.showGridlines === false
    updateActiveSheetFromCurrent(
      toggleSheetGridlines,
      nextVisible ? 'Gridlines shown' : 'Gridlines hidden',
    )
    void univerEditorRef.current?.setShowGridlines?.(nextVisible)
  }

  function changeZoom(ratio: number): void {
    setZoom(ratio)
    void univerEditorRef.current?.setZoom?.(ratio)
    setStatusMessage(`Zoom ${Math.round(ratio * 100)}%`)
  }

  function printActiveSheet(): void {
    const html = sheetsSnapshotHtml(document)
    if (typeof window === 'undefined') return
    const win = window.open('', '_blank', 'width=960,height=720')
    if (win) {
      win.document.write(
        `<!doctype html><html><head><title>${document.metadata.title} — ${activeSheet.name}</title>` +
          `<style>body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:24px;color:#14171a;}table{border-collapse:collapse;width:100%;font-size:11px;}td,th{border:1px solid #c8ccd2;padding:4px 6px;}@media print{@page{size:landscape;margin:12mm;}}</style>` +
          `</head><body><h2>${document.metadata.title} — ${activeSheet.name}</h2>${html || ''}<script>window.onload=()=>window.print()</script></body></html>`,
      )
      win.document.close()
    } else {
      window.print?.()
    }
    setStatusMessage('Opened print / PDF preview')
  }

  function saveNamedRange(name: string, range: string, comment?: string): void {
    const norm = normalizeRangeToken(range)
    if (!norm) {
      setStatusMessage(`Invalid range ${range}`)
      return
    }
    const next = upsertNamedRange(documentRef.current, {
      name,
      sheetId: activeSheet.id,
      range: norm,
      comment,
    })
    if (next === documentRef.current) {
      setStatusMessage(`Invalid named range "${name}" (use letters, digits, _ — not a cell key like A1)`)
      return
    }
    commitDocument(() => next, `Defined named range ${name} (${norm})`)
  }

  function removeNamedRangeById(id: string): void {
    commitDocument(
      current => deleteNamedRange(current, id),
      'Deleted named range',
    )
  }

  function jumpToCell(sheetId: string, rangeOrCell: string): void {
    const norm = normalizeRangeToken(rangeOrCell) ?? rangeOrCell
    const firstCell = norm.split(':')[0] || 'A1'
    if (sheetId !== activeSheet.id) {
      selectSheet(sheetId)
    }
    setSelectedCell(firstCell)
    setSelectedRange(norm.includes(':') ? norm : `${firstCell}:${firstCell}`)
    void univerEditorRef.current?.revealCell?.(firstCell)
  }

  function jumpToNamedRange(nr: NamedRange): void {
    jumpToCell(nr.sheetId, nr.range)
    setStatusMessage(`Jumped to ${nr.name} (${nr.range})`)
  }

  function addConditionalRuleFull(rule: Omit<ConditionalFormatRule, 'id'>): void {
    updateActiveSheetFromCurrent(
      sheet => addConditionalFormatRule(sheet, rule),
      `Added conditional rule on ${rule.range}`,
    )
    refreshUniverSurface()
  }

  function deleteConditionalRuleById(ruleId: string): void {
    updateActiveSheetFromCurrent(
      sheet => removeConditionalFormatRule(sheet, ruleId),
      'Removed conditional format rule',
    )
    refreshUniverSurface()
  }

  function moveConditionalRuleById(ruleId: string, delta: -1 | 1): void {
    updateActiveSheetFromCurrent(
      sheet => moveConditionalFormatRule(sheet, ruleId, delta),
      'Reordered conditional format rules',
    )
    refreshUniverSurface()
  }

  function addValidationRuleFull(rule: Omit<ValidationRule, 'id'>): void {
    updateActiveSheetFromCurrent(
      sheet => addValidationRule(sheet, rule),
      `Added ${rule.kind} validation on ${rule.range}`,
    )
  }

  function deleteValidationRuleById(ruleId: string): void {
    updateActiveSheetFromCurrent(
      sheet => removeValidationRule(sheet, ruleId),
      'Removed validation rule',
    )
  }

  function createPivotFromSelection(): void {
    if (activeSheet.pivot) {
      setInspectorTab('pivot')
      setInspectorOpen(true)
      return
    }
    const range = activeRange()
    if (!range) return
    const fields = inferPivotFields(activeSheet, range, document.workbook)
    if (!fields.length) {
      setStatusMessage('Select a table range with headers to build a Pivot Table')
      return
    }
    const firstCat = fields.find(f => !f.isNumeric) ?? fields[0]
    const firstNum = fields.find(f => f.isNumeric) ?? fields[fields.length - 1]
    const def: PivotDefinition = {
      id: `pivot-${Date.now()}`,
      sourceSheetId: activeSheet.id,
      sourceRange: range,
      rows: [{ columnIndex: firstCat.columnIndex, label: firstCat.label }],
      columns: [],
      values: [
        {
          columnIndex: firstNum.columnIndex,
          label: firstNum.label,
          aggregate: firstNum.isNumeric ? 'sum' : 'count',
        },
      ],
      refreshedAt: new Date().toISOString(),
    }
    const pivotName = uniqueSheetName(`Pivot · ${activeSheet.name}`)
    const pivotSheet = computePivotSheet(
      activeSheet,
      def,
      { id: `sheet-pivot-${Date.now()}`, name: pivotName },
      document.workbook,
    )
    commitDocument(
      current => addSheetToWorkbook(current, pivotSheet),
      `Created Pivot Table "${pivotName}"`,
    )
    refreshUniverSurface()
    setInspectorTab('pivot')
    setInspectorOpen(true)
  }

  function updatePivotDefinition(def: PivotDefinition): void {
    const source = document.workbook.sheets.find(s => s.id === def.sourceSheetId)
    if (!source) {
      setStatusMessage('Pivot source sheet no longer exists')
      return
    }
    const recomputed = computePivotSheet(
      source,
      def,
      { id: activeSheet.id, name: activeSheet.name },
      document.workbook,
    )
    updateActiveSheet(recomputed, 'Updated Pivot Table')
    refreshUniverSurface()
  }

  function refreshActivePivot(): void {
    if (!activeSheet.pivot) return
    updatePivotDefinition(activeSheet.pivot)
    showToast('Pivot Table refreshed')
  }

  function moveSheet(sheetId: string, delta: -1 | 1): void {
    const idx = document.workbook.sheets.findIndex(s => s.id === sheetId)
    if (idx < 0) return
    commitDocument(
      current => moveSheetInWorkbook(current, sheetId, idx + delta),
      'Reordered sheets',
    )
    refreshUniverSurface()
  }

  function changeSheetTabColor(sheetId: string, color: string | undefined): void {
    commitDocument(current => ({
      ...current,
      workbook: {
        ...current.workbook,
        sheets: current.workbook.sheets.map(s =>
          s.id === sheetId ? setSheetTabColor(s, color) : s,
        ),
      },
    }), color ? 'Updated tab color' : 'Cleared tab color')
    void univerEditorRef.current?.setTabColor?.(sheetId, color ?? '')
  }

  function hideSheetById(sheetId: string): void {
    commitDocument(
      current => setSheetHidden(current, sheetId, true),
      'Hid sheet',
    )
    refreshUniverSurface()
  }

  function unhideSheetById(sheetId: string): void {
    commitDocument(
      current => ({
        ...setSheetHidden(current, sheetId, false),
        workbook: {
          ...setSheetHidden(current, sheetId, false).workbook,
          activeSheetId: sheetId,
        },
      }),
      'Shown hidden sheet',
    )
    refreshUniverSurface()
  }

  function revertAgentEntry(entry: AgentLogEntry): void {
    if (!entry.cells?.length) return
    commitDocument(current => {
      const target = current.workbook.sheets.find(s => s.id === entry.sheetId)
      if (!target) return current
      let nextSheet = target
      for (const change of entry.cells!) {
        nextSheet = upsertCell(nextSheet, change.cell, change.before ?? '')
      }
      return cloneWithUpdatedSheet(current, nextSheet)
    }, `Reverted "${entry.summary}"`)
    refreshUniverSurface()
    showToast(`Reverted ${entry.cells.length} cell change(s)`)
  }

  const activeFilterCount = activeFilterColumnIndices(activeSheet).length
  const activePivotSource = activeSheet.pivot
    ? document.workbook.sheets.find(s => s.id === activeSheet.pivot!.sourceSheetId)
    : undefined
  const activePivotStale = activeSheet.pivot
    ? isPivotStale(activeSheet, activePivotSource, document.workbook)
    : false
  const selectionHasMerge = useMemo(
    () =>
      normalizedRange
        ? hasMergeInRange(activeSheet, normalizedRange)
        : false,
    [activeSheet, normalizedRange],
  )

  const paletteCommands = useMemo<PaletteCommand[]>(() => {
    const cmds: PaletteCommand[] = [
      { id: 'bold', category: 'Format', label: 'Toggle Bold', shortcut: '⌘B', run: () => toggleStyle('bold') },
      { id: 'italic', category: 'Format', label: 'Toggle Italic', shortcut: '⌘I', run: () => toggleStyle('italic') },
      { id: 'underline', category: 'Format', label: 'Toggle Underline', shortcut: '⌘U', run: () => toggleStyle('underline') },
      { id: 'strike', category: 'Format', label: 'Toggle Strikethrough', run: () => toggleStyle('strikethrough') },
      { id: 'wrap', category: 'Format', label: 'Toggle Wrap Text', run: () => applyCellStyle({ wrap: selectedStyle?.wrap === 'wrap' ? 'overflow' : 'wrap' }) },
      { id: 'merge', category: 'Format', label: selectionHasMerge ? 'Unmerge Cells' : 'Merge Selected Cells', run: toggleMergeSelectedRange },
      { id: 'clear-fmt', category: 'Format', label: 'Clear Formatting', run: clearSelectedFormatting },
      { id: 'painter', category: 'Format', label: 'Format Painter', run: toggleFormatPainter },
      { id: 'cur-usd', category: 'Number', label: 'Format as Currency (USD $)', run: () => applyCellStyle({ numberFormat: 'currency', currencySymbol: '$' }) },
      { id: 'cur-eur', category: 'Number', label: 'Format as Currency (EUR €)', run: () => applyCellStyle({ numberFormat: 'currency', currencySymbol: '€' }) },
      { id: 'cur-gbp', category: 'Number', label: 'Format as Currency (GBP £)', run: () => applyCellStyle({ numberFormat: 'currency', currencySymbol: '£' }) },
      { id: 'acct', category: 'Number', label: 'Format as Accounting', run: () => applyCellStyle({ numberFormat: 'accounting' }) },
      { id: 'pct', category: 'Number', label: 'Format as Percent (%)', run: () => applyCellStyle({ numberFormat: 'percent' }) },
      { id: 'sci', category: 'Number', label: 'Format as Scientific', run: () => applyCellStyle({ numberFormat: 'scientific' }) },
      { id: 'border-dbl', category: 'Format', label: 'Double Bottom Border (Accounting Total)', run: () => applyBorderPreset('double-bottom') },
      { id: 'pivot', category: 'Data', label: 'Insert Pivot Table Sheet', keywords: 'summarize group aggregate', run: createPivotFromSelection },
      { id: 'dedupe', category: 'Data', label: 'Remove Duplicate Rows', run: () => runRemoveDuplicates(true) },
      { id: 'split', category: 'Data', label: 'Split Text to Columns (Comma)', run: () => runSplitTextToColumns(',') },
      { id: 'trim', category: 'Data', label: 'Trim Whitespace in Selection', run: runTrimWhitespace },
      { id: 'checkbox', category: 'Data', label: 'Toggle Checkboxes in Selection', run: runToggleCheckboxes },
      { id: 'transpose', category: 'Edit', label: 'Paste Transpose', run: pasteTransposed },
      { id: 'paste-val', category: 'Edit', label: 'Paste Values Only', shortcut: '⌘⇧V', run: () => pasteSpecial('values') },
      { id: 'fill-down', category: 'Edit', label: 'Fill Down', shortcut: '⌘D', run: () => fillSelectedRange('down') },
      { id: 'fill-right', category: 'Edit', label: 'Fill Right', shortcut: '⌘R', run: () => fillSelectedRange('right') },
      { id: 'autofit-col', category: 'Edit', label: 'Auto-Fit Selected Columns', run: autoFitSelectedColumns },
      { id: 'autofit-row', category: 'Edit', label: 'Auto-Fit Selected Rows', run: autoFitSelectedRows },
      { id: 'freeze-here', category: 'View', label: `Freeze Panes Up To ${selectedCell}`, run: () => applyCustomFreeze((selectedPosition?.row ?? 0) + 1, (selectedPosition?.column ?? 0) + 1) },
      { id: 'unfreeze', category: 'View', label: 'Unfreeze All Panes', run: () => applyCustomFreeze(0, 0) },
      { id: 'gridlines', category: 'View', label: activeSheet.showGridlines === false ? 'Show Gridlines' : 'Hide Gridlines', run: toggleGridlines },
      { id: 'show-formulas', category: 'View', label: showFormulas ? 'Show Values' : 'Show Formulas', shortcut: 'Ctrl+`', run: () => setShowFormulas(v => !v) },
      { id: 'find', category: 'Edit', label: 'Find and Replace', shortcut: '⌘F', run: () => setFindOpen(true) },
      { id: 'print', category: 'File', label: 'Print / Export PDF', run: printActiveSheet },
      { id: 'insp-rules', category: 'Inspector', label: 'Open Conditional Formatting Rules Manager', run: () => { setInspectorTab('rules'); setInspectorOpen(true) } },
      { id: 'insp-valid', category: 'Inspector', label: 'Open Data Validation Manager', run: () => { setInspectorTab('validate'); setInspectorOpen(true) } },
      { id: 'insp-names', category: 'Inspector', label: 'Open Named Ranges Manager', run: () => { setInspectorTab('names'); setInspectorOpen(true) } },
      { id: 'insp-audit', category: 'Inspector', label: 'Open Formula Audit (Precedents & Dependents)', run: () => { setInspectorTab('audit'); setInspectorOpen(true) } },
      { id: 'insp-log', category: 'Inspector', label: 'Open Agent Activity & Revert Log', run: () => { setInspectorTab('activity'); setInspectorOpen(true) } },
      ...visibleSheets.map(s => ({
        id: `sheet-${s.id}`,
        category: 'Sheet',
        label: `Go to sheet: ${s.name}`,
        run: () => selectSheet(s.id),
      })),
      ...(document.workbook.namedRanges ?? []).map(nr => ({
        id: `nr-${nr.id}`,
        category: 'Name',
        label: `Jump to ${nr.name} (${nr.range})`,
        run: () => jumpToNamedRange(nr),
      })),
    ]
    return cmds
  }, [activeSheet, selectedCell, normalizedRange, selectedStyle, showFormulas, visibleSheets, document.workbook.namedRanges])
  // Data-validation feedback (Phase S3), computed from the LIVE workbook for
  // every sheet: the Univer surface paints these cells and the fallback grid
  // titles them. Flag, never reject — commits still land.
  const invalidFlagsBySheet = useMemo(() => {
    const flagged: Record<string, InvalidCellFlag[]> = {}
    for (const sheet of document.workbook.sheets) {
      const flags = invalidCellFlagsForSheet(sheet, document.workbook)
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
  const { rangeSummary, statPills } = useMemo(() => {
    if (!normalizedRange) return { rangeSummary: null, statPills: [] as StatPill[] }
    const keys = cellsInRange(normalizedRange, 20000)
    if (keys.length < 2) return { rangeSummary: null, statPills: [] as StatPill[] }
    const numbers: number[] = []
    for (const key of keys) {
      const value = computedCellValue(activeSheet, key, document.workbook)
      if (value.trim() === '') continue
      const numeric = Number(value)
      if (Number.isFinite(numeric)) numbers.push(numeric)
    }
    if (!numbers.length) return { rangeSummary: null, statPills: [] as StatPill[] }
    const sum = numbers.reduce((total, value) => total + value, 0)
    const avg = sum / numbers.length
    const min = Math.min(...numbers)
    const max = Math.max(...numbers)
    const compact = (value: number): string =>
      Number.isInteger(value) ? String(value) : String(Number(value.toFixed(4)))
    return {
      rangeSummary: `Sum ${compact(sum)} · Avg ${compact(avg)} · Min ${compact(min)} · Max ${compact(max)} · Count ${numbers.length}`,
      statPills: [
        { label: 'Sum', display: compact(sum), raw: String(sum) },
        { label: 'Avg', display: compact(avg), raw: String(avg) },
        { label: 'Count', display: String(numbers.length), raw: String(numbers.length) },
      ],
    }
  }, [activeSheet, normalizedRange, document.workbook])

  const findMatchesList = useMemo(
    () =>
      findText.trim()
        ? findMatches(activeSheet, findText, { matchCase }, document.workbook)
        : [],
    [activeSheet, document.workbook, findText, matchCase],
  )
  const findActiveIndex = findMatchesList.indexOf(selectedCell)

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
      inert={switching || undefined}
      aria-busy={switching}
      data-compact-room={compact || undefined}
      headerDocumentName={
        document.metadata.title?.trim() ||
        (lifecycle.doc.path
          ? lifecycle.doc.path.split(/[\\/]/).pop()
          : undefined)
      }
      headerActions={
        <DocumentHeaderActions
          lifecycle={lifecycle}
          title={document.metadata.title || 'Untitled'}
          onOpenSwitcher={() => setSwitcherOpen(true)}
        />
      }
    >
      {compact && <style>{COMPACT_ROOM_CSS}</style>}
      {toast && <Toast role="status">{toast}</Toast>}
      {resourceOpen.failed && (
        <ToolTextButton type="button" onClick={resourceOpen.retry}>
          Retry opening workbook
        </ToolTextButton>
      )}
      <Shell data-app="sheets">
        <DocumentSwitcher
          appSlug={PURESHEETS_APP_SLUG}
          suffixes={['.sheets', '.sheets.html']}
          variant="modal"
          loadPreview={loadWorkbookPreview}
          open={switcherOpen && !compact}
          onClose={() => setSwitcherOpen(false)}
          onOpenDocument={path => {
            setSwitcherOpen(false)
            void openWorkbookPath(path).catch(reportFailure)
          }}
          onCreateNew={() => void newWorkbook().catch(reportFailure)}
          newLabel="New workbook"
          title="Open a workbook"
          itemNoun="workbook"
          previewStyle="grid"
        />
        <TopBar data-compact-hidden={compact || undefined}>
          <TitleBlock>
            <FileNameInput
              aria-label="Workbook title"
              value={titleDraft ?? document.metadata.title}
              placeholder="Untitled"
              onFocus={() => {
                skipTitleBlurCommitRef.current = false
                setTitleDraft(document.metadata.title)
              }}
              onChange={event => setTitleDraft(event.currentTarget.value)}
              onBlur={event => {
                if (skipTitleBlurCommitRef.current) {
                  skipTitleBlurCommitRef.current = false
                  setTitleDraft(null)
                  return
                }
                commitWorkbookTitle(event.currentTarget.value)
              }}
              onKeyDown={event => {
                if (event.key === 'Enter') {
                  event.currentTarget.blur()
                } else if (event.key === 'Escape') {
                  event.preventDefault()
                  skipTitleBlurCommitRef.current = true
                  setTitleDraft(null)
                  event.currentTarget.blur()
                }
              }}
            />
          </TitleBlock>
          <CommandBar
            analysisActions={<ToolbarMenu label="Analyze" displayValue="Analyze" options={[...analysisOptions]} minWidth={165}
              onSelect={value => void openAnalysis(value as AnalysisMode).catch(reportFailure)} />}
            activeSheet={activeSheet}
            workbook={document.workbook}
            selectedCell={selectedCell}
            selectedRow={selectedPosition?.row ?? 0}
            selectedColumn={selectedPosition?.column ?? 0}
            selectedStyle={selectedStyle}
            isMerged={selectionHasMerge}
            formatPainterActive={Boolean(formatPainterStyle)}
            showGridlines={activeSheet.showGridlines !== false}
            showFormulas={showFormulas}
            inspectorOpen={inspectorOpen}
            activeFilterCount={activeFilterCount}
            canUndo={undoStack.length > 0}
            canRedo={redoStack.length > 0}
            findOpen={findOpen}
            onFindOpenChange={setFindOpen}
            findFocusRequest={findFocusRequest}
            findText={findText}
            replaceText={replaceText}
            findMatchCount={findMatchesList.length}
            findActiveIndex={findActiveIndex}
            onUndo={undo}
            onRedo={redo}
            onCopySelectedRange={copySelectedRange}
            onPasteFromClipboard={() => pasteFromClipboard()}
            onToggleStyle={toggleStyle}
            onApplyCellStyle={applyCellStyle}
            onApplyBorderPreset={applyBorderPreset}
            onClearSelectedRange={clearSelectedRange}
            onClearFormatting={clearSelectedFormatting}
            onToggleFormatPainter={toggleFormatPainter}
            onToggleMerge={toggleMergeSelectedRange}
            onSortSelectedRange={sortSelectedRange}
            sortHasHeader={sortHasHeader}
            onToggleSortHeader={() => setSortHasHeader(current => !current)}
            onMultiSort={applyMultiSort}
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
            onApplyColumnFilter={applySmartColumnFilter}
            onClearAllFilters={clearAllSmartFilters}
            onToggleFirstRowFreeze={toggleFirstRowFreeze}
            onToggleFirstColumnFreeze={toggleFirstColumnFreeze}
            onSetFreezePanes={applyCustomFreeze}
            firstRowFrozen={(activeSheet.frozenRows ?? 0) > 0}
            firstColumnFrozen={(activeSheet.frozenColumns ?? 0) > 0}
            onCreateChart={createChart}
            onCreatePivot={createPivotFromSelection}
            onInsertRow={insertSelectedRow}
            onDeleteRow={deleteSelectedRow}
            onInsertColumn={insertSelectedColumn}
            onDeleteColumn={deleteSelectedColumn}
            onAutoFitColumns={autoFitSelectedColumns}
            onAutoFitRows={autoFitSelectedRows}
            onHideRows={hideSelectedRows}
            onUnhideRows={unhideAllRows}
            onHideColumns={hideSelectedColumns}
            onUnhideColumns={unhideAllColumns}
            onToggleGridlines={toggleGridlines}
            onToggleShowFormulas={() => setShowFormulas(v => !v)}
            onPrintSheet={printActiveSheet}
            onResizeColumn={delta =>
              resizeSelectedColumn(selectedColumnWidth() + delta)
            }
            onResizeRow={delta =>
              resizeSelectedRow(selectedRowHeight() + delta)
            }
            onPasteValuesOnly={() => pasteSpecial('values')}
            onPasteFormattingOnly={() => pasteSpecial('formats')}
            onPasteTranspose={pasteTransposed}
            onFillDown={() => fillSelectedRange('down')}
            onFillRight={() => fillSelectedRange('right')}
            onAdjustDecimals={adjustDecimals}
            onApplyTableStyle={applyTableStyle}
            onInsertFormulaTemplate={insertFormulaTemplate}
            onRemoveDuplicates={runRemoveDuplicates}
            onSplitTextToColumns={runSplitTextToColumns}
            onTrimWhitespace={runTrimWhitespace}
            onToggleCheckboxes={runToggleCheckboxes}
            onSetHyperlink={runSetHyperlink}
            onFindTextChange={setFindText}
            onReplaceTextChange={setReplaceText}
            onFindNext={findNext}
            onFindPrevious={findPrevious}
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
            onToggleInspector={() => setInspectorOpen(open => !open)}
            onOpenCommandPalette={() => setPaletteOpen(true)}
            agentLogCount={(document.workbook.agentLog ?? []).length}
          />
        </TopBar>

        {analysis && !compact && <WorkbookAnalysisDialog
          mode={analysis.mode} document={analysis.source} sheet={analysis.sheet} range={analysis.range} cell={analysis.cell}
          onClose={() => setAnalysis(null)} onNavigate={navigateAnalysis}
          onApply={(next, message) => {
            if (analysis.version !== transitions.version || analysis.source !== documentRef.current)
              throw new Error('The workbook changed while this preview was open. Close it and preview again.')
            commitDocument(() => next, message)
            if (analysis.sheet.id !== next.workbook.activeSheetId) selectCell('A1')
            refreshUniverSurface()
            setAnalysis(null)
          }}
        />}

        {activeSheet.pivot && (
          <PivotBanner aria-label="Pivot table status">
            <span>
              <strong>Pivot Table</strong> from{' '}
              {activePivotSource?.name ?? 'deleted sheet'}!{activeSheet.pivot.sourceRange}
              {activePivotStale ? ' · Source changed — refresh to update' : ''}
            </span>
            <PivotBannerActions>
              <ImportWarningsDismiss type="button" onClick={refreshActivePivot}>
                Refresh
              </ImportWarningsDismiss>
              <ImportWarningsDismiss
                type="button"
                onClick={() => {
                  setInspectorTab('pivot')
                  setInspectorOpen(true)
                }}
              >
                Configure
              </ImportWarningsDismiss>
            </PivotBannerActions>
          </PivotBanner>
        )}

        {importWarnings && importWarnings.length > 0 && (
          <ImportWarningsBanner aria-label="Import warnings">
            <ImportWarningsTitle>{warningsTitle}</ImportWarningsTitle>
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
            selectedHyperlink={activeSheet.cells[selectedCell]?.hyperlink}
            selectedCheckbox={activeSheet.cells[selectedCell]?.checkbox}
            onToggleCheckboxValue={() => {
              const nextVal =
                activeSheet.cells[selectedCell]?.value.trim().toUpperCase() === 'TRUE'
                  ? 'FALSE'
                  : 'TRUE'
              commitSelectedCell(nextVal)
            }}
            namedRanges={document.workbook.namedRanges ?? []}
            onSelectNamedRange={jumpToNamedRange}
            onRangeChange={range =>
              changeSelectedRange(range, { manual: true })
            }
            onRangeInvalid={setStatusMessage}
            onEditValueChange={changeFormulaValue}
            onCommitSelectedCell={commitSelectedCell}
            onCancelEdit={cancelFormulaEdit}
          />

          {showFormulas && (
            <FormulasStrip aria-label="Sheet formulas overview">
              <strong>Formulas mode</strong>
              {Object.entries(activeSheet.cells)
                .filter(([, cell]) => cell.kind === 'formula')
                .slice(0, 24)
                .map(([key, cell]) => (
                  <FormulaChip
                    key={key}
                    type="button"
                    onClick={() => jumpToCell(activeSheet.id, key)}
                    title={`Jump to ${key}`}
                  >
                    <span>{key}:</span> {cell.value}
                  </FormulaChip>
                ))}
              {Object.values(activeSheet.cells).filter(c => c.kind === 'formula')
                .length === 0 && (
                <span>No formulas on {activeSheet.name}.</span>
              )}
              <ImportWarningsDismiss
                type="button"
                onClick={() => setShowFormulas(false)}
              >
                Done
              </ImportWarningsDismiss>
            </FormulasStrip>
          )}

          {agentActivityOpen && (
            <AgentActivityBar
              hiddenInRoom={compact}
              entries={document.workbook.agentLog ?? []}
              onClose={() => setAgentActivityOpen(false)}
            />
          )}

          {commentBarOpen && (
            <CommentBar
              hiddenInRoom={compact}
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
                {...workbookAgentRefs(filePath, document.metadata.title, activeSheet)}
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
                    workbook={document.workbook}
                    visible={showCompatibilityGrid}
                    gridTemplate={gridTemplate}
                    visibleRows={visibleRows}
                    selectedCell={selectedCell}
                    normalizedRange={normalizedRange}
                    showFormulas={showFormulas}
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
                hiddenInRoom={compact}
                charts={activeSheet.charts ?? []}
                onDeleteChart={deleteChart}
                onUpdateChart={updateChartConfig}
              />
            </SheetColumn>

            {inspectorOpen && !compact && (
              <InspectorRail
                tab={inspectorTab}
                onSelectTab={setInspectorTab}
                onClose={() => setInspectorOpen(false)}
                document={document}
                activeSheet={activeSheet}
                selectedCell={selectedCell}
                selectedRange={normalizedRange ?? `${selectedCell}:${selectedCell}`}
                invalidFlags={invalidFlagsBySheet[activeSheet.id] ?? []}
                showFormulas={showFormulas}
                onToggleShowFormulas={() => setShowFormulas(v => !v)}
                onJumpToCell={jumpToCell}
                onAddConditionalRule={addConditionalRuleFull}
                onDeleteConditionalRule={deleteConditionalRuleById}
                onMoveConditionalRule={moveConditionalRuleById}
                onAddValidationRule={addValidationRuleFull}
                onDeleteValidationRule={deleteValidationRuleById}
                onSaveNamedRange={saveNamedRange}
                onDeleteNamedRange={removeNamedRangeById}
                onJumpToNamedRange={jumpToNamedRange}
                onCreatePivotFromSelection={createPivotFromSelection}
                onUpdatePivotDefinition={updatePivotDefinition}
                onRefreshPivot={refreshActivePivot}
                onRevertAgentEntry={revertAgentEntry}
              />
            )}
          </MainArea>

          <SheetTabsBar
            compact={compact}
            sheets={visibleSheets}
            hiddenSheets={hiddenSheets}
            activeSheetId={activeSheet.id}
            renamingSheetId={renamingSheetId}
            renameDraft={renameDraft}
            statusText={statusText}
            statPills={statPills}
            invalidCount={(invalidFlagsBySheet[activeSheet.id] ?? []).length}
            zoom={zoom}
            onCopyStat={pill => {
              void writeClipboard(pill.raw).catch(() => undefined)
              showToast(`Copied ${pill.label}: ${pill.raw}`)
            }}
            onOpenInvalidAudit={() => {
              setInspectorTab('audit')
              setInspectorOpen(true)
            }}
            onZoomChange={changeZoom}
            onAddSheet={addSheet}
            onSelectSheet={selectSheet}
            onRenameSheet={renameCurrentSheet}
            onBeginRenameSheet={beginRenameSheet}
            onDuplicateSheet={duplicateCurrentSheet}
            onDeleteSheet={deleteCurrentSheet}
            onRenameDraftChange={setRenameDraft}
            onCommitSheetRename={commitSheetRename}
            onCancelSheetRename={cancelSheetRename}
            onMoveSheet={moveSheet}
            onSetSheetTabColor={changeSheetTabColor}
            onHideSheet={hideSheetById}
            onUnhideSheet={unhideSheetById}
          />
          {!compact && (
            <div
              data-testid="sheets-live-data-bar"
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                padding: '4px 10px',
                borderTop: '1px solid var(--n-line, #d4d7dc)',
                background: 'var(--n-panel, #f6f7f8)',
                fontSize: '11.5px',
                flexWrap: 'wrap',
              }}
            >
              <strong style={{ fontSize: '10.5px', textTransform: 'uppercase', letterSpacing: '0.05em', color: 'var(--n-soft, #474c54)' }}>
                Live Data &amp; Transform:
              </strong>
              <button
                type="button"
                onClick={() => {
                  void publishSheetLiveData(document, activeSheet, bridge)
                    .then(ds =>
                      showToast(`Published live dataset "${ds.name}" (${ds.metrics.length} KPIs, ${ds.table.rows.length} rows).`),
                    )
                    .catch(reportFailure)
                }}
                style={{ padding: '2px 8px', borderRadius: '3px', border: '1px solid var(--n-line, #c8ccd2)', background: 'var(--n-sheet, #fff)', cursor: 'pointer', fontSize: '11px', fontWeight: 600 }}
              >
                Publish live dataset
              </button>
              <button
                type="button"
                onClick={() => {
                  void sendSheetToApp(document, activeSheet, 'pureslides', bridge)
                    .then(() =>
                      showToast(`Sent "${activeSheet.name}" live metrics & table to PureSlides.`),
                    )
                    .catch(reportFailure)
                }}
                style={{ padding: '2px 8px', borderRadius: '3px', border: '1px solid var(--n-line, #c8ccd2)', background: 'var(--n-sheet, #fff)', cursor: 'pointer', fontSize: '11px' }}
              >
                Send to PureSlides
              </button>
              <button
                type="button"
                onClick={() => {
                  void sendSheetToApp(document, activeSheet, 'canvas2', bridge)
                    .then(() =>
                      showToast(`Sent "${activeSheet.name}" live metrics & table to PureCanvas.`),
                    )
                    .catch(reportFailure)
                }}
                style={{ padding: '2px 8px', borderRadius: '3px', border: '1px solid var(--n-line, #c8ccd2)', background: 'var(--n-sheet, #fff)', cursor: 'pointer', fontSize: '11px' }}
              >
                Send to PureCanvas
              </button>
              <button
                type="button"
                onClick={() => {
                  void sendSheetToApp(document, activeSheet, 'puresite', bridge)
                    .then(() =>
                      showToast(`Sent "${activeSheet.name}" live metrics & table to PureSite.`),
                    )
                    .catch(reportFailure)
                }}
                style={{ padding: '2px 8px', borderRadius: '3px', border: '1px solid var(--n-line, #c8ccd2)', background: 'var(--n-sheet, #fff)', cursor: 'pointer', fontSize: '11px' }}
              >
                Send to PureSite
              </button>
            </div>
          )}
          {!compact && <WorkspaceThreadStrip app="puresheets" />}
        </SheetSurface>

        <CommandPalette
          open={paletteOpen}
          commands={paletteCommands}
          onClose={() => setPaletteOpen(false)}
        />
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
// Rendered only beside the Executive room, so with compact off none of this
// exists. Hiding with CSS keeps the header, command bar and panels mounted, with
// their state and tools alive, for when Expand ends the split. AppFrame's own
// header is its only direct <header> child; the command bar's is nested.
const COMPACT_ROOM_CSS = `
  [data-compact-room] > header,
  [data-compact-room] [data-compact-hidden] {
    display: none !important;
  }
`

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
  display: flex;
  flex-direction: column;
  min-height: 0;
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

// The document title and grouped commands share a full-width, quiet header.
/* The title, the tool tabs and the quick actions share one line; the tools for the open tab run full width under it. */
const TopBar = styled.header`
  display: grid;
  grid-template-columns: auto minmax(0, 1fr);
  align-items: center;
  flex-shrink: 0;
  gap: 10px 20px;
  padding: 14px var(--pure-chrome-inset) 12px;
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-bar);
`

/* The title is as wide as its words, closed by a hairline, so the tabs start right after it rather than floating mid-bar. */
const TitleBlock = styled.div`
  min-width: 0;
  padding-right: 20px;
  border-right: 1px solid var(--pure-chrome-line);
`

const FileNameInput = styled.input`
  field-sizing: content;
  width: auto;
  min-width: 6ch;
  max-width: 32ch;
  overflow: hidden;
  margin: 0;
  padding: 0;
  border: 0;
  border-bottom: 1px solid transparent;
  background: transparent;
  color: var(--sheets-ink);
  font: inherit;
  font-size: 16px;
  font-weight: 600;
  text-overflow: ellipsis;
  white-space: nowrap;

  &:hover,
  &:focus {
    border-bottom-color: var(--sheets-line);
    outline: 0;
  }
`

const SheetSurface = styled.div`
  flex: 1;
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

const PivotBanner = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 6px var(--pure-chrome-inset);
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--n-panel, var(--pure-chrome-well));
  color: var(--sheets-ink);
  font-size: 12px;
`

const PivotBannerActions = styled.div`
  display: inline-flex;
  gap: 6px;
`

const FormulasStrip = styled.div`
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 6px;
  padding: 6px var(--pure-chrome-inset);
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--n-panel, var(--pure-chrome-well));
  color: var(--sheets-ink);
  font-size: 11.5px;
`

const FormulaChip = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 2px 6px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-family: var(--sheets-mono);
  font-size: 11px;
  cursor: pointer;
  span {
    font-weight: 700;
    color: var(--sheets-accent);
  }
  &:hover {
    background: var(--pure-chrome-hover);
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
