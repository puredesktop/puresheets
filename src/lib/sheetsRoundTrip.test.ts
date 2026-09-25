/**
 * Phase S0 HARD GATE (MAIL_SHEETS_MASTER_PLAN.md, Track S):
 * existing `.sheets` and `.sheets.html` files must open, edit, save, and
 * reopen through the Univer snapshot bridge WITHOUT data loss.
 *
 * The bridge path under test mirrors the app exactly:
 *   fixture file → loadSheetsDocument (workbook model + fresh snapshot)
 *   → snapshot edited the way a Univer edit exports it (workbook.save())
 *   → documentFromUniverSnapshot (real headless Univer core validates it)
 *   → serializeSheetsDocument / workbookFilesForSave → reload → compare.
 */
import { describe, expect, it } from 'vitest'
import { CellValueType } from '@univerjs/core'
import {
  loadSheetsDocument,
  loadSheetsHtmlDocument,
  serializeSheetsDocument,
  serializeSheetsHtmlDocument,
} from './sheetsDocument'
import { workbookFilesForSave } from './workbookFiles'
import { documentFromUniverSnapshot } from './univerAdapter'
import type { PureSheetsUniverSnapshot } from './univerAdapter'
import {
  UNIVER_SURFACE_COLUMN_COUNT,
  UNIVER_SURFACE_ROW_COUNT,
} from '../constants'
import {
  cloneWithUpdatedSheet,
  deleteColumn,
  insertColumn,
  insertRow,
} from './workbookModel'
import type { PureSheetsDocument, WorkbookSheet } from '../types'

/**
 * Realistic legacy `.sheets` fixture: two user sheets plus an agent draft,
 * formulas, styles, number formats, sizing, freeze/sort/filter state, a
 * chart, and a full agent-QA record. Written as literal JSON (no engine
 * snapshot) the way pre-Univer PureSheets saved files.
 */
const LEGACY_SHEETS_FIXTURE = JSON.stringify({
  app: 'PureSheets',
  version: 1,
  engine: { name: 'univer' },
  metadata: {
    title: 'Q3 Forecast',
    createdAt: '2026-01-10T09:00:00.000Z',
    updatedAt: '2026-06-30T10:30:00.000Z',
  },
  workbook: {
    activeSheetId: 'data',
    sheets: [
      {
        id: 'data',
        name: 'Data',
        rowCount: 40,
        columnCount: 18,
        cells: {
          A1: {
            value: 'Region',
            kind: 'text',
            style: { bold: true, fillColor: '#eef1eb', border: true },
          },
          B1: {
            value: 'Revenue',
            kind: 'text',
            style: { bold: true, align: 'right' },
          },
          A2: { value: 'North', kind: 'text' },
          B2: {
            value: '1200',
            kind: 'number',
            style: { numberFormat: 'currency', decimals: 0 },
          },
          A3: { value: 'South', kind: 'text' },
          B3: {
            value: '870',
            kind: 'number',
            style: { numberFormat: 'currency' },
          },
          A4: { value: 'West', kind: 'text', style: { italic: true } },
          B4: { value: '450', kind: 'number' },
          B5: { value: '=SUM(B2:B4)', kind: 'formula', style: { bold: true } },
          C2: { value: '=B2*0.1', kind: 'formula' },
          D1: {
            value: 'N/A',
            kind: 'text',
            style: { textColor: '#982d25', fontSize: 10 },
          },
        },
        columnWidths: { '0': 140, '1': 96 },
        rowHeights: { '0': 36 },
        frozenRows: 1,
        filterRow: 0,
        filterColumn: 0,
        filterQuery: 'orth',
        sortState: { range: 'A2:B4', column: 0, direction: 'asc' },
        charts: [
          {
            id: 'chart-1',
            title: 'A2:B4 bar chart',
            type: 'bar',
            range: 'A2:B4',
            labels: ['North', 'South', 'West'],
            values: [1200, 870, 450],
          },
          {
            id: 'chart-2',
            title: 'Revenue share',
            type: 'donut',
            range: 'A2:B4',
            labels: ['North', 'South', 'West'],
            values: [1200, 870, 450],
            legend: 'bottom',
            showAxes: false,
            xLabel: 'Region',
            yLabel: 'Revenue',
          },
        ],
        comments: {
          B2: 'Confirmed with regional lead',
          D1: 'Data source unavailable this quarter',
        },
        validations: [
          {
            id: 'validation-1',
            range: 'B2:B4',
            kind: 'number-range',
            min: 0,
            max: 100000,
          },
          {
            id: 'validation-2',
            range: 'A2:A4',
            kind: 'list',
            options: ['North', 'South', 'West'],
          },
        ],
        protection: { lockedRanges: ['A1:D1'] },
        conditionalFormats: [
          {
            id: 'conditional-format-1',
            range: 'B2:B4',
            condition: { kind: 'greater', value: '1000' },
            style: { fillColor: '#e5f3df' },
          },
          {
            id: 'conditional-format-2',
            range: 'A2:A4',
            condition: { kind: 'text-contains', value: 'south' },
            style: { textColor: '#982d25', bold: true },
          },
        ],
      },
      {
        id: 'notes',
        name: 'Notes',
        rowCount: 40,
        columnCount: 18,
        cells: {
          A1: { value: 'Assumptions reviewed 2026-06-28', kind: 'text' },
        },
        frozenColumns: 1,
      },
      {
        id: 'draft-1',
        name: 'Data - Agent draft',
        rowCount: 40,
        columnCount: 18,
        cells: {
          A2: {
            value: 'NORTH',
            kind: 'text',
            style: { fillColor: '#fff4bf', border: true },
          },
        },
        agentDraft: {
          sourceSheetId: 'data',
          sourceSheetName: 'Data',
          agentRunId: 'agent-run-1',
          agentName: 'PureSheets Assistant',
          createdAt: '2026-06-30T10:00:00.000Z',
          selectedSourceRange: 'A2:A4',
          prompt: 'uppercase region names',
          qaStatus: 'pending',
        },
      },
    ],
    agentQa: {
      operations: [
        {
          id: 'op-1',
          draftSheetId: 'draft-1',
          sourceSheetId: 'data',
          agentName: 'PureSheets Assistant',
          prompt: 'uppercase region names',
          affectedSourceRange: 'A2:A4',
          affectedDraftRange: 'A2:A4',
          changeIds: ['change-1'],
          explanation: 'PureSheets Assistant proposed 1 change.',
          approvalStatus: 'pending',
          createdAt: '2026-06-30T10:00:00.000Z',
          updatedAt: '2026-06-30T10:00:00.000Z',
        },
      ],
      changes: [
        {
          id: 'change-1',
          operationId: 'op-1',
          draftSheetId: 'draft-1',
          sourceSheetId: 'data',
          sourceCell: 'A2',
          draftCell: 'A2',
          original: { value: 'North', kind: 'text' },
          proposed: { value: 'NORTH', kind: 'text' },
          explanation: 'Proposed from prompt "uppercase region names" for A2.',
          status: 'pending',
          createdAt: '2026-06-30T10:00:00.000Z',
        },
      ],
      history: [
        {
          id: 'history-1',
          operationId: 'op-1',
          draftSheetId: 'draft-1',
          sourceSheetId: 'data',
          timestamp: '2026-06-30T10:00:00.000Z',
          status: 'proposed',
          text: 'PureSheets Assistant proposed 1 change to Data.',
          affectedRanges: ['A2:A4'],
        },
      ],
    },
  },
})

function sheetById(document: PureSheetsDocument, id: string): WorkbookSheet {
  const sheet = document.workbook.sheets.find(item => item.id === id)
  if (!sheet) throw new Error(`expected sheet ${id}`)
  return sheet
}

/** Emulate one interactive Univer edit: the editor mutates its workbook and
 * exports the whole snapshot via `workbook.save()`. */
function withEditedCell(
  document: PureSheetsDocument,
  sheetId: string,
  row: number,
  column: number,
  value: string,
): PureSheetsUniverSnapshot {
  const snapshot = structuredClone(
    document.engine.snapshot,
  ) as PureSheetsUniverSnapshot
  const sheet = snapshot.sheets[sheetId]
  if (!sheet) throw new Error(`snapshot is missing sheet ${sheetId}`)
  sheet.cellData ??= {}
  sheet.cellData[row] = {
    ...sheet.cellData[row],
    [column]: { v: value, t: CellValueType.STRING },
  }
  return snapshot
}

function expectNoDataLoss(
  reloaded: PureSheetsDocument,
  original: PureSheetsDocument,
  editedCellKey: string,
  editedValue: string,
): void {
  // Sheet inventory: ids, names, order, dimensions.
  expect(reloaded.workbook.sheets.map(sheet => sheet.id)).toEqual(
    original.workbook.sheets.map(sheet => sheet.id),
  )
  expect(reloaded.workbook.sheets.map(sheet => sheet.name)).toEqual(
    original.workbook.sheets.map(sheet => sheet.name),
  )

  for (const originalSheet of original.workbook.sheets) {
    const reloadedSheet = sheetById(reloaded, originalSheet.id)
    // The rendered canvas floor grew to a full spreadsheet grid (#272), so a
    // sheet smaller than the floor round-trips up to it — no data is lost,
    // the grid is just at least surface-sized.
    expect(reloadedSheet.rowCount).toBe(
      Math.max(originalSheet.rowCount, UNIVER_SURFACE_ROW_COUNT),
    )
    expect(reloadedSheet.columnCount).toBe(
      Math.max(originalSheet.columnCount, UNIVER_SURFACE_COLUMN_COUNT),
    )

    // Every original cell survives: value, kind, and style.
    for (const [key, cell] of Object.entries(originalSheet.cells)) {
      if (originalSheet.id === 'data' && key === editedCellKey) continue
      const reloadedCell = reloadedSheet.cells[key]
      expect(reloadedCell, `${originalSheet.id}!${key}`).toBeTruthy()
      expect(reloadedCell.value, `${originalSheet.id}!${key} value`).toBe(
        cell.value,
      )
      expect(reloadedCell.kind, `${originalSheet.id}!${key} kind`).toBe(
        cell.kind,
      )
      expect(
        reloadedCell.style ?? {},
        `${originalSheet.id}!${key} style`,
      ).toEqual(cell.style ?? {})
    }

    // Sizing, freeze, and filter state.
    expect(reloadedSheet.columnWidths ?? {}).toEqual(
      originalSheet.columnWidths ?? {},
    )
    expect(reloadedSheet.rowHeights ?? {}).toEqual(
      originalSheet.rowHeights ?? {},
    )
    expect(reloadedSheet.frozenRows ?? 0).toBe(originalSheet.frozenRows ?? 0)
    expect(reloadedSheet.frozenColumns ?? 0).toBe(
      originalSheet.frozenColumns ?? 0,
    )
    expect(reloadedSheet.filterRow).toBe(originalSheet.filterRow)
    expect(reloadedSheet.filterColumn).toBe(originalSheet.filterColumn)
    expect(reloadedSheet.filterQuery).toBe(originalSheet.filterQuery)

    // Phase S3 metadata: comments, validation, protection, conditional
    // formats all survive the bridge in full.
    expect(reloadedSheet.comments).toEqual(originalSheet.comments)
    expect(reloadedSheet.validations).toEqual(originalSheet.validations)
    expect(reloadedSheet.protection).toEqual(originalSheet.protection)
    expect(reloadedSheet.conditionalFormats).toEqual(
      originalSheet.conditionalFormats,
    )

    // Charts and agent-draft metadata.
    expect(reloadedSheet.charts).toEqual(originalSheet.charts)
    expect(reloadedSheet.agentDraft).toEqual(originalSheet.agentDraft)
  }

  // The edit landed.
  expect(sheetById(reloaded, 'data').cells[editedCellKey]?.value).toBe(
    editedValue,
  )

  // Metadata and format invariants.
  expect(reloaded.app).toBe('PureSheets')
  expect(reloaded.version).toBe(1)
  expect(reloaded.metadata.title).toBe(original.metadata.title)
  expect(reloaded.metadata.createdAt).toBe(original.metadata.createdAt)
}

describe('Phase S0 hard gate: .sheets round-trip through the Univer bridge', () => {
  it('opens a legacy .sheets file, edits via the snapshot bridge, saves, and reloads without data loss', () => {
    const original = loadSheetsDocument(LEGACY_SHEETS_FIXTURE)

    // Edit D2 (empty in the fixture) the way the Univer surface exports it.
    const editedSnapshot = withEditedCell(original, 'data', 1, 3, 'Edited')
    const edited = documentFromUniverSnapshot(editedSnapshot, original)

    const reloaded = loadSheetsDocument(serializeSheetsDocument(edited))
    expectNoDataLoss(reloaded, original, 'D2', 'Edited')

    // The saved file keeps a fresh embedded Univer snapshot (the bridge).
    const saved = JSON.parse(serializeSheetsDocument(edited)) as {
      engine: { snapshot?: { sheets?: Record<string, unknown> } }
    }
    expect(saved.engine.snapshot?.sheets).toBeTruthy()
  })

  it('survives a second full bridge round-trip (open → save → open → save → open)', () => {
    const original = loadSheetsDocument(LEGACY_SHEETS_FIXTURE)
    const editedSnapshot = withEditedCell(original, 'data', 1, 3, 'Edited')
    const edited = documentFromUniverSnapshot(editedSnapshot, original)
    const once = loadSheetsDocument(serializeSheetsDocument(edited))
    const twice = loadSheetsDocument(serializeSheetsDocument(once))
    expectNoDataLoss(twice, original, 'D2', 'Edited')
  })

  it('round-trips .sheets.html files through the same bridge path', () => {
    const original = loadSheetsHtmlDocument(
      serializeSheetsHtmlDocument(loadSheetsDocument(LEGACY_SHEETS_FIXTURE)),
    )
    const editedSnapshot = withEditedCell(original, 'data', 1, 3, 'Edited')
    const edited = documentFromUniverSnapshot(editedSnapshot, original)
    const reloaded = loadSheetsHtmlDocument(serializeSheetsHtmlDocument(edited))
    expectNoDataLoss(reloaded, original, 'D2', 'Edited')
  })

  it('serializes package saves (manifest + workbook.json) that reload identically', () => {
    const original = loadSheetsDocument(LEGACY_SHEETS_FIXTURE)
    const editedSnapshot = withEditedCell(original, 'data', 1, 3, 'Edited')
    const edited = documentFromUniverSnapshot(editedSnapshot, original)
    const files = workbookFilesForSave(edited, {
      boundPath: '/workspace/Q3 Forecast.sheets',
      isPackage: true,
    })
    expect(files.map(file => file.name)).toEqual([
      'manifest.json',
      'workbook.json',
    ])
    const manifest = JSON.parse(files[0].content) as Record<string, unknown>
    expect(manifest.kind).toBe('purescience.sheets.document')
    expect(manifest.title).toBe('Q3 Forecast')
    const reloaded = loadSheetsDocument(files[1].content)
    expectNoDataLoss(reloaded, original, 'D2', 'Edited')
  })

  it('persists row/column structure edits (insert + delete) through the bridge', () => {
    const original = loadSheetsDocument(LEGACY_SHEETS_FIXTURE)

    // Insert a row above row 2 (index 1) on the data sheet — same model
    // transform the shell's insert-row fallback uses; the Univer path
    // produces the equivalent snapshot via worksheet.insertRows.
    const inserted = cloneWithUpdatedSheet(
      original,
      insertRow(sheetById(original, 'data'), 1),
    )
    const reloadedInsert = loadSheetsDocument(serializeSheetsDocument(inserted))
    const dataAfterInsert = sheetById(reloadedInsert, 'data')
    // The small fixture stays under the rendered-canvas floor, so its row
    // count reports the floor rather than the data extent (#272); the row
    // shifts below are what the insert actually has to preserve.
    expect(dataAfterInsert.rowCount).toBe(UNIVER_SURFACE_ROW_COUNT)
    // Rows at and below the insertion point shifted down…
    expect(dataAfterInsert.cells.A2).toBeUndefined()
    expect(dataAfterInsert.cells.A3?.value).toBe('North')
    expect(dataAfterInsert.cells.B3?.value).toBe('1200')
    expect(dataAfterInsert.cells.B3?.style).toEqual({
      numberFormat: 'currency',
      decimals: 0,
    })
    // …and formula references followed the shift, surviving save + reload.
    expect(dataAfterInsert.cells.B6?.value).toBe('=SUM(B3:B5)')
    expect(dataAfterInsert.cells.C3?.value).toBe('=B3*0.1')
    // Header row (above the insertion point) unmoved, sizing intact.
    expect(dataAfterInsert.cells.A1?.value).toBe('Region')
    expect(dataAfterInsert.rowHeights).toEqual({ '0': 36 })
    expect(dataAfterInsert.frozenRows).toBe(1)
    // Positional metadata (comments, validation/protection/conditional
    // ranges) shifted with the insert and persisted through the bridge.
    expect(dataAfterInsert.comments).toEqual({
      B3: 'Confirmed with regional lead',
      D1: 'Data source unavailable this quarter',
    })
    expect(dataAfterInsert.validations?.map(rule => rule.range)).toEqual([
      'B3:B5',
      'A3:A5',
    ])
    expect(dataAfterInsert.protection).toEqual({ lockedRanges: ['A1:D1'] })
    expect(
      dataAfterInsert.conditionalFormats?.map(rule => rule.range),
    ).toEqual(['B3:B5', 'A3:A5'])

    // Insert a column before B (index 1): cells, widths, and references
    // shift right and survive save + reload. Inserting into the floor-sized
    // grid grows it by one column past the floor, and that larger count
    // persists through the round-trip (#272); the cell/reference shifts below
    // are the real subject of this check.
    const widened = cloneWithUpdatedSheet(
      reloadedInsert,
      insertColumn(dataAfterInsert, 1),
    )
    const reloadedWiden = loadSheetsDocument(serializeSheetsDocument(widened))
    const dataAfterWiden = sheetById(reloadedWiden, 'data')
    expect(dataAfterWiden.columnCount).toBe(UNIVER_SURFACE_COLUMN_COUNT + 1)
    expect(dataAfterWiden.cells.B3).toBeUndefined()
    expect(dataAfterWiden.cells.C3?.value).toBe('1200')
    expect(dataAfterWiden.cells.C6?.value).toBe('=SUM(C3:C5)')
    expect(dataAfterWiden.cells.D3?.value).toBe('=C3*0.1')
    expect(dataAfterWiden.cells.E1?.value).toBe('N/A')
    expect(dataAfterWiden.comments).toEqual({
      C3: 'Confirmed with regional lead',
      E1: 'Data source unavailable this quarter',
    })
    expect(dataAfterWiden.protection).toEqual({ lockedRanges: ['A1:E1'] })
    expect(dataAfterWiden.validations?.map(rule => rule.range)).toEqual([
      'C3:C5',
      'A3:A5',
    ])

    // Delete the inserted column again and round-trip: references come back,
    // nothing else moved. (Counts never shrink below the rendered-canvas
    // floor — that floor is rendering minimum, not data.)
    const deleted = cloneWithUpdatedSheet(
      reloadedWiden,
      deleteColumn(dataAfterWiden, 1),
    )
    const reloadedDelete = loadSheetsDocument(serializeSheetsDocument(deleted))
    const dataAfterDelete = sheetById(reloadedDelete, 'data')
    expect(dataAfterDelete.columnCount).toBe(UNIVER_SURFACE_COLUMN_COUNT)
    expect(dataAfterDelete.cells.B3?.value).toBe('1200')
    expect(dataAfterDelete.cells.B6?.value).toBe('=SUM(B3:B5)')
    expect(dataAfterDelete.cells.C3?.value).toBe('=B3*0.1')
    expect(dataAfterDelete.cells.D1?.value).toBe('N/A')
    expect(dataAfterDelete.cells.D1?.style).toEqual({
      textColor: '#982d25',
      fontSize: 10,
    })
    // Untouched cells survive all three structural round-trips.
    expect(dataAfterDelete.cells.A3?.value).toBe('North')
    expect(dataAfterDelete.cells.A1?.value).toBe('Region')
    // Positional metadata returned to its pre-widen positions.
    expect(dataAfterDelete.comments).toEqual({
      B3: 'Confirmed with regional lead',
      D1: 'Data source unavailable this quarter',
    })
    expect(dataAfterDelete.protection).toEqual({ lockedRanges: ['A1:D1'] })
    expect(dataAfterDelete.validations?.map(rule => rule.range)).toEqual([
      'B3:B5',
      'A3:A5',
    ])
  })

  it('serializes legacy flat saves in the original single-file format', () => {
    const original = loadSheetsDocument(LEGACY_SHEETS_FIXTURE)
    const htmlFiles = workbookFilesForSave(original, {
      boundPath: '/workspace/legacy.sheets.html',
      isPackage: false,
    })
    expect(htmlFiles).toHaveLength(1)
    expect(htmlFiles[0].name).toBeNull()
    expect(htmlFiles[0].content).toContain('<!doctype html>')
    const flatFiles = workbookFilesForSave(original, {
      boundPath: '/workspace/legacy.sheets',
      isPackage: false,
    })
    expect(flatFiles[0].content.trimStart().startsWith('{')).toBe(true)
  })
})
