import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import { addValidationRule } from '../lib/sheetDataFeatures'
import {
  addSheet,
  applyAgentCellChangesDirect,
  createBlankWorkbook,
  lockRange,
} from '../lib/workbookModel'
import type { PureSheetsDocument } from '../types'
import { PURESHEETS_ATOMIC_TOOL_SCHEMAS } from './atomicToolSchemas'
import {
  AgentSheetsToolError,
  EXPOSED_TOOL_NAMES,
  PURESHEETS_AGENT_TOOL_NAMES,
  PURESHEETS_ALL_ATOMIC_TOOL_NAMES,
  PURESHEETS_CORE_TOOL_NAMES,
  PURESHEETS_EXTENDED_TOOL_NAMES,
  PURESHEETS_FUNCTION_NAMES,
  PURESHEETS_META_TOOL_NAMES,
  type SheetsAgentToolContext,
} from './catalog'
import {
  PURESHEETS_OPERATIONS_REGISTRY,
  buildPureSheetsHandlers,
} from './operationsRegistry'

const invoke = (shortName: string, args: Record<string, unknown> = {}) => ({
  toolCallId: `tc_${shortName}`,
  shortName,
  arguments: args,
})

const RELEASE_VERBS = new Set([
  'send',
  'sends',
  'sending',
  'merge',
  'merges',
  'merging',
  'publish',
  'publishes',
  'publishing',
  'deposit',
  'deposits',
  'depositing',
  'deploy',
  'deploys',
  'deploying',
  'post',
  'posts',
  'posting',
])

function makeHarness(
  initial: PureSheetsDocument = createBlankWorkbook('Test Workbook'),
  overrides: Partial<SheetsAgentToolContext> = {},
) {
  let doc: PureSheetsDocument | null = initial
  let setCount = 0
  const ctx: SheetsAgentToolContext = {
    get document() {
      return doc
    },
    setDocument(updater) {
      setCount += 1
      if (doc) doc = updater(doc)
    },
    filePath: '/workspace/budget.sheets',
    currentFilePath: () => '/workspace/budget.sheets',
    selectedRange: 'A1:B2',
    openWorkbook: vi.fn(async path => ({
      title: 'Imported',
      sourcePath: path,
      filePath: '/workspace/imported.sheets',
      warnings: [],
    })),
    createWorkbook: vi.fn(async title => ({ title, filePath: null })),
    saveAsXlsx: vi.fn(async (path, _overwrite) => ({
      path,
      sheets: ['Sheet 1'],
      warnings: [],
    })),
    autoFitRows: vi.fn(async () => true),
    completeWorkbook: vi.fn(async () => ({
      taskOutcome: 'saved',
      artifactPaths: ['/workspace/budget.sheets'],
      observations: ['verified'],
    })),
    ...overrides,
  }
  return {
    ctx,
    getDoc: () => doc!,
    setDoc: (next: PureSheetsDocument) => {
      doc = next
    },
    get setCount() {
      return setCount
    },
  }
}

function makeRichWorkbook(): PureSheetsDocument {
  let doc = createBlankWorkbook('Q1 Model')
  const dataSheet = doc.workbook.sheets[0]
  dataSheet.name = 'Sales'
  doc = applyAgentCellChangesDirect(
    doc,
    dataSheet.id,
    'A1:D6',
    'Seed sales data',
    [
      { cell: 'A1', value: 'Region', kind: 'text' },
      { cell: 'B1', value: 'Units', kind: 'text' },
      { cell: 'C1', value: 'Price', kind: 'text' },
      { cell: 'D1', value: 'Revenue', kind: 'text' },
      { cell: 'A2', value: 'North', kind: 'text' },
      { cell: 'B2', value: '10', kind: 'number' },
      { cell: 'C2', value: '25', kind: 'number' },
      { cell: 'D2', value: '=B2*C2', kind: 'formula' },
      { cell: 'A3', value: 'South', kind: 'text' },
      { cell: 'B3', value: '20', kind: 'number' },
      { cell: 'C3', value: '30', kind: 'number' },
      { cell: 'D3', value: '=B3*C3', kind: 'formula' },
      { cell: 'A4', value: 'North', kind: 'text' },
      { cell: 'B4', value: '5', kind: 'number' },
      { cell: 'C4', value: '25', kind: 'number' },
      { cell: 'D4', value: '=B4*C4', kind: 'formula' },
      { cell: 'A5', value: 'West', kind: 'text' },
      { cell: 'B5', value: '15', kind: 'number' },
      { cell: 'C5', value: '40', kind: 'number' },
      { cell: 'D5', value: '=B5*C5', kind: 'formula' },
    ],
  ).document

  doc = addSheet(doc, 'Summary')
  const summarySheet = doc.workbook.sheets[1]
  doc = applyAgentCellChangesDirect(
    doc,
    summarySheet.id,
    'A1:B3',
    'Seed summary formulas',
    [
      { cell: 'A1', value: 'Metric', kind: 'text' },
      { cell: 'B1', value: 'Value', kind: 'text' },
      { cell: 'A2', value: 'Total Revenue', kind: 'text' },
      { cell: 'B2', value: '=SUM(Sales!D2:D5)', kind: 'formula' },
      { cell: 'A3', value: 'Total Units', kind: 'text' },
      { cell: 'B3', value: '=SUM(Sales!B2:B5)', kind: 'formula' },
    ],
  ).document

  doc.workbook.namedRanges = [
    { id: 'nr-1', name: 'SalesRevenue', sheetId: dataSheet.id, range: 'D2:D5' },
  ]
  return doc
}

describe('PureSheets Operations Registry invariants', () => {
  it('exposes the exact 3-tier layout and syncs 1:1 with plugin.json and catalog.ts', () => {
    const manifest = JSON.parse(readFileSync('plugin.json', 'utf8')) as {
      app: {
        agents: {
          tools: Array<{
            name: string
            description: string
            risk: string
            requiresApproval: boolean
            inputSchema: Record<string, unknown>
          }>
        }
      }
    }

    const summary = PURESHEETS_OPERATIONS_REGISTRY.auditSummary()
    expect(summary).toEqual({
      appSlug: 'sheets',
      functions: 5,
      coreTools: 4,
      extendedTools: 13,
      metaTools: 3,
      totalExposed: 12,
      totalAtomicTools: 17,
      totalRegistered: 22,
      schemaReductionPercent: 29,
    })

    expect(PURESHEETS_OPERATIONS_REGISTRY.exposedToolNames).toEqual([
      ...EXPOSED_TOOL_NAMES,
    ])
    expect([...PURESHEETS_AGENT_TOOL_NAMES]).toEqual([...EXPOSED_TOOL_NAMES])
    expect(
      PURESHEETS_OPERATIONS_REGISTRY.functions.map(f => f.name),
    ).toEqual([...PURESHEETS_FUNCTION_NAMES])
    expect(
      PURESHEETS_OPERATIONS_REGISTRY.coreTools.map(t => t.name),
    ).toEqual([...PURESHEETS_CORE_TOOL_NAMES])
    expect(PURESHEETS_OPERATIONS_REGISTRY.extendedToolNames).toEqual([
      ...PURESHEETS_EXTENDED_TOOL_NAMES,
    ])
    expect([...PURESHEETS_ALL_ATOMIC_TOOL_NAMES].sort()).toEqual(
      PURESHEETS_ATOMIC_TOOL_SCHEMAS.map(t => t.name).sort(),
    )
    expect(manifest.app.agents.tools).toEqual(
      PURESHEETS_OPERATIONS_REGISTRY.manifestTools,
    )
    expect(PURESHEETS_META_TOOL_NAMES).toEqual([
      'discoverAppTools',
      'callExtendedReadTool',
      'callExtendedTool',
    ])
  })

  it('satisfies manifest rules: descriptions >= 80 chars, explicit booleans, no non-release release verbs', () => {
    for (const tool of PURESHEETS_OPERATIONS_REGISTRY.manifestTools) {
      const normalized = tool.description.trim().replace(/\s+/g, ' ')
      expect(
        normalized.length,
        `${tool.name} description too short (${normalized.length})`,
      ).toBeGreaterThanOrEqual(80)
      expect(typeof tool.requiresApproval).toBe('boolean')
      const nameFirst =
        tool.name
          .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
          .split(/[^A-Za-z]+/)
          .find(Boolean)
          ?.toLowerCase() ?? ''
      const descFirst =
        normalized
          .split(/[^A-Za-z]+/)
          .find(Boolean)
          ?.toLowerCase() ?? ''
      if (tool.risk !== 'release') {
        expect(
          RELEASE_VERBS.has(nameFirst),
          `${tool.name} starts with release verb`,
        ).toBe(false)
        expect(
          RELEASE_VERBS.has(descFirst),
          `${tool.name} description starts with release verb`,
        ).toBe(false)
      }
    }
  })

  it('builds one handler per exposed tool and none for extended tools directly', () => {
    const { ctx } = makeHarness()
    const handlers = buildPureSheetsHandlers(() => ctx)
    expect(Object.keys(handlers).sort()).toEqual([...EXPOSED_TOOL_NAMES].sort())
    for (const ext of PURESHEETS_EXTENDED_TOOL_NAMES) {
      expect(handlers[ext]).toBeUndefined()
    }
  })
})

describe('Universal escape-hatch meta-tools & guards', () => {
  it('discoverAppTools searches extended tools by query, category, exact names, and includeFunctions', async () => {
    const { ctx } = makeHarness(makeRichWorkbook())
    const handlers = buildPureSheetsHandlers(() => ctx)

    const byQuery = JSON.parse(
      (
        await handlers.discoverAppTools(
          invoke('discoverAppTools', { query: 'excel xlsx export' }),
        )
      ).content,
    )
    expect(byQuery.appSlug).toBe('sheets')
    expect(
      byQuery.operations.some(
        (o: { name: string }) => o.name === 'saveWorkbookAsXlsx',
      ),
    ).toBe(true)

    const byCategory = JSON.parse(
      (
        await handlers.discoverAppTools(
          invoke('discoverAppTools', { category: 'sheets' }),
        )
      ).content,
    )
    expect(
      byCategory.operations.map((o: { name: string }) => o.name).sort(),
    ).toEqual(['addSheet', 'deleteSheet', 'renameSheet'])

    const byNames = JSON.parse(
      (
        await handlers.discoverAppTools(
          invoke('discoverAppTools', {
            names: ['deleteSheet', 'findInWorkbook', 'doesNotExist'],
          }),
        )
      ).content,
    )
    expect(byNames.mode).toBe('describe')
    expect(byNames.operations[0]).toMatchObject({
      name: 'deleteSheet',
      found: true,
      exposure: 'extended',
      invokeVia: 'callExtendedTool',
      requiresConfirmDestructive: true,
    })
    expect(byNames.operations[1]).toMatchObject({
      name: 'findInWorkbook',
      found: true,
      exposure: 'extended',
      invokeVia: 'callExtendedReadTool',
    })
    expect(byNames.operations[2]).toMatchObject({
      name: 'doesNotExist',
      found: false,
    })

    const withFunctions = JSON.parse(
      (
        await handlers.discoverAppTools(
          invoke('discoverAppTools', {
            query: 'audit formula',
            includeFunctions: true,
          }),
        )
      ).content,
    )
    expect(
      withFunctions.operations.some(
        (o: { name: string; invokeVia: string }) =>
          o.name === 'fn_auditWorkbookFormulasAndData' &&
          o.invokeVia === 'direct',
      ),
    ).toBe(true)
  })

  it('callExtendedReadTool routes read-only extended tools and refuses mutating, direct, and unknown tools', async () => {
    const { ctx } = makeHarness(makeRichWorkbook())
    const handlers = buildPureSheetsHandlers(() => ctx)

    const found = JSON.parse(
      (
        await handlers.callExtendedReadTool(
          invoke('callExtendedReadTool', {
            tool: 'findInWorkbook',
            arguments: { query: 'North' },
          }),
        )
      ).content,
    )
    expect(found.matches).toHaveLength(2)

    const log = JSON.parse(
      (
        await handlers.callExtendedReadTool(
          invoke('callExtendedReadTool', {
            tool: 'listSheetChanges',
            arguments: { limit: 5 },
          }),
        )
      ).content,
    )
    expect(log.entries.length).toBeGreaterThanOrEqual(2)

    const opened = JSON.parse(
      (
        await handlers.callExtendedReadTool(
          invoke('callExtendedReadTool', {
            tool: 'openWorkbook',
            arguments: { path: '/workspace/other.xlsx' },
          }),
        )
      ).content,
    )
    expect(opened.sourcePath).toBe('/workspace/other.xlsx')

    await expect(
      handlers.callExtendedReadTool(
        invoke('callExtendedReadTool', {
          tool: 'addSheet',
          arguments: { name: 'Blocked' },
        }),
      ),
    ).rejects.toThrow(/cannot run through callExtendedReadTool/)

    await expect(
      handlers.callExtendedReadTool(
        invoke('callExtendedReadTool', { tool: 'readRange' }),
      ),
    ).rejects.toThrow(/already directly exposed/)

    await expect(
      handlers.callExtendedReadTool(
        invoke('callExtendedReadTool', { tool: 'ghostTool' }),
      ),
    ).rejects.toThrow(/Unknown extended tool/)
  })

  it('callExtendedTool routes mutating extended tools, enforces confirmDestructive on deleteSheet, and refuses read/direct/unknown', async () => {
    const harness = makeHarness(makeRichWorkbook())
    const recordOperation = vi.fn()
    const beforeInvoke = vi.fn()
    const handlers = buildPureSheetsHandlers(() => harness.ctx, {
      beforeInvoke,
      recordOperation,
    })

    const added = JSON.parse(
      (
        await handlers.callExtendedTool(
          invoke('callExtendedTool', {
            tool: 'addSheet',
            arguments: { name: 'Forecast' },
          }),
        )
      ).content,
    )
    expect(added.sheet).toBe('Forecast')
    expect(beforeInvoke).toHaveBeenCalledWith(
      expect.objectContaining({
        name: 'addSheet',
        opClass: 'tool',
        risk: 'write',
        viaExtended: true,
      }),
    )
    expect(recordOperation).toHaveBeenCalledWith(
      harness.ctx,
      'extended.addSheet',
      expect.stringContaining('addSheet'),
    )

    // Destructive tool deleteSheet MUST refuse without confirmDestructive: true
    await expect(
      handlers.callExtendedTool(
        invoke('callExtendedTool', {
          tool: 'deleteSheet',
          arguments: { sheet: 'Forecast' },
        }),
      ),
    ).rejects.toThrow(/confirmDestructive: true/)

    // Succeeds when confirmDestructive: true is supplied
    const deleted = JSON.parse(
      (
        await handlers.callExtendedTool(
          invoke('callExtendedTool', {
            tool: 'deleteSheet',
            arguments: { sheet: 'Forecast' },
            confirmDestructive: true,
          }),
        )
      ).content,
    )
    expect(deleted.deleted).toBe('Forecast')

    // Refuses read-only, direct, and unknown tools
    await expect(
      handlers.callExtendedTool(
        invoke('callExtendedTool', {
          tool: 'findInWorkbook',
          arguments: { query: 'North' },
        }),
      ),
    ).rejects.toThrow(/read-only/)

    await expect(
      handlers.callExtendedTool(
        invoke('callExtendedTool', {
          tool: 'applyCellChanges',
          arguments: {},
        }),
      ),
    ).rejects.toThrow(/already directly exposed/)

    await expect(
      handlers.callExtendedTool(
        invoke('callExtendedTool', { tool: 'missingTool' }),
      ),
    ).rejects.toThrow(/Unknown extended tool/)
  })

  it('enforces beforeInvoke write guard uniformly across direct tools, functions, and extended tools', async () => {
    let mayChange = false
    const harness = makeHarness(makeRichWorkbook())
    const handlers = buildPureSheetsHandlers(() => harness.ctx, {
      beforeInvoke: guard => {
        if (
          guard.risk !== 'read' &&
          guard.name !== 'createWorkbook' &&
          !mayChange
        ) {
          throw new AgentSheetsToolError('Wait for workbook to finish opening.')
        }
      },
    })

    // Reads and createWorkbook are allowed while opening
    const brief = JSON.parse(
      (
        await handlers.fn_workbookBriefing(invoke('fn_workbookBriefing', {}))
      ).content,
    )
    expect(brief.ok).toBe(true)

    const created = JSON.parse(
      (
        await handlers.callExtendedTool(
          invoke('callExtendedTool', {
            tool: 'createWorkbook',
            arguments: { title: 'Fresh' },
          }),
        )
      ).content,
    )
    expect(created.title).toBe('Fresh')

    // Mutating direct, function, and extended calls are blocked
    await expect(
      handlers.completeWorkbook(invoke('completeWorkbook', {})),
    ).rejects.toThrow(/Wait for workbook/)
    await expect(
      handlers.fn_buildStructuredSheetTable(
        invoke('fn_buildStructuredSheetTable', {
          sheet: 'New',
          createSheet: true,
          headers: ['A'],
          rows: [['1']],
          summary: 'test',
        }),
      ),
    ).rejects.toThrow(/Wait for workbook/)
    await expect(
      handlers.callExtendedTool(
        invoke('callExtendedTool', {
          tool: 'addSheet',
          arguments: { name: 'Blocked' },
        }),
      ),
    ).rejects.toThrow(/Wait for workbook/)

    mayChange = true
    const allowed = JSON.parse(
      (await handlers.completeWorkbook(invoke('completeWorkbook', {}))).content,
    )
    expect(allowed.taskOutcome).toBe('saved')
  })
})

describe('Composite App Functions (fn_*)', () => {
  it('fn_workbookBriefing returns inventory, version, headers, computed previews, named ranges, and recent changes in 1 call', async () => {
    const { ctx } = makeHarness(makeRichWorkbook())
    const handlers = buildPureSheetsHandlers(() => ctx)

    const env = JSON.parse(
      (
        await handlers.fn_workbookBriefing(
          invoke('fn_workbookBriefing', { previewRows: 3 }),
        )
      ).content,
    )
    expect(env).toMatchObject({
      ok: true,
      opClass: 'function',
      function: 'fn_workbookBriefing',
    })
    expect(env.data.title).toBe('Q1 Model')
    expect(typeof env.data.version).toBe('string')
    expect(env.data.namedRanges).toHaveLength(1)
    expect(env.data.sheets).toHaveLength(2)
    expect(env.data.totals).toMatchObject({
      sheets: 2,
      formulas: 6,
      formulaErrors: 0,
      validationFlags: 0,
    })
    const sales = env.data.sheets[0]
    expect(sales.name).toBe('Sales')
    expect(sales.usedRange).toBe('A1:D5')
    expect(sales.headers.map((h: { value: string }) => h.value)).toEqual([
      'Region',
      'Units',
      'Price',
      'Revenue',
    ])
    expect(sales.preview).toHaveLength(3)
    expect(sales.preview[1].values.D).toBe('250')
    expect(env.data.sheets[1].preview[1].values.B).toBe('1575')
    expect(env.data.recentChanges.length).toBeGreaterThanOrEqual(2)

    // Filtering to a single sheet works, unknown sheet rejects
    const single = JSON.parse(
      (
        await handlers.fn_workbookBriefing(
          invoke('fn_workbookBriefing', { sheet: 'Summary' }),
        )
      ).content,
    )
    expect(single.data.sheets).toHaveLength(1)
    expect(single.data.sheets[0].name).toBe('Summary')

    await expect(
      handlers.fn_workbookBriefing(
        invoke('fn_workbookBriefing', { sheet: 'Nonexistent' }),
      ),
    ).rejects.toThrow(/No sheet matching/)
  })

  it('fn_auditWorkbookFormulasAndData detects formula errors, broken sheet refs, validation violations, hardcoded overrides, and text numbers', async () => {
    // Clean workbook passes cleanly
    const cleanHarness = makeHarness(makeRichWorkbook())
    const cleanHandlers = buildPureSheetsHandlers(() => cleanHarness.ctx)
    const cleanAudit = JSON.parse(
      (
        await cleanHandlers.fn_auditWorkbookFormulasAndData(
          invoke('fn_auditWorkbookFormulasAndData', {}),
        )
      ).content,
    )
    expect(cleanAudit.data.healthy).toBe(true)
    expect(cleanAudit.data.counts).toMatchObject({
      formulas: 6,
      formulaErrors: 0,
      brokenSheetRefs: 0,
      validationViolations: 0,
      hardcodedOverrides: 0,
    })
    expect(cleanAudit.data.crossSheetDependencies).toEqual([
      { edge: 'Summary -> Sales', referenceCount: 2 },
    ])

    // Now inject real spreadsheet defects
    let flawed = makeRichWorkbook()
    const sales = flawed.workbook.sheets[0]
    // Hardcoded override in formula column D (3 formulas + 1 hardcoded number = 75% formulas)
    flawed = applyAgentCellChangesDirect(
      flawed,
      sales.id,
      'A6:D6',
      'Add overridden & text-number row',
      [
        { cell: 'A6', value: 'East', kind: 'text' },
        { cell: 'B6', value: '$42', kind: 'text' }, // numeric-looking string in numeric col B
        { cell: 'C6', value: '20', kind: 'number' },
        { cell: 'D6', value: '999', kind: 'number' }, // hardcoded override in formula col D
      ],
    ).document

    // Add validation rule on C2:C6 requiring <= 35 (C5 is 40 -> violation)
    flawed.workbook.sheets[0] = addValidationRule(flawed.workbook.sheets[0], {
      range: 'C2:C6',
      kind: 'number-range',
      min: 1,
      max: 35,
    })

    // Add formula errors (#DIV/0!) and a broken cross-sheet reference on Summary
    const summarySheet = flawed.workbook.sheets[1]
    flawed = applyAgentCellChangesDirect(
      flawed,
      summarySheet.id,
      'B4:B5',
      'Add defective formulas',
      [
        { cell: 'B4', value: '=100/0', kind: 'formula' },
        { cell: 'B5', value: '=MissingSheet!A1+1', kind: 'formula' },
      ],
    ).document

    const flawedHarness = makeHarness(flawed)
    const flawedHandlers = buildPureSheetsHandlers(() => flawedHarness.ctx)
    const audit = JSON.parse(
      (
        await flawedHandlers.fn_auditWorkbookFormulasAndData(
          invoke('fn_auditWorkbookFormulasAndData', {}),
        )
      ).content,
    )
    expect(audit.data.healthy).toBe(false)
    expect(audit.data.counts.formulaErrors).toBeGreaterThanOrEqual(1)
    expect(audit.data.errorBreakdown['#DIV/0!']).toBe(1)
    expect(audit.data.counts.brokenSheetRefs).toBe(1)
    expect(audit.data.brokenSheetRefs[0]).toMatchObject({
      sheet: 'Summary',
      cell: 'B5',
      missingSheet: 'MissingSheet',
    })
    expect(audit.data.counts.validationViolations).toBe(1)
    expect(audit.data.validationViolations[0]).toMatchObject({
      sheet: 'Sales',
      cell: 'C5',
    })
    expect(audit.data.counts.hardcodedOverrides).toBe(1)
    expect(audit.data.hardcodedOverrides[0]).toMatchObject({
      sheet: 'Sales',
      cell: 'D6',
      column: 'D',
    })
    expect(audit.data.counts.numericTextCells).toBe(1)
    expect(audit.data.numericTextCells[0]).toMatchObject({
      sheet: 'Sales',
      cell: 'B6',
      value: '$42',
    })
  })

  it('fn_profileSheetData computes exact column statistics and grouped rollups', async () => {
    const { ctx } = makeHarness(makeRichWorkbook())
    const handlers = buildPureSheetsHandlers(() => ctx)

    const profile = JSON.parse(
      (
        await handlers.fn_profileSheetData(
          invoke('fn_profileSheetData', {
            sheet: 'Sales',
            groupByColumn: 'Region',
            valueColumn: 'Revenue',
          }),
        )
      ).content,
    )
    expect(profile.ok).toBe(true)
    expect(profile.data).toMatchObject({
      sheet: 'Sales',
      range: 'A1:D5',
      dataRowCount: 4,
      columnCount: 4,
    })

    const revenueCol = profile.data.columns.find(
      (c: { header: string }) => c.header === 'Revenue',
    )
    expect(revenueCol.inferredType).toBe('number')
    expect(revenueCol.formulas).toBe(4)
    expect(revenueCol.stats).toEqual({
      min: 125,
      max: 600,
      sum: 1575,
      mean: 393.75,
      median: 425,
    })

    expect(profile.data.groupedRollup.groups).toEqual([
      {
        group: 'South',
        rowCount: 1,
        numericCount: 1,
        sum: 600,
        avg: 600,
        min: 600,
        max: 600,
      },
      {
        group: 'West',
        rowCount: 1,
        numericCount: 1,
        sum: 600,
        avg: 600,
        min: 600,
        max: 600,
      },
      {
        group: 'North',
        rowCount: 2,
        numericCount: 2,
        sum: 375,
        avg: 187.5,
        min: 125,
        max: 250,
      },
    ])
  })

  it('fn_buildStructuredSheetTable supports dryRun, atomic new-sheet build with auto-sum, formatting, notes, chart, autoFitRows, and existing-sheet version checks', async () => {
    const harness = makeHarness(makeRichWorkbook())
    const recordOperation = vi.fn()
    const handlers = buildPureSheetsHandlers(() => harness.ctx, {
      recordOperation,
    })

    // 1. dryRun validates without mutating
    const dry = JSON.parse(
      (
        await handlers.fn_buildStructuredSheetTable(
          invoke('fn_buildStructuredSheetTable', {
            sheet: 'Forecast',
            createSheet: true,
            headers: ['Quarter', 'Revenue', 'Margin'],
            rows: [
              ['Q1', 1000, '=B2*0.2'],
              ['Q2', 1250, '=B3*0.22'],
            ],
            summaryRow: { label: 'FY Total', mode: 'auto-sum' },
            summary: 'Build Q1-Q2 forecast table',
            dryRun: true,
          }),
        )
      ).content,
    )
    expect(dry.data).toMatchObject({
      dryRun: true,
      range: 'A1:C4',
      cellCount: 12,
    })
    expect(dry.data.summaryFormulas).toEqual([
      { cell: 'A4', value: 'FY Total' },
      { cell: 'B4', value: '=SUM(B2:B3)' },
      { cell: 'C4', value: '=SUM(C2:C3)' },
    ])
    expect(harness.setCount).toBe(0)
    expect(recordOperation).not.toHaveBeenCalled()

    // 2. Real commit builds sheet, table, auto-sum, formats, notes, chart, and autoFitRows in 1 commit
    const built = JSON.parse(
      (
        await handlers.fn_buildStructuredSheetTable(
          invoke('fn_buildStructuredSheetTable', {
            sheet: 'Forecast',
            createSheet: true,
            headers: ['Quarter', 'Revenue', 'Margin'],
            rows: [
              ['Q1', 1000, '=B2*0.2'],
              ['Q2', 1250, '=B3*0.22'],
            ],
            summaryRow: { label: 'FY Total', mode: 'auto-sum' },
            columnFormats: [
              { column: 'Revenue', numberFormat: 'currency', decimals: 0, columnWidth: 140 },
              { column: 2, numberFormat: 'currency', decimals: 2 },
            ],
            notes: [{ cell: 'C2', comment: '20% target margin assumption' }],
            chart: { type: 'bar' },
            autoFitRows: true,
            summary: 'Build Q1-Q2 forecast table',
          }),
        )
      ).content,
    )
    expect(harness.setCount).toBe(1)
    expect(recordOperation).toHaveBeenCalledTimes(1)
    expect(built.data).toMatchObject({
      dryRun: false,
      sheet: 'Forecast',
      createdSheet: true,
      range: 'A1:C4',
      applied: 12,
      notesApplied: 1,
      chartAdded: true,
      rowsAutoFitted: true,
      formulaErrors: [],
    })
    expect(built.data.summaryReadback).toEqual([
      { cell: 'A4', value: 'FY Total', computed: 'FY Total' },
      { cell: 'B4', value: '=SUM(B2:B3)', computed: '2250' },
      { cell: 'C4', value: '=SUM(C2:C3)', computed: '475' },
    ])

    const forecast = harness
      .getDoc()
      .workbook.sheets.find(s => s.name === 'Forecast')!
    expect(forecast.cells.A1?.style?.bold).toBe(true)
    expect(forecast.cells.B2?.style?.numberFormat).toBe('currency')
    expect(forecast.columnWidths?.['1']).toBe(140)
    expect(forecast.comments?.C2).toBe('20% target margin assumption')
    expect(forecast.charts).toHaveLength(1)

    // 3. Refuses duplicate sheet name
    await expect(
      handlers.fn_buildStructuredSheetTable(
        invoke('fn_buildStructuredSheetTable', {
          sheet: 'Forecast',
          createSheet: true,
          headers: ['A'],
          rows: [['1']],
          summary: 'Duplicate',
        }),
      ),
    ).rejects.toThrow(/already exists/)

    // 4. Existing-sheet mode rejects missing/stale expectedVersion
    await expect(
      handlers.fn_buildStructuredSheetTable(
        invoke('fn_buildStructuredSheetTable', {
          sheet: 'Forecast',
          expectedVersion: 'stale-version',
          startCell: 'E1',
          headers: ['K'],
          rows: [['1']],
          summary: 'Stale write',
        }),
      ),
    ).rejects.toThrow(/expectedVersion/)
  })

  it('fn_applyFormulaAndFormatBatch supports dryRun, atomic edits + multi-range formats + comments + formula readback, and enforces locks & version', async () => {
    const doc = makeRichWorkbook()
    // Lock A2 on Sales so we verify locked-cell protection
    doc.workbook.sheets[0] = lockRange(doc.workbook.sheets[0], 'A2:A2')
    const harness = makeHarness(doc)
    const handlers = buildPureSheetsHandlers(() => harness.ctx)

    const ctxBefore = JSON.parse(
      (await handlers.getSheetsContext(invoke('getSheetsContext', {}))).content,
    )

    // 1. dryRun detects locked cells without mutating
    const dry = JSON.parse(
      (
        await handlers.fn_applyFormulaAndFormatBatch(
          invoke('fn_applyFormulaAndFormatBatch', {
            sheet: 'Sales',
            expectedVersion: ctxBefore.version,
            summary: 'Add tax column and format',
            cellChanges: [
              { cell: 'A2', value: 'Tampered' },
              { cell: 'E1', value: 'Tax' },
              { cell: 'E2', value: '=D2*0.1' },
              { cell: 'E3', value: '=D3*0.1' },
            ],
            formats: [
              { range: 'E1:E1', bold: true },
              { range: 'E2:E3', numberFormat: 'currency', decimals: 2 },
            ],
            comments: [{ cell: 'E1', comment: '10% VAT' }],
            dryRun: true,
          }),
        )
      ).content,
    )
    expect(dry.data).toMatchObject({
      dryRun: true,
      plannedCellChanges: 4,
      plannedFormats: 2,
      plannedComments: 1,
      lockedCellSkips: ['A2'],
    })
    expect(harness.setCount).toBe(0)

    // 2. Real commit applies unlocked edits, skips locked A2, applies formats + comments, and returns computed readback
    const applied = JSON.parse(
      (
        await handlers.fn_applyFormulaAndFormatBatch(
          invoke('fn_applyFormulaAndFormatBatch', {
            sheet: 'Sales',
            expectedVersion: ctxBefore.version,
            summary: 'Add tax column and format',
            cellChanges: [
              { cell: 'A2', value: 'Tampered' },
              { cell: 'E1', value: 'Tax' },
              { cell: 'E2', value: '=D2*0.1' },
              { cell: 'E3', value: '=D3*0.1' },
            ],
            formats: [
              { range: 'E1:E1', bold: true },
              { range: 'E2:E3', numberFormat: 'currency', decimals: 2 },
            ],
            comments: [{ cell: 'E1', comment: '10% VAT' }],
          }),
        )
      ).content,
    )
    expect(harness.setCount).toBe(1)
    expect(applied.data).toMatchObject({
      dryRun: false,
      appliedCells: 3,
      skippedLockedCells: 1,
      appliedFormats: 2,
      appliedComments: 1,
      formulaErrorCount: 0,
    })
    expect(applied.data.formulaReadback).toEqual([
      { cell: 'E2', formula: '=D2*0.1', computed: '25', isError: false },
      { cell: 'E3', formula: '=D3*0.1', computed: '60', isError: false },
    ])
    expect(applied.data.version).not.toBe(ctxBefore.version)
    expect(harness.getDoc().workbook.sheets[0].cells.A2?.value).toBe('North')
    expect(harness.getDoc().workbook.sheets[0].comments?.E1).toBe('10% VAT')

    // 3. Stale version is rejected
    await expect(
      handlers.fn_applyFormulaAndFormatBatch(
        invoke('fn_applyFormulaAndFormatBatch', {
          sheet: 'Sales',
          expectedVersion: ctxBefore.version,
          summary: 'Stale',
          cellChanges: [{ cell: 'F1', value: 'x' }],
        }),
      ),
    ).rejects.toThrow(/expectedVersion/)

    // 4. Out-of-range cellChange is rejected
    await expect(
      handlers.fn_applyFormulaAndFormatBatch(
        invoke('fn_applyFormulaAndFormatBatch', {
          sheet: 'Sales',
          expectedVersion: applied.data.version,
          range: 'E1:E2',
          summary: 'Out of bounds',
          cellChanges: [{ cell: 'F1', value: 'x' }],
        }),
      ),
    ).rejects.toThrow(/outside declared target range/)

    // 5. Resizing on a protected sheet is refused
    await expect(
      handlers.fn_applyFormulaAndFormatBatch(
        invoke('fn_applyFormulaAndFormatBatch', {
          sheet: 'Sales',
          expectedVersion: applied.data.version,
          summary: 'Resize protected',
          formats: [{ range: 'E1:E3', columnWidth: 180 }],
        }),
      ),
    ).rejects.toThrow(/protected sheets/)
  })

  it('preserves core atomic tools (getSheetsContext, readRange, applyCellChanges, completeWorkbook)', async () => {
    const harness = makeHarness(makeRichWorkbook())
    const handlers = buildPureSheetsHandlers(() => harness.ctx)

    const ctxInfo = JSON.parse(
      (await handlers.getSheetsContext(invoke('getSheetsContext', {}))).content,
    )
    expect(ctxInfo.title).toBe('Q1 Model')

    const write = JSON.parse(
      (
        await handlers.applyCellChanges(
          invoke('applyCellChanges', {
            sheet: 'Sales',
            range: 'F1:F2',
            expectedVersion: ctxInfo.version,
            summary: 'Add direct atomic cell',
            cellChanges: [
              { cell: 'F1', value: 'Note', kind: 'text' },
              { cell: 'F2', value: '=D2*2', kind: 'formula' },
            ],
          }),
        )
      ).content,
    )
    expect(write.applied).toBe(2)

    const read = JSON.parse(
      (
        await handlers.readRange(
          invoke('readRange', { sheet: 'Sales', range: 'F1:F2' }),
        )
      ).content,
    )
    expect(read.cells[1]).toMatchObject({ cell: 'F2', computed: '500' })

    const receipt = JSON.parse(
      (await handlers.completeWorkbook(invoke('completeWorkbook', {}))).content,
    )
    expect(receipt.artifactPaths).toEqual(['/workspace/budget.sheets'])
  })
})
