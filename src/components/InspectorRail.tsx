import { useMemo, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import {
  ArrowDown,
  ArrowUp,
  RefreshCw,
  RotateCcw,
  Trash2,
  X,
} from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import type {
  AgentLogEntry,
  CellStyle,
  ConditionalFormatRule,
  NamedRange,
  PivotAggregate,
  PivotDefinition,
  PureSheetsDocument,
  ValidationRule,
  WorkbookSheet,
} from '../types'
import {
  extractFormulaPrecedents,
  traceDependents,
} from '../lib/formulaCatalog'
import { inferPivotFields, isPivotStale } from '../lib/pivotEngine'
import { computedCellValue, isFormulaError } from '../lib/workbookModel'
import type { InvalidCellFlag } from '../lib/sheetDataFeatures'

export type InspectorTab =
  | 'rules'
  | 'validate'
  | 'pivot'
  | 'names'
  | 'audit'
  | 'activity'

const CF_KINDS: Array<{
  id: ConditionalFormatRule['condition']['kind']
  label: string
}> = [
  { id: 'greater', label: '>' },
  { id: 'less', label: '<' },
  { id: 'equal', label: '=' },
  { id: 'not-equal', label: '≠' },
  { id: 'between', label: 'Between' },
  { id: 'text-contains', label: 'Contains' },
  { id: 'blank', label: 'Blank' },
  { id: 'not-blank', label: 'Not blank' },
  { id: 'duplicate', label: 'Duplicates' },
  { id: 'top-n', label: 'Top N' },
  { id: 'bottom-n', label: 'Bottom N' },
  { id: 'color-scale', label: 'Color scale' },
  { id: 'data-bar', label: 'Data bar' },
  { id: 'custom-formula', label: 'Formula' },
]

const CF_PRESETS: Array<{ id: string; label: string; style: CellStyle }> = [
  { id: 'red', label: 'Red', style: { textColor: '#982d25', fillColor: '#fdecea', bold: true } },
  { id: 'green', label: 'Green', style: { textColor: '#1f5c39', fillColor: '#e5f3df' } },
  { id: 'amber', label: 'Amber', style: { textColor: '#7a4a0b', fillColor: '#fff4bf' } },
  { id: 'blue', label: 'Blue', style: { textColor: '#1a5683', fillColor: '#e4f0fb' } },
]

const PIVOT_AGGS: PivotAggregate[] = [
  'sum',
  'count',
  'avg',
  'min',
  'max',
  'percent-of-total',
]

export function InspectorRail({
  tab,
  onSelectTab,
  onClose,
  document,
  activeSheet,
  selectedCell,
  selectedRange,
  invalidFlags,
  showFormulas,
  onToggleShowFormulas,
  onJumpToCell,
  // Conditional formatting
  onAddConditionalRule,
  onDeleteConditionalRule,
  onMoveConditionalRule,
  // Validation
  onAddValidationRule,
  onDeleteValidationRule,
  // Named ranges
  onSaveNamedRange,
  onDeleteNamedRange,
  onJumpToNamedRange,
  // Pivot
  onCreatePivotFromSelection,
  onUpdatePivotDefinition,
  onRefreshPivot,
  // Agent activity
  onRevertAgentEntry,
}: {
  tab: InspectorTab
  onSelectTab: (tab: InspectorTab) => void
  onClose: () => void
  document: PureSheetsDocument
  activeSheet: WorkbookSheet
  selectedCell: string
  selectedRange: string
  invalidFlags: InvalidCellFlag[]
  showFormulas: boolean
  onToggleShowFormulas: () => void
  onJumpToCell: (sheetId: string, rangeOrCell: string) => void
  onAddConditionalRule: (rule: Omit<ConditionalFormatRule, 'id'>) => void
  onDeleteConditionalRule: (ruleId: string) => void
  onMoveConditionalRule: (ruleId: string, delta: -1 | 1) => void
  onAddValidationRule: (rule: Omit<ValidationRule, 'id'>) => void
  onDeleteValidationRule: (ruleId: string) => void
  onSaveNamedRange: (name: string, range: string, comment?: string) => void
  onDeleteNamedRange: (id: string) => void
  onJumpToNamedRange: (nr: NamedRange) => void
  onCreatePivotFromSelection: () => void
  onUpdatePivotDefinition: (def: PivotDefinition) => void
  onRefreshPivot: () => void
  onRevertAgentEntry: (entry: AgentLogEntry) => void
}): React.ReactElement {
  // CF builder state
  const [cfKind, setCfKind] =
    useState<ConditionalFormatRule['condition']['kind']>('greater')
  const [cfVal1, setCfVal1] = useState('')
  const [cfVal2, setCfVal2] = useState('')
  const [cfPreset, setCfPreset] = useState('red')

  // Validation builder state
  const [valKind, setValKind] = useState<ValidationRule['kind']>('list')
  const [valListText, setValListText] = useState('')
  const [valSourceRange, setValSourceRange] = useState('')
  const [valMin, setValMin] = useState('')
  const [valMax, setValMax] = useState('')
  const [valQuery, setValQuery] = useState('')
  const [valFormula, setValFormula] = useState('')

  // Named range state
  const [nrName, setNrName] = useState('')
  const [nrComment, setNrComment] = useState('')

  // Formula audit memo
  const selectedCellObj = activeSheet.cells[selectedCell]
  const precedents = useMemo(() => {
    if (selectedCellObj?.kind !== 'formula') return []
    return extractFormulaPrecedents(selectedCellObj.value, activeSheet.name)
  }, [selectedCellObj, activeSheet.name])

  const dependents = useMemo(
    () => traceDependents(document.workbook, activeSheet.id, selectedCell),
    [document.workbook, activeSheet.id, selectedCell],
  )

  const formulaErrors = useMemo(() => {
    const errs: Array<{ cell: string; error: string; formula: string }> = []
    for (const [key, cell] of Object.entries(activeSheet.cells)) {
      if (cell.kind !== 'formula') continue
      const res = computedCellValue(activeSheet, key, document.workbook)
      if (isFormulaError(res)) {
        errs.push({ cell: key, error: res, formula: cell.value })
      }
    }
    return errs
  }, [activeSheet, document.workbook])

  // Pivot state
  const pivotDef = activeSheet.pivot
  const pivotSourceSheet = pivotDef
    ? document.workbook.sheets.find(s => s.id === pivotDef.sourceSheetId)
    : undefined
  const pivotFields = useMemo(() => {
    if (!pivotDef || !pivotSourceSheet) return []
    return inferPivotFields(pivotSourceSheet, pivotDef.sourceRange, document.workbook)
  }, [pivotDef, pivotSourceSheet, document.workbook])
  const pivotStale = pivotDef
    ? isPivotStale(activeSheet, pivotSourceSheet, document.workbook)
    : false

  const tabs: Array<{ id: InspectorTab; label: string }> = [
    { id: 'rules', label: 'Rules' },
    { id: 'validate', label: 'Valid' },
    { id: 'pivot', label: 'Pivot' },
    { id: 'names', label: 'Names' },
    { id: 'audit', label: 'Audit' },
    { id: 'activity', label: 'Log' },
  ]

  function handleAddCf(): void {
    const needsValue =
      cfKind !== 'blank' &&
      cfKind !== 'not-blank' &&
      cfKind !== 'duplicate' &&
      cfKind !== 'color-scale' &&
      cfKind !== 'data-bar'
    if (needsValue && !cfVal1.trim()) return
    const style =
      CF_PRESETS.find(p => p.id === cfPreset)?.style ?? CF_PRESETS[0].style
    onAddConditionalRule({
      range: selectedRange,
      condition: {
        kind: cfKind,
        value: cfVal1.trim(),
        value2: cfKind === 'between' ? cfVal2.trim() || undefined : undefined,
      },
      style: cfKind === 'color-scale' || cfKind === 'data-bar' ? {} : style,
      ...(cfKind === 'color-scale'
        ? {
            colorScale: {
              minColor: '#fdecea',
              midColor: '#fff4bf',
              maxColor: '#e5f3df',
            },
          }
        : {}),
      ...(cfKind === 'data-bar' ? { dataBarColor: '#2f6ea2' } : {}),
    })
    setCfVal1('')
    setCfVal2('')
  }

  function handleAddValidation(): void {
    if (valKind === 'list') {
      const options = valListText
        .split(',')
        .map(s => s.trim())
        .filter(Boolean)
      onAddValidationRule({
        range: selectedRange,
        kind: 'list',
        ...(valSourceRange.trim()
          ? { sourceRange: valSourceRange.trim() }
          : options.length
          ? { options }
          : {}),
      })
    } else if (valKind === 'number-range' || valKind === 'text-length') {
      const min = valMin.trim() === '' ? undefined : Number(valMin)
      const max = valMax.trim() === '' ? undefined : Number(valMax)
      onAddValidationRule({
        range: selectedRange,
        kind: valKind,
        min: Number.isFinite(min as number) ? min : undefined,
        max: Number.isFinite(max as number) ? max : undefined,
      })
    } else if (valKind === 'text-contains') {
      onAddValidationRule({
        range: selectedRange,
        kind: 'text-contains',
        textQuery: valQuery.trim(),
      })
    } else if (valKind === 'custom-formula') {
      if (!valFormula.trim()) return
      onAddValidationRule({
        range: selectedRange,
        kind: 'custom-formula',
        formula: valFormula.trim(),
      })
    } else {
      onAddValidationRule({ range: selectedRange, kind: 'date' })
    }
  }

  return (
    <Rail aria-label="Sheet inspector">
      <RailHeader>
        <TabStrip role="tablist" aria-label="Inspector sections">
          {tabs.map(t => (
            <TabBtn
              key={t.id}
              type="button"
              role="tab"
              aria-selected={tab === t.id}
              $active={tab === t.id}
              onClick={() => onSelectTab(t.id)}
            >
              {t.label}
            </TabBtn>
          ))}
        </TabStrip>
        <CloseBtn type="button" aria-label="Close inspector" onClick={onClose}>
          <PlatformIcon icon={X} size={14} />
        </CloseBtn>
      </RailHeader>

      <RailBody>
        {tab === 'rules' && (
          <Section>
            <SectionHead>Conditional formatting · {selectedRange}</SectionHead>
            <ChipGrid>
              {CF_KINDS.map(k => (
                <Chip
                  key={k.id}
                  type="button"
                  $active={cfKind === k.id}
                  onClick={() => setCfKind(k.id)}
                >
                  {k.label}
                </Chip>
              ))}
            </ChipGrid>
            {cfKind !== 'blank' &&
              cfKind !== 'not-blank' &&
              cfKind !== 'duplicate' &&
              cfKind !== 'color-scale' &&
              cfKind !== 'data-bar' && (
                <Row>
                  <Input
                    value={cfVal1}
                    onChange={e => setCfVal1(e.currentTarget.value)}
                    placeholder={
                      cfKind === 'custom-formula'
                        ? '=A1>100'
                        : cfKind === 'top-n' || cfKind === 'bottom-n'
                        ? 'N (e.g. 5)'
                        : 'Value'
                    }
                    aria-label="Rule condition value"
                  />
                  {cfKind === 'between' && (
                    <Input
                      value={cfVal2}
                      onChange={e => setCfVal2(e.currentTarget.value)}
                      placeholder="Upper"
                      aria-label="Rule condition upper value"
                    />
                  )}
                </Row>
              )}
            {cfKind !== 'color-scale' && cfKind !== 'data-bar' && (
              <Row>
                {CF_PRESETS.map(p => (
                  <Chip
                    key={p.id}
                    type="button"
                    $active={cfPreset === p.id}
                    onClick={() => setCfPreset(p.id)}
                  >
                    {p.label}
                  </Chip>
                ))}
              </Row>
            )}
            <PrimaryBtn type="button" onClick={handleAddCf}>
              Add rule
            </PrimaryBtn>

            <SubHead>
              Active rules ({activeSheet.conditionalFormats?.length ?? 0})
            </SubHead>
            {(activeSheet.conditionalFormats ?? []).length === 0 ? (
              <Empty>No conditional rules on this sheet.</Empty>
            ) : (
              <ItemList>
                {(activeSheet.conditionalFormats ?? []).map((rule, idx, arr) => (
                  <ItemCard key={rule.id}>
                    <ItemInfo>
                      <ItemTitle>
                        {rule.range} · {rule.condition.kind}
                        {rule.condition.value ? ` ${rule.condition.value}` : ''}
                        {rule.condition.value2 ? `–${rule.condition.value2}` : ''}
                      </ItemTitle>
                      {rule.style.fillColor && (
                        <SwatchDot $color={rule.style.fillColor} />
                      )}
                    </ItemInfo>
                    <ItemActions>
                      <MiniBtn
                        type="button"
                        disabled={idx === 0}
                        aria-label="Move rule up"
                        onClick={() => onMoveConditionalRule(rule.id, -1)}
                      >
                        <PlatformIcon icon={ArrowUp} size={12} />
                      </MiniBtn>
                      <MiniBtn
                        type="button"
                        disabled={idx === arr.length - 1}
                        aria-label="Move rule down"
                        onClick={() => onMoveConditionalRule(rule.id, 1)}
                      >
                        <PlatformIcon icon={ArrowDown} size={12} />
                      </MiniBtn>
                      <MiniBtn
                        type="button"
                        aria-label="Delete rule"
                        onClick={() => onDeleteConditionalRule(rule.id)}
                      >
                        <PlatformIcon icon={Trash2} size={12} />
                      </MiniBtn>
                    </ItemActions>
                  </ItemCard>
                ))}
              </ItemList>
            )}
          </Section>
        )}

        {tab === 'validate' && (
          <Section>
            <SectionHead>Data validation · {selectedRange}</SectionHead>
            <ChipGrid>
              {(
                [
                  ['list', 'Dropdown list'],
                  ['number-range', 'Number'],
                  ['date', 'Date'],
                  ['text-length', 'Text length'],
                  ['text-contains', 'Text contains'],
                  ['custom-formula', 'Formula'],
                ] as const
              ).map(([id, label]) => (
                <Chip
                  key={id}
                  type="button"
                  $active={valKind === id}
                  onClick={() => setValKind(id)}
                >
                  {label}
                </Chip>
              ))}
            </ChipGrid>

            {valKind === 'list' && (
              <>
                <Input
                  value={valListText}
                  onChange={e => setValListText(e.currentTarget.value)}
                  placeholder="Options: Open, In Progress, Done"
                  aria-label="Validation list options"
                />
                <Input
                  value={valSourceRange}
                  onChange={e => setValSourceRange(e.currentTarget.value)}
                  placeholder="Or range: Sheet2!A2:A10"
                  aria-label="Validation source range"
                />
              </>
            )}
            {(valKind === 'number-range' || valKind === 'text-length') && (
              <Row>
                <Input
                  value={valMin}
                  onChange={e => setValMin(e.currentTarget.value)}
                  placeholder="Min"
                  aria-label="Validation min"
                />
                <Input
                  value={valMax}
                  onChange={e => setValMax(e.currentTarget.value)}
                  placeholder="Max"
                  aria-label="Validation max"
                />
              </Row>
            )}
            {valKind === 'text-contains' && (
              <Input
                value={valQuery}
                onChange={e => setValQuery(e.currentTarget.value)}
                placeholder="Must contain text…"
                aria-label="Validation text query"
              />
            )}
            {valKind === 'custom-formula' && (
              <Input
                value={valFormula}
                onChange={e => setValFormula(e.currentTarget.value)}
                placeholder="=A1>0"
                aria-label="Validation custom formula"
              />
            )}
            <PrimaryBtn type="button" onClick={handleAddValidation}>
              Apply validation
            </PrimaryBtn>

            <SubHead>
              Rules on sheet ({activeSheet.validations?.length ?? 0})
            </SubHead>
            {(activeSheet.validations ?? []).length === 0 ? (
              <Empty>No validation rules yet.</Empty>
            ) : (
              <ItemList>
                {(activeSheet.validations ?? []).map(r => (
                  <ItemCard key={r.id}>
                    <ItemInfo>
                      <ItemTitle>
                        {r.range} · {r.kind}
                      </ItemTitle>
                      <ItemSub>
                        {r.sourceRange
                          ? `From ${r.sourceRange}`
                          : r.options?.length
                          ? r.options.slice(0, 4).join(', ')
                          : r.formula ??
                            (r.min !== undefined || r.max !== undefined
                              ? `${r.min ?? '−∞'}..${r.max ?? '+∞'}`
                              : '')}
                      </ItemSub>
                    </ItemInfo>
                    <MiniBtn
                      type="button"
                      aria-label="Delete validation rule"
                      onClick={() => onDeleteValidationRule(r.id)}
                    >
                      <PlatformIcon icon={Trash2} size={12} />
                    </MiniBtn>
                  </ItemCard>
                ))}
              </ItemList>
            )}
          </Section>
        )}

        {tab === 'pivot' && (
          <Section>
            {!pivotDef ? (
              <>
                <SectionHead>Pivot Table</SectionHead>
                <Note>
                  Summarize any table into a protected derived Pivot sheet with
                  row grouping, column cross-tabs, and totals.
                </Note>
                <PrimaryBtn type="button" onClick={onCreatePivotFromSelection}>
                  Create Pivot from {selectedRange}
                </PrimaryBtn>
              </>
            ) : (
              <>
                <PivotHeadRow>
                  <div>
                    <SectionHead>Pivot · {pivotDef.sourceRange}</SectionHead>
                    <ItemSub>
                      Source: {pivotSourceSheet?.name ?? 'Missing sheet'}
                      {pivotStale ? ' · Stale' : ' · Up to date'}
                    </ItemSub>
                  </div>
                  <QuietBtn type="button" onClick={onRefreshPivot}>
                    <PlatformIcon icon={RefreshCw} size={12} /> Refresh
                  </QuietBtn>
                </PivotHeadRow>

                <SubHead>Rows (group by)</SubHead>
                <ChipGrid>
                  {pivotFields.map(f => {
                    const active = pivotDef.rows.some(
                      r => r.columnIndex === f.columnIndex,
                    )
                    return (
                      <Chip
                        key={`r-${f.columnIndex}`}
                        type="button"
                        $active={active}
                        onClick={() => {
                          const nextRows = active
                            ? pivotDef.rows.filter(
                                r => r.columnIndex !== f.columnIndex,
                              )
                            : [
                                ...pivotDef.rows,
                                { columnIndex: f.columnIndex, label: f.label },
                              ]
                          onUpdatePivotDefinition({ ...pivotDef, rows: nextRows })
                        }}
                      >
                        {f.label}
                      </Chip>
                    )
                  })}
                </ChipGrid>

                <SubHead>Columns (cross-tab)</SubHead>
                <ChipGrid>
                  {pivotFields.map(f => {
                    const active = pivotDef.columns.some(
                      c => c.columnIndex === f.columnIndex,
                    )
                    return (
                      <Chip
                        key={`c-${f.columnIndex}`}
                        type="button"
                        $active={active}
                        onClick={() => {
                          const nextCols = active
                            ? []
                            : [{ columnIndex: f.columnIndex, label: f.label }]
                          onUpdatePivotDefinition({
                            ...pivotDef,
                            columns: nextCols,
                          })
                        }}
                      >
                        {f.label}
                      </Chip>
                    )
                  })}
                </ChipGrid>

                <SubHead>Values</SubHead>
                <ChipGrid>
                  {pivotFields.map(f => {
                    const active = pivotDef.values.some(
                      v => v.columnIndex === f.columnIndex,
                    )
                    return (
                      <Chip
                        key={`v-${f.columnIndex}`}
                        type="button"
                        $active={active}
                        onClick={() => {
                          const nextVals = active
                            ? pivotDef.values.filter(
                                v => v.columnIndex !== f.columnIndex,
                              )
                            : [
                                ...pivotDef.values,
                                {
                                  columnIndex: f.columnIndex,
                                  label: f.label,
                                  aggregate: (f.isNumeric
                                    ? 'sum'
                                    : 'count') as PivotAggregate,
                                },
                              ]
                          onUpdatePivotDefinition({
                            ...pivotDef,
                            values: nextVals,
                          })
                        }}
                      >
                        {f.label}
                      </Chip>
                    )
                  })}
                </ChipGrid>

                {pivotDef.values.length > 0 && (
                  <ItemList>
                    {pivotDef.values.map((v, idx) => (
                      <ItemCard key={`${v.columnIndex}-${idx}`}>
                        <ItemTitle>{v.label}</ItemTitle>
                        <AggRow>
                          {PIVOT_AGGS.map(agg => (
                            <TinyChip
                              key={agg}
                              type="button"
                              $active={v.aggregate === agg}
                              onClick={() => {
                                const nextVals = pivotDef.values.map((item, i) =>
                                  i === idx ? { ...item, aggregate: agg } : item,
                                )
                                onUpdatePivotDefinition({
                                  ...pivotDef,
                                  values: nextVals,
                                })
                              }}
                            >
                              {agg === 'percent-of-total' ? '%' : agg}
                            </TinyChip>
                          ))}
                        </AggRow>
                      </ItemCard>
                    ))}
                  </ItemList>
                )}
              </>
            )}
          </Section>
        )}

        {tab === 'names' && (
          <Section>
            <SectionHead>Named ranges</SectionHead>
            <Input
              value={nrName}
              onChange={e => setNrName(e.currentTarget.value)}
              placeholder="Name (e.g. TaxRate, Q1_Sales)"
              aria-label="Named range name"
            />
            <Input
              value={nrComment}
              onChange={e => setNrComment(e.currentTarget.value)}
              placeholder="Note (optional)"
              aria-label="Named range note"
            />
            <PrimaryBtn
              type="button"
              onClick={() => {
                if (!nrName.trim()) return
                onSaveNamedRange(nrName.trim(), selectedRange, nrComment.trim())
                setNrName('')
                setNrComment('')
              }}
            >
              Define for {selectedRange}
            </PrimaryBtn>

            <SubHead>
              Workbook names ({document.workbook.namedRanges?.length ?? 0})
            </SubHead>
            {(document.workbook.namedRanges ?? []).length === 0 ? (
              <Empty>No named ranges defined yet.</Empty>
            ) : (
              <ItemList>
                {(document.workbook.namedRanges ?? []).map(nr => {
                  const targetSheet = document.workbook.sheets.find(
                    s => s.id === nr.sheetId,
                  )
                  return (
                    <ItemCard key={nr.id}>
                      <JumpBtn
                        type="button"
                        onClick={() => onJumpToNamedRange(nr)}
                      >
                        <ItemTitle>{nr.name}</ItemTitle>
                        <ItemSub>
                          {targetSheet?.name ?? '?'}!{nr.range}
                        </ItemSub>
                      </JumpBtn>
                      <MiniBtn
                        type="button"
                        aria-label={`Delete named range ${nr.name}`}
                        onClick={() => onDeleteNamedRange(nr.id)}
                      >
                        <PlatformIcon icon={Trash2} size={12} />
                      </MiniBtn>
                    </ItemCard>
                  )
                })}
              </ItemList>
            )}
          </Section>
        )}

        {tab === 'audit' && (
          <Section>
            <SectionHead>Formula & sheet audit · {selectedCell}</SectionHead>
            <QuietBtn type="button" onClick={onToggleShowFormulas}>
              {showFormulas ? 'Show calculated values' : 'Show all formulas (Ctrl+`)'}
            </QuietBtn>

            <SubHead>Precedents ({precedents.length})</SubHead>
            {precedents.length === 0 ? (
              <Empty>Selected cell references no other cells.</Empty>
            ) : (
              <ItemList>
                {precedents.map(ref => {
                  const targetSheet = document.workbook.sheets.find(
                    s => s.name.toLowerCase() === ref.sheetName.toLowerCase(),
                  )
                  return (
                    <ItemCard key={ref.token}>
                      <JumpBtn
                        type="button"
                        onClick={() =>
                          onJumpToCell(targetSheet?.id ?? activeSheet.id, ref.range)
                        }
                      >
                        <ItemTitle>
                          {ref.sheetName}!{ref.range}
                        </ItemTitle>
                      </JumpBtn>
                    </ItemCard>
                  )
                })}
              </ItemList>
            )}

            <SubHead>Dependents ({dependents.length})</SubHead>
            {dependents.length === 0 ? (
              <Empty>No formulas depend on {selectedCell}.</Empty>
            ) : (
              <ItemList>
                {dependents.map(dep => (
                  <ItemCard key={`${dep.sheetId}-${dep.cell}`}>
                    <JumpBtn
                      type="button"
                      onClick={() => onJumpToCell(dep.sheetId, dep.cell)}
                    >
                      <ItemTitle>
                        {dep.sheetName}!{dep.cell}
                      </ItemTitle>
                      <ItemSub>{dep.formula}</ItemSub>
                    </JumpBtn>
                  </ItemCard>
                ))}
              </ItemList>
            )}

            <SubHead>Formula errors ({formulaErrors.length})</SubHead>
            {formulaErrors.length === 0 ? (
              <Empty>No formula errors on this sheet.</Empty>
            ) : (
              <ItemList>
                {formulaErrors.map(err => (
                  <ItemCard key={err.cell}>
                    <JumpBtn
                      type="button"
                      onClick={() => onJumpToCell(activeSheet.id, err.cell)}
                    >
                      <ItemTitle>
                        {err.cell} · {err.error}
                      </ItemTitle>
                      <ItemSub>{err.formula}</ItemSub>
                    </JumpBtn>
                  </ItemCard>
                ))}
              </ItemList>
            )}

            <SubHead>Invalid cells ({invalidFlags.length})</SubHead>
            {invalidFlags.length === 0 ? (
              <Empty>All validated cells pass.</Empty>
            ) : (
              <ItemList>
                {invalidFlags.map(flag => (
                  <ItemCard key={flag.cell}>
                    <JumpBtn
                      type="button"
                      onClick={() => onJumpToCell(activeSheet.id, flag.cell)}
                    >
                      <ItemTitle>
                        {flag.cell} · {flag.value}
                      </ItemTitle>
                      <ItemSub>{flag.message}</ItemSub>
                    </JumpBtn>
                  </ItemCard>
                ))}
              </ItemList>
            )}
          </Section>
        )}

        {tab === 'activity' && (
          <Section>
            <SectionHead>
              Agent activity ({(document.workbook.agentLog ?? []).length})
            </SectionHead>
            {(document.workbook.agentLog ?? []).length === 0 ? (
              <Empty>No agent writes recorded for this workbook.</Empty>
            ) : (
              <ItemList>
                {(document.workbook.agentLog ?? []).map(entry => (
                  <ItemCard key={entry.id}>
                    <ItemInfo>
                      <ItemTitle>{entry.summary}</ItemTitle>
                      <ItemSub>
                        {entry.sheetName}
                        {entry.range ? ` · ${entry.range}` : ''} · {entry.tool}
                      </ItemSub>
                    </ItemInfo>
                    {entry.cells && entry.cells.length > 0 && (
                      <MiniBtn
                        type="button"
                        title="Revert these cell changes"
                        aria-label={`Revert ${entry.summary}`}
                        onClick={() => onRevertAgentEntry(entry)}
                      >
                        <PlatformIcon icon={RotateCcw} size={12} />
                      </MiniBtn>
                    )}
                  </ItemCard>
                ))}
              </ItemList>
            )}
          </Section>
        )}
      </RailBody>
    </Rail>
  )
}

const Rail = styled.aside`
  width: 316px;
  flex: 0 0 316px;
  display: flex;
  flex-direction: column;
  min-height: 0;
  border-left: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-surface);
`

const RailHeader = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 4px;
  padding: 6px 8px;
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--n-panel, var(--pure-chrome-bar));
`

const TabStrip = styled.div`
  display: flex;
  gap: 2px;
  overflow-x: auto;
`

const TabBtn = styled.button<{ $active?: boolean }>`
  height: 24px;
  padding: 0 7px;
  border: 0;
  background: ${({ $active }) =>
    $active ? 'var(--p-chosen, var(--pure-chrome-selection))' : 'transparent'};
  box-shadow: ${({ $active }) =>
    $active ? 'var(--p-chosen-edge, none)' : 'none'};
  color: ${({ $active }) =>
    $active ? 'var(--sheets-ink)' : 'var(--pure-chrome-soft)'};
  font-size: 11.5px;
  font-weight: ${({ $active }) => ($active ? 700 : 500)};
  cursor: pointer;
`

const CloseBtn = styled.button`
  display: inline-grid;
  place-items: center;
  width: 24px;
  height: 24px;
  border: 0;
  background: transparent;
  color: var(--pure-chrome-soft);
  cursor: pointer;
`

const RailBody = styled.div`
  flex: 1 1 auto;
  overflow-y: auto;
  padding: 10px;
`

const Section = styled.div`
  display: grid;
  gap: 8px;
`

const SectionHead = styled.div`
  font-size: 12.5px;
  font-weight: 700;
  color: var(--sheets-ink);
`

const SubHead = styled.div`
  margin-top: 6px;
  font-size: 10.5px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: var(--pure-chrome-muted);
`

const Note = styled.p`
  margin: 0;
  font-size: 12px;
  line-height: 1.45;
  color: var(--pure-chrome-soft);
`

const Empty = styled.div`
  font-size: 12px;
  color: var(--pure-chrome-muted);
`

const ChipGrid = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
`

const Chip = styled.button<{ $active?: boolean }>`
  height: 24px;
  padding: 0 7px;
  border: 1px solid
    ${({ $active }) => ($active ? 'var(--sheets-accent)' : 'var(--pure-chrome-line)')};
  border-radius: 2px;
  background: ${({ $active }) =>
    $active ? 'var(--p-chosen, var(--pure-chrome-selection))' : 'var(--pure-chrome-surface)'};
  color: var(--sheets-ink);
  font-size: 11.5px;
  cursor: pointer;
`

const TinyChip = styled(Chip)`
  height: 20px;
  padding: 0 5px;
  font-size: 10.5px;
`

const Row = styled.div`
  display: flex;
  gap: 6px;
`

const PivotHeadRow = styled.div`
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 8px;
`

const Input = styled.input`
  width: 100%;
  box-sizing: border-box;
  height: 28px;
  padding: 0 8px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-size: 12px;
`

const PrimaryBtn = styled.button`
  height: 28px;
  padding: 0 10px;
  border: 1px solid var(--pure-chrome-accent);
  border-radius: 2px;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
`

const QuietBtn = styled.button`
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 5px;
  height: 26px;
  padding: 0 8px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-size: 11.5px;
  font-weight: 600;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const ItemList = styled.div`
  display: grid;
  gap: 5px;
`

const ItemCard = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
  padding: 6px 8px;
  border: 1px solid var(--pure-chrome-fence, var(--pure-chrome-line));
  background: var(--n-panel, var(--pure-chrome-surface));
`

const ItemInfo = styled.div`
  display: grid;
  gap: 2px;
  min-width: 0;
  flex: 1 1 auto;
`

const JumpBtn = styled.button`
  display: grid;
  gap: 2px;
  min-width: 0;
  flex: 1 1 auto;
  border: 0;
  padding: 0;
  background: transparent;
  text-align: left;
  cursor: pointer;
`

const ItemTitle = styled.div`
  font-size: 12px;
  font-weight: 600;
  color: var(--sheets-ink);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`

const ItemSub = styled.div`
  font-family: var(--sheets-mono);
  font-size: 10.5px;
  color: var(--pure-chrome-muted);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`

const ItemActions = styled.div`
  display: inline-flex;
  gap: 2px;
  flex: 0 0 auto;
`

const AggRow = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 2px;
`

const SwatchDot = styled.span<{ $color: string }>`
  width: 14px;
  height: 6px;
  border: 1px solid var(--pure-chrome-line);
  background: ${({ $color }) => $color};
`

const MiniBtn = styled.button`
  display: inline-grid;
  place-items: center;
  width: 22px;
  height: 22px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--pure-chrome-soft);
  cursor: pointer;
  &:disabled {
    opacity: 0.35;
    cursor: default;
  }
  &:hover:not(:disabled) {
    background: var(--pure-chrome-hover);
    color: var(--sheets-ink);
  }
`
