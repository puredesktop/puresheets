import { useEffect, useMemo, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { Filter, X } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import type { ColumnFilterRule, WorkbookSheet } from '../types'
import { profileColumn } from '../lib/dataTransforms'
import { columnLabel, type FormulaWorkbook } from '../lib/workbookModel'

const OPERATORS: Array<{
  id: '' | NonNullable<ColumnFilterRule['operator']>
  label: string
}> = [
  { id: '', label: 'Any condition' },
  { id: 'contains', label: 'Contains' },
  { id: 'not-contains', label: 'Does not contain' },
  { id: 'equals', label: 'Equals' },
  { id: 'gt', label: 'Greater than (>)' },
  { id: 'lt', label: 'Less than (<)' },
  { id: 'between', label: 'Between' },
  { id: 'blank', label: 'Is blank' },
  { id: 'not-blank', label: 'Is not blank' },
]

function compactNum(n: number): string {
  return Number.isInteger(n) ? String(n) : String(Number(n.toFixed(2)))
}

/**
 * Smart multi-column filter popover (Wave 3):
 * - Instant mini column profile (unique, blanks, min/median/max + spark-histogram)
 * - Condition operator (contains, equals, >, <, between, blank/not-blank)
 * - Distinct value checklist with counts + Select All / Clear
 */
export function SmartFilterControl({
  sheet,
  workbook,
  column,
  activeFilterCount,
  onApplyColumnFilter,
  onClearAllFilters,
}: {
  sheet: WorkbookSheet
  workbook?: FormulaWorkbook
  column: number
  activeFilterCount: number
  onApplyColumnFilter: (column: number, rule: ColumnFilterRule | null) => void
  onClearAllFilters: () => void
}): React.ReactElement {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement | null>(null)
  const existing = sheet.columnFilters?.[String(column)]

  const [operator, setOperator] = useState<
    '' | NonNullable<ColumnFilterRule['operator']>
  >(existing?.operator ?? '')
  const [val1, setVal1] = useState(existing?.value ?? '')
  const [val2, setVal2] = useState(existing?.value2 ?? '')
  const [search, setSearch] = useState('')
  const [selectedSet, setSelectedSet] = useState<Set<string> | null>(
    existing?.selectedValues ? new Set(existing.selectedValues) : null,
  )

  useEffect(() => {
    if (!open) return
    const current = sheet.columnFilters?.[String(column)]
    setOperator(current?.operator ?? '')
    setVal1(current?.value ?? '')
    setVal2(current?.value2 ?? '')
    setSearch('')
    setSelectedSet(
      current?.selectedValues ? new Set(current.selectedValues) : null,
    )
  }, [open, column, sheet.columnFilters])

  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent): void => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) {
        setOpen(false)
      }
    }
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') setOpen(false)
    }
    document.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])

  const profile = useMemo(
    () => (open ? profileColumn(sheet, column, sheet.filterHasHeader !== false, workbook) : null),
    [open, sheet, column, workbook],
  )

  const visibleValues = useMemo(() => {
    if (!profile) return []
    const q = search.trim().toLowerCase()
    const list = q
      ? profile.valueCounts.filter(v => v.label.toLowerCase().includes(q))
      : profile.valueCounts
    return list.slice(0, 120)
  }, [profile, search])

  const allLabels = useMemo(
    () => (profile ? profile.valueCounts.map(v => v.label) : []),
    [profile],
  )

  function isChecked(label: string): boolean {
    return selectedSet === null || selectedSet.has(label)
  }

  function toggleValue(label: string): void {
    setSelectedSet(prev => {
      const next = new Set(prev ?? allLabels)
      if (next.has(label)) next.delete(label)
      else next.add(label)
      return next.size === allLabels.length ? null : next
    })
  }

  function apply(): void {
    const rule: ColumnFilterRule = {}
    if (selectedSet !== null) rule.selectedValues = [...selectedSet]
    if (operator) {
      rule.operator = operator
      if (operator !== 'blank' && operator !== 'not-blank') {
        rule.value = val1
        if (operator === 'between') rule.value2 = val2
      }
    }
    const empty = !rule.selectedValues && !rule.operator
    onApplyColumnFilter(column, empty ? null : rule)
    setOpen(false)
  }

  const colLetter = columnLabel(column)

  return (
    <Wrap ref={wrapRef}>
      <Trigger
        type="button"
        title="Smart column filter & profile"
        aria-label="Smart column filter"
        aria-pressed={activeFilterCount > 0}
        aria-expanded={open}
        onClick={() => setOpen(o => !o)}
      >
        <PlatformIcon icon={Filter} size={14} strokeWidth={1.9} />
        <span>
          Filter{activeFilterCount > 0 ? ` (${activeFilterCount})` : ''}
        </span>
      </Trigger>

      {open && profile && (
        <Panel aria-label={`Filter column ${colLetter}`}>
          <PanelHead>
            <div>
              <Title>
                Col {colLetter} · {profile.header}
              </Title>
              <SubMeta>
                {profile.uniqueCount} unique · {profile.blankCount} blank ·{' '}
                {profile.nonBlankCount} filled
              </SubMeta>
            </div>
            <IconClose type="button" aria-label="Close filter" onClick={() => setOpen(false)}>
              <PlatformIcon icon={X} size={13} />
            </IconClose>
          </PanelHead>

          {profile.numericStats && (
            <ProfileBox aria-label="Column distribution">
              <SparkBars>
                {profile.numericStats.bins.map((count, i) => {
                  const peak = Math.max(1, ...profile.numericStats!.bins)
                  const pct = Math.max(8, Math.round((count / peak) * 100))
                  return <Bar key={i} $pct={pct} title={`Bin ${i + 1}: ${count}`} />
                })}
              </SparkBars>
              <StatsLine>
                <span>Min {compactNum(profile.numericStats.min)}</span>
                <span>Med {compactNum(profile.numericStats.median)}</span>
                <span>Max {compactNum(profile.numericStats.max)}</span>
              </StatsLine>
            </ProfileBox>
          )}

          <SectionLabel>Condition</SectionLabel>
          <OpGrid>
            {OPERATORS.map(op => (
              <OpChip
                key={op.id || 'none'}
                type="button"
                $active={operator === op.id}
                onClick={() => setOperator(op.id)}
              >
                {op.label}
              </OpChip>
            ))}
          </OpGrid>

          {operator && operator !== 'blank' && operator !== 'not-blank' && (
            <FieldRow>
              <SmallInput
                value={val1}
                onChange={e => setVal1(e.currentTarget.value)}
                placeholder="Value"
                aria-label="Filter condition value"
              />
              {operator === 'between' && (
                <SmallInput
                  value={val2}
                  onChange={e => setVal2(e.currentTarget.value)}
                  placeholder="and"
                  aria-label="Filter upper value"
                />
              )}
            </FieldRow>
          )}

          <ChecklistHeader>
            <SectionLabel>Values ({profile.valueCounts.length})</SectionLabel>
            <LinkRow>
              <SmallLink type="button" onClick={() => setSelectedSet(null)}>
                All
              </SmallLink>
              <SmallLink type="button" onClick={() => setSelectedSet(new Set())}>
                None
              </SmallLink>
            </LinkRow>
          </ChecklistHeader>

          <SmallInput
            value={search}
            onChange={e => setSearch(e.currentTarget.value)}
            placeholder="Search values…"
            aria-label="Search column values"
          />

          <Checklist>
            {visibleValues.map(item => (
              <CheckRow key={item.label}>
                <input
                  type="checkbox"
                  checked={isChecked(item.label)}
                  onChange={() => toggleValue(item.label)}
                />
                <ValText>{item.label}</ValText>
                <ValCount>{item.count}</ValCount>
              </CheckRow>
            ))}
          </Checklist>

          <FooterRow>
            <PrimaryBtn type="button" onClick={apply}>
              Apply
            </PrimaryBtn>
            <QuietBtn
              type="button"
              onClick={() => {
                onApplyColumnFilter(column, null)
                setOpen(false)
              }}
            >
              Clear col
            </QuietBtn>
            {activeFilterCount > 0 && (
              <QuietBtn
                type="button"
                onClick={() => {
                  onClearAllFilters()
                  setOpen(false)
                }}
              >
                Clear all
              </QuietBtn>
            )}
          </FooterRow>
        </Panel>
      )}
    </Wrap>
  )
}

const Wrap = styled.div`
  position: relative;
  display: inline-flex;
`

const Trigger = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 5px;
  height: var(--pure-chrome-control-height, 28px);
  padding: 0 8px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius, 3px);
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
  &[aria-pressed='true'] {
    background: var(--p-chosen, var(--pure-chrome-selection));
    border-color: var(--sheets-accent);
  }
`

const Panel = styled.div`
  position: absolute;
  top: calc(100% + 6px);
  right: 0;
  z-index: 65;
  display: grid;
  gap: 8px;
  width: 310px;
  padding: 10px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius, 3px);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
`

const PanelHead = styled.div`
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 8px;
`

const Title = styled.div`
  font-size: 12.5px;
  font-weight: 700;
  color: var(--sheets-ink);
`

const SubMeta = styled.div`
  font-family: var(--sheets-mono);
  font-size: 10.5px;
  color: var(--pure-chrome-muted);
`

const IconClose = styled.button`
  border: 0;
  background: transparent;
  color: var(--pure-chrome-soft);
  cursor: pointer;
  padding: 2px;
`

const ProfileBox = styled.div`
  display: grid;
  gap: 4px;
  padding: 6px 8px;
  border: 1px solid var(--pure-chrome-fence, var(--pure-chrome-line));
  background: var(--n-panel, var(--pure-chrome-well));
`

const SparkBars = styled.div`
  display: grid;
  grid-template-columns: repeat(8, 1fr);
  align-items: end;
  gap: 3px;
  height: 28px;
`

const Bar = styled.div<{ $pct: number }>`
  height: ${({ $pct }) => `${$pct}%`};
  background: var(--sheets-accent);
  opacity: 0.78;
`

const StatsLine = styled.div`
  display: flex;
  justify-content: space-between;
  font-family: var(--sheets-mono);
  font-size: 10.5px;
  color: var(--pure-chrome-soft);
`

const SectionLabel = styled.div`
  font-size: 10.5px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: var(--pure-chrome-muted);
`

const OpGrid = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 3px;
`

const OpChip = styled.button<{ $active?: boolean }>`
  height: 22px;
  padding: 0 6px;
  border: 1px solid
    ${({ $active }) => ($active ? 'var(--sheets-accent)' : 'var(--pure-chrome-line)')};
  border-radius: 2px;
  background: ${({ $active }) =>
    $active ? 'var(--p-chosen, var(--pure-chrome-selection))' : 'var(--pure-chrome-surface)'};
  color: var(--sheets-ink);
  font-size: 11px;
  cursor: pointer;
`

const FieldRow = styled.div`
  display: flex;
  gap: 6px;
`

const SmallInput = styled.input`
  width: 100%;
  box-sizing: border-box;
  height: 26px;
  padding: 0 7px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-size: 12px;
`

const ChecklistHeader = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
`

const LinkRow = styled.div`
  display: flex;
  gap: 8px;
`

const SmallLink = styled.button`
  border: 0;
  padding: 0;
  background: transparent;
  color: var(--sheets-accent);
  font-size: 11px;
  font-weight: 600;
  cursor: pointer;
  &:hover {
    text-decoration: underline;
  }
`

const Checklist = styled.div`
  max-height: 148px;
  overflow-y: auto;
  border: 1px solid var(--pure-chrome-fence, var(--pure-chrome-line));
  padding: 4px;
  display: grid;
  gap: 2px;
`

const CheckRow = styled.label`
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 2px 4px;
  font-size: 12px;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const ValText = styled.span`
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--sheets-ink);
`

const ValCount = styled.span`
  flex: 0 0 auto;
  font-family: var(--sheets-mono);
  font-size: 10.5px;
  color: var(--pure-chrome-muted);
`

const FooterRow = styled.div`
  display: flex;
  gap: 6px;
`

const PrimaryBtn = styled.button`
  height: 26px;
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
  height: 26px;
  padding: 0 8px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-size: 12px;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
`
