import { useEffect, useMemo, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { ArrowUpDown, Plus, Sparkles, Trash2 } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import type { WorkbookSheet } from '../types'
import type { SortKey } from '../lib/dataTransforms'
import {
  cellKey,
  columnLabel,
  displayCellValue,
  normalizeRangeToken,
  parseCellKey,
} from '../lib/workbookModel'

function columnsInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): Array<{ index: number; label: string }> {
  const norm = normalizeRangeToken(rangeToken)
  if (!norm) return [{ index: 0, label: 'Col A' }]
  const [a, b] = norm.split(':')
  const s = parseCellKey(a)
  const e = parseCellKey(b ?? a)
  if (!s || !e) return [{ index: 0, label: 'Col A' }]
  const startCol = Math.min(s.column, e.column)
  const endCol = Math.max(s.column, e.column)
  const headerRow = Math.min(s.row, e.row)
  const cols: Array<{ index: number; label: string }> = []
  for (let c = startCol; c <= endCol; c += 1) {
    const header = displayCellValue(sheet, cellKey(headerRow, c)).trim()
    cols.push({
      index: c,
      label: header ? `${columnLabel(c)} (${header})` : `Col ${columnLabel(c)}`,
    })
  }
  return cols
}

export function MultiSortControl({
  sheet,
  selectedRange,
  onApplySort,
}: {
  sheet: WorkbookSheet
  selectedRange: string
  onApplySort: (keys: SortKey[], hasHeader: boolean) => void
}): React.ReactElement {
  const [open, setOpen] = useState(false)
  const [hasHeader, setHasHeader] = useState(true)
  const wrapRef = useRef<HTMLDivElement | null>(null)
  const cols = useMemo(
    () => columnsInRange(sheet, selectedRange),
    [sheet, selectedRange],
  )
  const [keys, setKeys] = useState<SortKey[]>([
    { column: cols[0]?.index ?? 0, direction: 'asc' },
  ])

  useEffect(() => {
    if (open) {
      setKeys([{ column: cols[0]?.index ?? 0, direction: 'asc' }])
    }
  }, [open, cols])

  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent): void => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', onDown)
    return () => document.removeEventListener('mousedown', onDown)
  }, [open])

  return (
    <Wrap ref={wrapRef}>
      <Trigger
        type="button"
        title="Multi-level sort"
        aria-label="Multi-level sort"
        aria-expanded={open}
        onClick={() => setOpen(o => !o)}
      >
        <PlatformIcon icon={ArrowUpDown} size={14} strokeWidth={1.9} />
        <span>Sort</span>
      </Trigger>

      {open && (
        <Panel aria-label="Multi-level sort">
          <Title>Sort {selectedRange}</Title>
          <CheckLabel>
            <input
              type="checkbox"
              checked={hasHeader}
              onChange={e => setHasHeader(e.currentTarget.checked)}
            />
            First row is a header
          </CheckLabel>
          <LevelStack>
            {keys.map((k, idx) => (
              <LevelRow key={idx}>
                <LevelTag>{idx === 0 ? 'By' : 'Then'}</LevelTag>
                <ColPills>
                  {cols.slice(0, 6).map(c => (
                    <SmallChip
                      key={c.index}
                      type="button"
                      $active={k.column === c.index}
                      onClick={() =>
                        setKeys(prev =>
                          prev.map((item, i) =>
                            i === idx ? { ...item, column: c.index } : item,
                          ),
                        )
                      }
                    >
                      {c.label}
                    </SmallChip>
                  ))}
                </ColPills>
                <DirBtn
                  type="button"
                  onClick={() =>
                    setKeys(prev =>
                      prev.map((item, i) =>
                        i === idx
                          ? {
                              ...item,
                              direction: item.direction === 'asc' ? 'desc' : 'asc',
                            }
                          : item,
                      ),
                    )
                  }
                >
                  {k.direction === 'asc' ? 'A→Z' : 'Z→A'}
                </DirBtn>
                {keys.length > 1 && (
                  <IconBtn
                    type="button"
                    aria-label="Remove sort level"
                    onClick={() =>
                      setKeys(prev => prev.filter((_, i) => i !== idx))
                    }
                  >
                    <PlatformIcon icon={Trash2} size={12} />
                  </IconBtn>
                )}
              </LevelRow>
            ))}
          </LevelStack>
          <Footer>
            {keys.length < 3 && cols.length > 1 && (
              <QuietBtn
                type="button"
                onClick={() =>
                  setKeys(prev => [
                    ...prev,
                    {
                      column:
                        cols[Math.min(prev.length, cols.length - 1)]?.index ?? 0,
                      direction: 'asc',
                    },
                  ])
                }
              >
                <PlatformIcon icon={Plus} size={12} /> Add level
              </QuietBtn>
            )}
            <PrimaryBtn
              type="button"
              onClick={() => {
                onApplySort(keys, hasHeader)
                setOpen(false)
              }}
            >
              Sort now
            </PrimaryBtn>
          </Footer>
        </Panel>
      )}
    </Wrap>
  )
}

export function DataCleanupControl({
  selectedRange,
  onRemoveDuplicates,
  onSplitTextToColumns,
  onTrimWhitespace,
  onToggleCheckboxes,
  onSetHyperlink,
}: {
  selectedRange: string
  onRemoveDuplicates: (hasHeader: boolean) => void
  onSplitTextToColumns: (delimiter: string) => void
  onTrimWhitespace: () => void
  onToggleCheckboxes: () => void
  onSetHyperlink: (url: string, label?: string) => void
}): React.ReactElement {
  const [open, setOpen] = useState(false)
  const [delimiter, setDelimiter] = useState(',')
  const [dedupeHeader, setDedupeHeader] = useState(true)
  const [linkUrl, setLinkUrl] = useState('')
  const [linkLabel, setLinkLabel] = useState('')
  const wrapRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent): void => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', onDown)
    return () => document.removeEventListener('mousedown', onDown)
  }, [open])

  return (
    <Wrap ref={wrapRef}>
      <Trigger
        type="button"
        title="Data cleanup & smart cells"
        aria-label="Data cleanup"
        aria-expanded={open}
        onClick={() => setOpen(o => !o)}
      >
        <PlatformIcon icon={Sparkles} size={14} strokeWidth={1.9} />
        <span>Clean</span>
      </Trigger>

      {open && (
        <Panel aria-label="Data cleanup and transforms">
          <Title>Transform {selectedRange}</Title>

          <Block>
            <SectionTitle>Remove duplicates</SectionTitle>
            <CheckLabel>
              <input
                type="checkbox"
                checked={dedupeHeader}
                onChange={e => setDedupeHeader(e.currentTarget.checked)}
              />
              Keep header row
            </CheckLabel>
            <QuietBtn
              type="button"
              onClick={() => {
                onRemoveDuplicates(dedupeHeader)
                setOpen(false)
              }}
            >
              Remove duplicate rows
            </QuietBtn>
          </Block>

          <Block>
            <SectionTitle>Split text to columns</SectionTitle>
            <InlineRow>
              <SmallInput
                value={delimiter}
                onChange={e => setDelimiter(e.currentTarget.value)}
                placeholder=","
                aria-label="Split delimiter"
              />
              <QuietBtn
                type="button"
                onClick={() => {
                  if (!delimiter) return
                  onSplitTextToColumns(delimiter)
                  setOpen(false)
                }}
              >
                Split
              </QuietBtn>
            </InlineRow>
          </Block>

          <Block>
            <SectionTitle>Quick tools</SectionTitle>
            <InlineRow>
              <QuietBtn
                type="button"
                onClick={() => {
                  onTrimWhitespace()
                  setOpen(false)
                }}
              >
                Trim whitespace
              </QuietBtn>
              <QuietBtn
                type="button"
                onClick={() => {
                  onToggleCheckboxes()
                  setOpen(false)
                }}
              >
                Toggle checkboxes
              </QuietBtn>
            </InlineRow>
          </Block>

          <Block>
            <SectionTitle>Cell hyperlink</SectionTitle>
            <SmallInput
              value={linkUrl}
              onChange={e => setLinkUrl(e.currentTarget.value)}
              placeholder="https://example.com"
              aria-label="Hyperlink URL"
            />
            <InlineRow>
              <SmallInput
                value={linkLabel}
                onChange={e => setLinkLabel(e.currentTarget.value)}
                placeholder="Display text (optional)"
                aria-label="Hyperlink label"
              />
              <PrimaryBtn
                type="button"
                onClick={() => {
                  onSetHyperlink(linkUrl, linkLabel || undefined)
                  setLinkUrl('')
                  setLinkLabel('')
                  setOpen(false)
                }}
              >
                Link
              </PrimaryBtn>
            </InlineRow>
          </Block>
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
`

const Panel = styled.div`
  position: absolute;
  top: calc(100% + 6px);
  right: 0;
  z-index: 65;
  display: grid;
  gap: 10px;
  width: 300px;
  padding: 10px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius, 3px);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
`

const Title = styled.div`
  font-size: 12.5px;
  font-weight: 700;
  color: var(--sheets-ink);
`

const SectionTitle = styled.div`
  font-size: 10.5px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: var(--pure-chrome-muted);
`

const Block = styled.div`
  display: grid;
  gap: 6px;
  padding-top: 6px;
  border-top: 1px solid var(--pure-chrome-fence, var(--pure-chrome-line));
`

const CheckLabel = styled.label`
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: var(--pure-chrome-soft);
  cursor: pointer;
`

const LevelStack = styled.div`
  display: grid;
  gap: 6px;
`

const LevelRow = styled.div`
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 5px;
  padding: 6px;
  border: 1px solid var(--pure-chrome-fence, var(--pure-chrome-line));
`

const LevelTag = styled.span`
  font-size: 11px;
  font-weight: 700;
  color: var(--pure-chrome-muted);
  min-width: 28px;
`

const ColPills = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 3px;
  flex: 1 1 auto;
`

const SmallChip = styled.button<{ $active?: boolean }>`
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

const DirBtn = styled.button`
  height: 22px;
  padding: 0 6px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-well);
  color: var(--sheets-ink);
  font-family: var(--sheets-mono);
  font-size: 10.5px;
  font-weight: 700;
  cursor: pointer;
`

const IconBtn = styled.button`
  border: 0;
  background: transparent;
  color: var(--pure-chrome-soft);
  cursor: pointer;
`

const InlineRow = styled.div`
  display: flex;
  gap: 6px;
`

const SmallInput = styled.input`
  flex: 1 1 auto;
  min-width: 0;
  height: 26px;
  padding: 0 7px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-size: 12px;
`

const Footer = styled.div`
  display: flex;
  justify-content: space-between;
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
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 4px;
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
