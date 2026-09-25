import { useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { PaintBucket, ShieldCheck } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { IconButton } from './controls'
import type { CellStyle, ConditionalFormatRule } from '../types'

/**
 * Compact popover controls for the Phase S3 data rules: data validation and
 * conditional formatting. Quiet enterprise styling — small option buttons,
 * no native selects, panels anchored under their toolbar trigger.
 */

export type ValidationKind = 'list' | 'number-range' | 'date'

export function ValidationControl({
  selectedRangeLabel,
  onApply,
  onClear,
}: {
  selectedRangeLabel: string
  onApply: (kind: ValidationKind, options: { min?: number; max?: number }) => void
  onClear: () => void
}): React.ReactElement {
  const [open, setOpen] = useState(false)
  const [kind, setKind] = useState<ValidationKind>('list')
  const [minText, setMinText] = useState('')
  const [maxText, setMaxText] = useState('')

  function apply(): void {
    const min = minText.trim() === '' ? undefined : Number(minText)
    const max = maxText.trim() === '' ? undefined : Number(maxText)
    onApply(kind, {
      min: Number.isFinite(min as number) ? min : undefined,
      max: Number.isFinite(max as number) ? max : undefined,
    })
    setOpen(false)
  }

  return (
    <ControlWrap>
      <IconButton
        type="button"
        title="Data validation"
        aria-label="Data validation"
        aria-expanded={open}
        onClick={() => setOpen(current => !current)}
      >
        <PlatformIcon icon={ShieldCheck} size={15} strokeWidth={1.9} />
      </IconButton>
      {open && (
        <Panel aria-label="Data validation rule">
          <PanelTitle>Validate {selectedRangeLabel}</PanelTitle>
          <OptionRow role="radiogroup" aria-label="Validation kind">
            <OptionButton
              type="button"
              role="radio"
              aria-checked={kind === 'list'}
              $active={kind === 'list'}
              onClick={() => setKind('list')}
            >
              List from values
            </OptionButton>
            <OptionButton
              type="button"
              role="radio"
              aria-checked={kind === 'number-range'}
              $active={kind === 'number-range'}
              onClick={() => setKind('number-range')}
            >
              Number range
            </OptionButton>
            <OptionButton
              type="button"
              role="radio"
              aria-checked={kind === 'date'}
              $active={kind === 'date'}
              onClick={() => setKind('date')}
            >
              Date
            </OptionButton>
          </OptionRow>
          {kind === 'list' && (
            <PanelNote>
              Allowed values are collected from the selection's current
              entries.
            </PanelNote>
          )}
          {kind === 'number-range' && (
            <FieldRow>
              <SmallInput
                value={minText}
                onChange={event => setMinText(event.currentTarget.value)}
                placeholder="Min"
                aria-label="Minimum value"
                inputMode="decimal"
              />
              <SmallInput
                value={maxText}
                onChange={event => setMaxText(event.currentTarget.value)}
                placeholder="Max"
                aria-label="Maximum value"
                inputMode="decimal"
              />
            </FieldRow>
          )}
          {kind === 'date' && (
            <PanelNote>Entries must parse as calendar dates.</PanelNote>
          )}
          <ActionRow>
            <ApplyButton type="button" onClick={apply}>
              Apply rule
            </ApplyButton>
            <QuietButton
              type="button"
              onClick={() => {
                onClear()
                setOpen(false)
              }}
            >
              Clear rules
            </QuietButton>
            <QuietButton type="button" onClick={() => setOpen(false)}>
              Close
            </QuietButton>
          </ActionRow>
        </Panel>
      )}
    </ControlWrap>
  )
}

export type ConditionKind = ConditionalFormatRule['condition']['kind']

const STYLE_PRESETS: Array<{ id: string; label: string; style: CellStyle }> = [
  { id: 'red-text', label: 'Red text', style: { textColor: '#982d25', bold: true } },
  { id: 'green-fill', label: 'Green fill', style: { fillColor: '#e5f3df' } },
  { id: 'amber-fill', label: 'Amber fill', style: { fillColor: '#fff4bf' } },
]

export function ConditionalFormatControl({
  selectedRangeLabel,
  onApply,
  onClear,
}: {
  selectedRangeLabel: string
  onApply: (
    condition: ConditionalFormatRule['condition'],
    style: CellStyle,
  ) => void
  onClear: () => void
}): React.ReactElement {
  const [open, setOpen] = useState(false)
  const [kind, setKind] = useState<ConditionKind>('greater')
  const [value, setValue] = useState('')
  const [value2, setValue2] = useState('')
  const [presetId, setPresetId] = useState(STYLE_PRESETS[0].id)

  const kindOptions: Array<{ id: ConditionKind; label: string }> = [
    { id: 'greater', label: 'Greater than' },
    { id: 'less', label: 'Less than' },
    { id: 'equal', label: 'Equal to' },
    { id: 'between', label: 'Between' },
    { id: 'text-contains', label: 'Text contains' },
  ]

  function apply(): void {
    if (!value.trim()) return
    const preset =
      STYLE_PRESETS.find(item => item.id === presetId) ?? STYLE_PRESETS[0]
    onApply(
      {
        kind,
        value: value.trim(),
        value2: kind === 'between' ? value2.trim() || undefined : undefined,
      },
      preset.style,
    )
    setOpen(false)
  }

  return (
    <ControlWrap>
      <IconButton
        type="button"
        title="Conditional formatting"
        aria-label="Conditional formatting"
        aria-expanded={open}
        onClick={() => setOpen(current => !current)}
      >
        <PlatformIcon icon={PaintBucket} size={15} strokeWidth={1.9} />
      </IconButton>
      {open && (
        <Panel aria-label="Conditional formatting rule">
          <PanelTitle>Format {selectedRangeLabel} when…</PanelTitle>
          <OptionRow role="radiogroup" aria-label="Condition">
            {kindOptions.map(option => (
              <OptionButton
                key={option.id}
                type="button"
                role="radio"
                aria-checked={kind === option.id}
                $active={kind === option.id}
                onClick={() => setKind(option.id)}
              >
                {option.label}
              </OptionButton>
            ))}
          </OptionRow>
          <FieldRow>
            <SmallInput
              value={value}
              onChange={event => setValue(event.currentTarget.value)}
              placeholder={kind === 'text-contains' ? 'Text' : 'Value'}
              aria-label="Condition value"
            />
            {kind === 'between' && (
              <SmallInput
                value={value2}
                onChange={event => setValue2(event.currentTarget.value)}
                placeholder="and"
                aria-label="Condition upper value"
              />
            )}
          </FieldRow>
          <OptionRow role="radiogroup" aria-label="Applied style">
            {STYLE_PRESETS.map(preset => (
              <OptionButton
                key={preset.id}
                type="button"
                role="radio"
                aria-checked={presetId === preset.id}
                $active={presetId === preset.id}
                onClick={() => setPresetId(preset.id)}
              >
                {preset.label}
              </OptionButton>
            ))}
          </OptionRow>
          <ActionRow>
            <ApplyButton type="button" onClick={apply}>
              Apply rule
            </ApplyButton>
            <QuietButton
              type="button"
              onClick={() => {
                onClear()
                setOpen(false)
              }}
            >
              Clear rules
            </QuietButton>
            <QuietButton type="button" onClick={() => setOpen(false)}>
              Close
            </QuietButton>
          </ActionRow>
        </Panel>
      )}
    </ControlWrap>
  )
}

const ControlWrap = styled.div`
  position: relative;
  display: inline-flex;
`

const Panel = styled.div`
  position: absolute;
  top: calc(100% + 6px);
  right: 0;
  z-index: 24;
  display: grid;
  gap: 8px;
  width: 300px;
  padding: 10px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
`

const PanelTitle = styled.div`
  color: var(--platform-colors-text);
  font-size: var(--pure-chrome-ui-size);
  font-weight: 600;
`

const PanelNote = styled.div`
  color: var(--pure-chrome-soft);
  font-size: var(--pure-chrome-ui-size);
`

const OptionRow = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
`

const OptionButton = styled.button<{ $active?: boolean }>`
  height: var(--pure-chrome-control-height);
  padding: 0 8px;
  border: 1px solid
    ${({ $active }) => ($active ? 'var(--pure-chrome-soft)' : 'var(--pure-chrome-line)')};
  border-radius: 7px;
  background: ${({ $active }) =>
    $active ? 'var(--pure-chrome-selection)' : 'var(--pure-chrome-surface)'};
  color: var(--platform-colors-text);
  font-size: 12px;
  font-weight: ${({ $active }) => ($active ? 600 : 500)};
`

const FieldRow = styled.div`
  display: flex;
  gap: 6px;
`

/** Typed loosely on purpose: styled-components' attrs rejects data-* literals. */
const fieldChrome: Record<string, string> = { 'data-chrome': 'field' }

const SmallInput = styled.input.attrs(fieldChrome)`
  width: 0;
  flex: 1 1 0;
`

const ActionRow = styled.div`
  display: flex;
  gap: 6px;
`

const ApplyButton = styled.button`
  height: var(--pure-chrome-control-height);
  padding: 0 10px;
  border: 1px solid var(--pure-chrome-accent);
  border-radius: 7px;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-size: 12px;
  font-weight: 600;
`

const QuietButton = styled.button`
  height: var(--pure-chrome-control-height);
  padding: 0 10px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 7px;
  background: var(--pure-chrome-surface);
  color: var(--platform-colors-text);
  font-size: 12px;
  font-weight: 500;
`
