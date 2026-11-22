import { useEffect, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { ChevronDown, Grid2X2, PaintBucket, PanelTop, RotateCcw, Type } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import type { CellStyle } from '../types'

const TEXT_SWATCHES = [
  '#14171a', '#474c54', '#666c75', '#8d939b', '#ffffff',
  '#982d25', '#b06b1f', '#256c43', '#1d6fa5', '#6a4496',
  '#c9382e', '#d18619', '#377d5a', '#2b82c9', '#8a5aa8',
]

const FILL_SWATCHES = [
  '#ffffff', '#f6f7f8', '#eef0f2', '#e2e5e8', '#d3d7dc',
  '#fdecea', '#fff4bf', '#e5f3df', '#e4f0fb', '#f1eaf8',
  '#f8d2cd', '#ffe399', '#c7e6be', '#c2dff7', '#dfcff2',
]

const BORDER_COLORS = [
  '#c8ccd2', '#8d939b', '#14171a', '#1d6fa5', '#256c43', '#982d25',
]

function useOutsideClose(
  open: boolean,
  ref: React.RefObject<HTMLElement | null>,
  onClose: () => void,
): void {
  useEffect(() => {
    if (!open) return
    const onPointerDown = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose()
    }
    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') onClose()
    }
    document.addEventListener('mousedown', onPointerDown)
    document.addEventListener('keydown', onKeyDown)
    return () => {
      document.removeEventListener('mousedown', onPointerDown)
      document.removeEventListener('keydown', onKeyDown)
    }
  }, [open, ref, onClose])
}

/**
 * Clean swatch color picker with custom color input + one-click Reset / No fill.
 * Renders a `<label title="...">` trigger so existing accessibility and test
 * selectors (`label[title="Text color"]`, `label[title="Fill color"]`) stay intact.
 */
export function ColorSwatchControl({
  mode,
  value,
  onChange,
}: {
  mode: 'text' | 'fill'
  value: string | undefined
  onChange: (color: string | undefined) => void
}): React.ReactElement {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement | null>(null)
  useOutsideClose(open, wrapRef, () => setOpen(false))

  const title = mode === 'text' ? 'Text color' : 'Fill color'
  const swatches = mode === 'text' ? TEXT_SWATCHES : FILL_SWATCHES
  const indicator =
    value ?? (mode === 'text' ? '#14171a' : 'transparent')

  return (
    <Wrap ref={wrapRef}>
      <ColorLabel
        title={title}
        onClick={event => {
          event.preventDefault()
          setOpen(o => !o)
        }}
      >
        <PlatformIcon
          icon={mode === 'text' ? Type : PaintBucket}
          size={15}
          strokeWidth={1.9}
        />
        <ColorBar $color={indicator} />
      </ColorLabel>
      {open && (
        <Popover aria-label={`${title} palette`}>
          <PopoverHead>
            <span>{title}</span>
            <ResetBtn
              type="button"
              onClick={() => {
                onChange(undefined)
                setOpen(false)
              }}
            >
              <PlatformIcon icon={RotateCcw} size={11} />
              {mode === 'fill' ? 'No fill' : 'Default'}
            </ResetBtn>
          </PopoverHead>
          <SwatchGrid>
            {swatches.map(hex => (
              <Swatch
                key={hex}
                type="button"
                $color={hex}
                $selected={value?.toLowerCase() === hex.toLowerCase()}
                title={hex}
                aria-label={`${title} ${hex}`}
                onClick={() => {
                  onChange(hex)
                  setOpen(false)
                }}
              />
            ))}
          </SwatchGrid>
          <CustomRow>
            <CustomLabel>Custom</CustomLabel>
            <CustomInput
              type="color"
              aria-label={`Custom ${title.toLowerCase()}`}
              value={value ?? (mode === 'text' ? '#20231f' : '#ffffff')}
              onChange={e => onChange(e.currentTarget.value)}
            />
          </CustomRow>
        </Popover>
      )}
    </Wrap>
  )
}

export type BorderPreset =
  | 'all'
  | 'outer'
  | 'bottom'
  | 'top'
  | 'top-bottom'
  | 'double-bottom'
  | 'left'
  | 'right'
  | 'none'

/**
 * Per-side border picker (All, Bottom, Double-bottom accounting underline,
 * Top+Bottom, Outer, Left, Right, Clear) + border color.
 */
export function BorderPickerControl({
  hasBorder,
  onToggleAllBorders,
  onApplyBorder,
}: {
  hasBorder: boolean
  onToggleAllBorders: () => void
  onApplyBorder: (preset: BorderPreset, color?: string) => void
}): React.ReactElement {
  const [open, setOpen] = useState(false)
  const [borderColor, setBorderColor] = useState<string>('#c8ccd2')
  const wrapRef = useRef<HTMLDivElement | null>(null)
  useOutsideClose(open, wrapRef, () => setOpen(false))

  const presets: Array<{ id: BorderPreset; label: string }> = [
    { id: 'all', label: 'All borders' },
    { id: 'bottom', label: 'Bottom' },
    { id: 'top-bottom', label: 'Top + bottom' },
    { id: 'double-bottom', label: 'Double bottom' },
    { id: 'top', label: 'Top' },
    { id: 'left', label: 'Left' },
    { id: 'right', label: 'Right' },
    { id: 'none', label: 'Clear borders' },
  ]

  return (
    <Wrap ref={wrapRef}>
      <TriggerBtn
        type="button"
        title="Toggle all borders"
        aria-label="Toggle all borders"
        aria-pressed={hasBorder}
        onClick={onToggleAllBorders}
      >
        <PlatformIcon icon={Grid2X2} size={15} strokeWidth={1.9} />
      </TriggerBtn>
      <CaretBtn
        type="button"
        title="Border styles"
        aria-label="Border styles"
        aria-expanded={open}
        onClick={() => setOpen(o => !o)}
      >
        <PlatformIcon icon={ChevronDown} size={11} />
      </CaretBtn>
      {open && (
        <Popover aria-label="Border options" $width={220}>
          <PopoverHead>
            <span>Borders</span>
          </PopoverHead>
          <PresetGrid>
            {presets.map(p => (
              <PresetBtn
                key={p.id}
                type="button"
                onClick={() => {
                  onApplyBorder(p.id, borderColor)
                  setOpen(false)
                }}
              >
                {p.label}
              </PresetBtn>
            ))}
          </PresetGrid>
          <BorderColorRow>
            <CustomLabel>Color</CustomLabel>
            <MiniSwatches>
              {BORDER_COLORS.map(c => (
                <Swatch
                  key={c}
                  type="button"
                  $color={c}
                  $selected={borderColor === c}
                  aria-label={`Border color ${c}`}
                  onClick={() => setBorderColor(c)}
                />
              ))}
            </MiniSwatches>
          </BorderColorRow>
        </Popover>
      )}
    </Wrap>
  )
}

/**
 * Freeze panes popover: Freeze Row 1, Freeze Col A, Freeze up to current cell, Unfreeze.
 */
export function FreezePanesControl({
  frozenRows,
  frozenColumns,
  selectedCell,
  selectedRow,
  selectedCol,
  onSetFreeze,
}: {
  frozenRows: number
  frozenColumns: number
  selectedCell: string
  selectedRow: number
  selectedCol: number
  onSetFreeze: (rows: number, columns: number) => void
}): React.ReactElement {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement | null>(null)
  useOutsideClose(open, wrapRef, () => setOpen(false))

  const active = frozenRows > 0 || frozenColumns > 0

  return (
    <Wrap ref={wrapRef}>
      <TriggerBtn
        type="button"
        title="Freeze panes"
        aria-label="Freeze panes"
        aria-pressed={active}
        aria-expanded={open}
        onClick={() => setOpen(o => !o)}
      >
        <PlatformIcon icon={PanelTop} size={15} strokeWidth={1.9} />
      </TriggerBtn>
      {open && (
        <Popover aria-label="Freeze panes menu" $width={224}>
          <PopoverHead>
            <span>Freeze panes</span>
          </PopoverHead>
          <MenuStack>
            <MenuItem
              type="button"
              onClick={() => {
                onSetFreeze(frozenRows === 1 ? 0 : 1, frozenColumns)
                setOpen(false)
              }}
            >
              {frozenRows === 1 ? 'Unfreeze top row' : 'Freeze top row (Row 1)'}
            </MenuItem>
            <MenuItem
              type="button"
              onClick={() => {
                onSetFreeze(frozenRows, frozenColumns === 1 ? 0 : 1)
                setOpen(false)
              }}
            >
              {frozenColumns === 1
                ? 'Unfreeze first column'
                : 'Freeze first column (Col A)'}
            </MenuItem>
            <MenuItem
              type="button"
              onClick={() => {
                onSetFreeze(selectedRow + 1, selectedCol + 1)
                setOpen(false)
              }}
            >
              Freeze up to {selectedCell} ({selectedRow + 1}R × {selectedCol + 1}C)
            </MenuItem>
            {active && (
              <MenuItem
                type="button"
                onClick={() => {
                  onSetFreeze(0, 0)
                  setOpen(false)
                }}
              >
                Unfreeze all panes
              </MenuItem>
            )}
          </MenuStack>
        </Popover>
      )}
    </Wrap>
  )
}

/**
 * Custom number format & currency symbol popover triggered from Format toolbar.
 */
export function CustomNumberFormatPopover({
  open,
  currentStyle,
  onApply,
  onClose,
}: {
  open: boolean
  currentStyle: CellStyle | undefined
  onApply: (patch: Partial<CellStyle>, message: string) => void
  onClose: () => void
}): React.ReactElement | null {
  const [pattern, setPattern] = useState(currentStyle?.numberPattern ?? '#,##0.00')
  const wrapRef = useRef<HTMLDivElement | null>(null)
  useOutsideClose(open, wrapRef, onClose)

  useEffect(() => {
    if (open) setPattern(currentStyle?.numberPattern ?? '#,##0.00')
  }, [open, currentStyle?.numberPattern])

  if (!open) return null

  const currencies: Array<{ sym: '$' | '€' | '£' | '¥'; label: string }> = [
    { sym: '$', label: 'USD ($)' },
    { sym: '€', label: 'EUR (€)' },
    { sym: '£', label: 'GBP (£)' },
    { sym: '¥', label: 'JPY (¥)' },
  ]

  const patterns = [
    { code: '#,##0.00', label: '1,234.56' },
    { code: '#,##0;(#,##0);"-"', label: 'Accounting' },
    { code: '0.0%', label: '12.5%' },
    { code: 'yyyy-mm-dd', label: '2026-11-22' },
    { code: 'mmm d, yyyy', label: 'Nov 22, 2026' },
    { code: 'hh:mm:ss', label: '14:30:00' },
    { code: '0.00E+00', label: '1.23E+03' },
  ]

  return (
    <Popover ref={wrapRef} aria-label="Custom number format" $width={260}>
      <PopoverHead>
        <span>Currency & custom format</span>
      </PopoverHead>
      <PresetGrid>
        {currencies.map(c => (
          <PresetBtn
            key={c.sym}
            type="button"
            onClick={() => {
              onApply(
                { numberFormat: 'currency', currencySymbol: c.sym },
                `Formatted as ${c.label}`,
              )
              onClose()
            }}
          >
            {c.label}
          </PresetBtn>
        ))}
      </PresetGrid>
      <SectionSub>Quick patterns</SectionSub>
      <PresetGrid>
        {patterns.map(p => (
          <PresetBtn
            key={p.code}
            type="button"
            onClick={() => {
              onApply({ numberPattern: p.code }, `Applied pattern ${p.code}`)
              onClose()
            }}
          >
            {p.label}
          </PresetBtn>
        ))}
      </PresetGrid>
      <SectionSub>Custom Excel format code</SectionSub>
      <CustomCodeRow>
        <CodeInput
          value={pattern}
          onChange={e => setPattern(e.currentTarget.value)}
          placeholder="#,##0.00;[Red]-#,##0.00"
          aria-label="Custom format code"
        />
        <ApplySmallBtn
          type="button"
          onClick={() => {
            if (!pattern.trim()) return
            onApply(
              { numberPattern: pattern.trim() },
              `Applied format ${pattern.trim()}`,
            )
            onClose()
          }}
        >
          Apply
        </ApplySmallBtn>
      </CustomCodeRow>
    </Popover>
  )
}

const Wrap = styled.div`
  position: relative;
  display: inline-flex;
`

const ColorLabel = styled.label`
  position: relative;
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  min-width: var(--pure-chrome-control-height, 28px);
  height: var(--pure-chrome-control-height, 28px);
  border-radius: var(--pure-chrome-radius, 3px);
  color: var(--sheets-ink);
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
  &[title]:hover::after,
  &:focus-within::after {
    position: absolute;
    top: calc(100% + 6px);
    left: 50%;
    z-index: 50;
    max-width: 220px;
    padding: 5px 7px;
    border: 1px solid var(--platform-colors-border);
    border-radius: 0;
    background: var(--platform-colors-popover);
    box-shadow: var(--platform-shadow-md);
    color: var(--platform-colors-popover-text);
    content: attr(title);
    font-size: var(--pure-chrome-meta-size);
    font-weight: 600;
    line-height: 1.2;
    pointer-events: none;
    text-align: center;
    transform: translateX(-50%);
    white-space: nowrap;
  }
`

const ColorBar = styled.span<{ $color: string }>`
  width: 14px;
  height: 3px;
  margin-top: 1px;
  border: 1px solid var(--pure-chrome-line);
  background: ${({ $color }) => $color};
`

const TriggerBtn = styled.button`
  display: inline-grid;
  place-items: center;
  min-width: var(--pure-chrome-control-height, 28px);
  height: var(--pure-chrome-control-height, 28px);
  border: 0;
  border-radius: var(--pure-chrome-radius, 3px);
  background: transparent;
  color: var(--sheets-ink);
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
  &[aria-pressed='true'] {
    background: var(--pure-chrome-selection);
  }
`

const CaretBtn = styled(TriggerBtn)`
  min-width: 16px;
  padding: 0 2px;
  color: var(--pure-chrome-soft);
`

const Popover = styled.div<{ $width?: number }>`
  position: absolute;
  top: calc(100% + 6px);
  left: 0;
  z-index: 60;
  display: grid;
  gap: 8px;
  width: ${({ $width }) => ($width ? `${$width}px` : '196px')};
  padding: 10px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius, 3px);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
`

const PopoverHead = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  color: var(--sheets-ink);
  font-size: 11.5px;
  font-weight: 700;
`

const ResetBtn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 2px 6px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--pure-chrome-soft);
  font-size: 11px;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
    color: var(--sheets-ink);
  }
`

const SwatchGrid = styled.div`
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 5px;
`

const MiniSwatches = styled.div`
  display: flex;
  gap: 5px;
`

const Swatch = styled.button<{ $color: string; $selected?: boolean }>`
  width: 22px;
  height: 22px;
  border: 1px solid
    ${({ $selected }) => ($selected ? 'var(--sheets-accent)' : 'var(--pure-chrome-line)')};
  outline: ${({ $selected }) => ($selected ? '2px solid var(--sheets-accent)' : 'none')};
  border-radius: 2px;
  background: ${({ $color }) => $color};
  cursor: pointer;
`

const CustomRow = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding-top: 6px;
  border-top: 1px solid var(--pure-chrome-fence, var(--pure-chrome-line));
`

const BorderColorRow = styled(CustomRow)`
  gap: 8px;
`

const CustomLabel = styled.span`
  color: var(--pure-chrome-soft);
  font-size: 11px;
  font-weight: 600;
`

const CustomInput = styled.input`
  width: 34px;
  height: 22px;
  padding: 0;
  border: 1px solid var(--pure-chrome-line);
  background: transparent;
  cursor: pointer;
`

const PresetGrid = styled.div`
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 4px;
`

const PresetBtn = styled.button`
  height: 26px;
  padding: 0 7px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-size: 11.5px;
  text-align: left;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const MenuStack = styled.div`
  display: grid;
  gap: 3px;
`

const MenuItem = styled.button`
  height: 28px;
  padding: 0 8px;
  border: 0;
  border-radius: 2px;
  background: transparent;
  color: var(--sheets-ink);
  font-size: 12px;
  text-align: left;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const SectionSub = styled.div`
  color: var(--pure-chrome-muted);
  font-size: 10.5px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.04em;
`

const CustomCodeRow = styled.div`
  display: flex;
  gap: 6px;
`

const CodeInput = styled.input`
  flex: 1 1 auto;
  min-width: 0;
  height: 26px;
  padding: 0 7px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-family: var(--sheets-mono);
  font-size: 11.5px;
`

const ApplySmallBtn = styled.button`
  height: 26px;
  padding: 0 10px;
  border: 1px solid var(--pure-chrome-accent);
  border-radius: 2px;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-size: 11.5px;
  font-weight: 600;
  cursor: pointer;
`
