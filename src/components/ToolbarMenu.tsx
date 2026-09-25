import { useCallback, useEffect, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { ChevronDown } from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { MenuDropdown } from '@purescience/platform-ui/components/common/dropdown/MenuDropdown'
import { isMenuInteractionTarget } from '@purescience/platform-ui/components/common/dropdown/menuPortalConstants'
import type { MenuEntry } from '@purescience/platform-ui/components/common/dropdown/menuTypes'

export interface ToolbarMenuOption {
  value: string
  label: string
}

/**
 * Compact toolbar dropdown built on the shared platform menu (custom menu,
 * not a native `<select>`, per the enterprise design rules). The trigger is
 * the platform's 28px outlined toolbar select; the popover is the shared
 * `MenuDropdown` panel.
 */
export function ToolbarMenu({
  label,
  displayValue,
  options,
  selectedValue,
  icon,
  minWidth,
  compact = false,
  onSelect,
}: {
  /** Accessible name and tooltip for the trigger, e.g. "Number format". */
  label: string
  /** Text shown inside the trigger, e.g. the current value. */
  displayValue: string
  options: ToolbarMenuOption[]
  selectedValue?: string
  icon?: LucideIcon
  minWidth?: number
  /** Icon + caret only; the label stays accessible via title/aria. */
  compact?: boolean
  onSelect: (value: string) => void
}): React.ReactElement {
  const triggerRef = useRef<HTMLButtonElement>(null)
  const [open, setOpen] = useState(false)
  const close = useCallback(() => setOpen(false), [])

  useEffect(() => {
    if (!open) return
    const onPointerDown = (event: MouseEvent): void => {
      if (isMenuInteractionTarget(event.target)) return
      if (triggerRef.current?.contains(event.target as Node)) return
      close()
    }
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') close()
    }
    document.addEventListener('mousedown', onPointerDown)
    document.addEventListener('keydown', onKeyDown)
    return () => {
      document.removeEventListener('mousedown', onPointerDown)
      document.removeEventListener('keydown', onKeyDown)
    }
  }, [close, open])

  const items: MenuEntry[] = options.map(option => ({
    id: option.value,
    label:
      option.value === selectedValue ? `${option.label} ✓` : option.label,
  }))

  return (
    <TriggerWrap onClick={event => event.stopPropagation()}>
      <Trigger
        ref={triggerRef}
        type="button"
        title={label}
        aria-label={label}
        aria-haspopup="menu"
        aria-expanded={open}
        data-platform-menu-trigger
        onClick={() => setOpen(current => !current)}
      >
        {icon ? (
          <PlatformIcon icon={icon} size={15} strokeWidth={1.9} />
        ) : null}
        {compact && icon ? null : (
          <TriggerValue>{displayValue}</TriggerValue>
        )}
        <PlatformIcon icon={ChevronDown} size={12} strokeWidth={1.9} />
      </Trigger>
      <MenuDropdown
        sectionId={`sheets-toolbar-${label}`}
        items={items}
        open={open}
        triggerRef={triggerRef}
        placement="below-start"
        minWidth={minWidth ?? 120}
        onSelectItem={item => onSelect(item.id)}
        onClose={close}
      />
    </TriggerWrap>
  )
}

const TriggerWrap = styled.div`
  position: relative;
  display: inline-flex;
`

/** Typed loosely on purpose: styled-components' attrs rejects data-* literals. */
const selectChrome: Record<string, string> = { 'data-chrome': 'toolbar-select' }

const Trigger = styled.button.attrs(selectChrome)`
  position: relative;
  gap: 6px;
  cursor: pointer;

  &:hover:not(:disabled) {
    background: var(--pure-chrome-hover);
  }

  &:focus-visible {
    outline: 2px solid var(--pure-chrome-accent);
    outline-offset: 1px;
  }

  &[aria-label]:hover::after,
  &[aria-label]:focus-visible::after {
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
    content: attr(aria-label);
    font-size: var(--pure-chrome-meta-size);
    font-weight: 600;
    line-height: 1.2;
    pointer-events: none;
    text-align: center;
    transform: translateX(-50%);
    white-space: nowrap;
  }
`

const TriggerValue = styled.span`
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`
