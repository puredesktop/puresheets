import type React from 'react'
import type { ButtonHTMLAttributes } from 'react'
import type { LucideIcon } from 'lucide-react'
import styled from 'styled-components'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'

/**
 * Shared compact controls for the PureSheets chrome (command bar, sheet tabs,
 * agent panel). They are the platform's 28px toolbar controls
 * (data-chrome="toolbar-control" / "toolbar-divider") with CSS tooltips
 * driven by the accessible label.
 */

/** Typed loosely on purpose: styled-components' attrs rejects data-* literals. */
const chrome = (kind: string): Record<string, string> => ({ 'data-chrome': kind })

export const IconButton = styled.button.attrs(chrome('toolbar-control'))`
  position: relative;

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

export const ToolbarDivider = styled.span.attrs(chrome('toolbar-divider'))`
  align-self: center;
`

export const ToolTextButton = styled.button.attrs(chrome('toolbar-control'))`
  padding: 0 9px;
`

export function ToolIconButton({
  icon,
  title,
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement> & {
  icon: LucideIcon
  title: string
}): React.ReactElement {
  return (
    <IconButton title={title} aria-label={title} {...props}>
      <PlatformIcon icon={icon} size={15} strokeWidth={1.9} />
    </IconButton>
  )
}
