import { useEffect, useMemo, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { Search, X } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'

export interface PaletteCommand {
  id: string
  label: string
  category: string
  shortcut?: string
  keywords?: string
  disabled?: boolean
  run: () => void
}

/**
 * Cmd/Ctrl+K command palette — gives power users instant fuzzy access to
 * every formatting, data, transform, view, sheet, and audit command without
 * hunting through tabs or ribbons.
 */
export function CommandPalette({
  open,
  commands,
  onClose,
}: {
  open: boolean
  commands: PaletteCommand[]
  onClose: () => void
}): React.ReactElement | null {
  const [query, setQuery] = useState('')
  const [activeIndex, setActiveIndex] = useState(0)
  const inputRef = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    if (open) {
      setQuery('')
      setActiveIndex(0)
      requestAnimationFrame(() => inputRef.current?.focus())
    }
  }, [open])

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    const list = commands.filter(cmd => !cmd.disabled)
    if (!q) return list.slice(0, 40)
    const terms = q.split(/\s+/).filter(Boolean)
    return list
      .filter(cmd => {
        const hay = `${cmd.label} ${cmd.category} ${cmd.keywords ?? ''}`.toLowerCase()
        return terms.every(term => hay.includes(term))
      })
      .slice(0, 40)
  }, [commands, query])

  if (!open) return null

  const clampedIndex = Math.min(activeIndex, Math.max(0, filtered.length - 1))

  function execute(cmd: PaletteCommand | undefined): void {
    if (!cmd) return
    onClose()
    cmd.run()
  }

  return (
    <Backdrop
      role="dialog"
      aria-modal="true"
      aria-label="Command palette"
      onMouseDown={event => {
        if (event.target === event.currentTarget) onClose()
      }}
    >
      <PaletteCard>
        <SearchRow>
          <PlatformIcon icon={Search} size={15} />
          <SearchInput
            ref={inputRef}
            value={query}
            onChange={event => {
              setQuery(event.currentTarget.value)
              setActiveIndex(0)
            }}
            onKeyDown={event => {
              if (event.key === 'Escape') {
                event.preventDefault()
                onClose()
              } else if (event.key === 'ArrowDown') {
                event.preventDefault()
                setActiveIndex(i =>
                  filtered.length ? (i + 1) % filtered.length : 0,
                )
              } else if (event.key === 'ArrowUp') {
                event.preventDefault()
                setActiveIndex(i =>
                  filtered.length
                    ? (i - 1 + filtered.length) % filtered.length
                    : 0,
                )
              } else if (event.key === 'Enter') {
                event.preventDefault()
                execute(filtered[clampedIndex])
              }
            }}
            placeholder="Type a command (merge, pivot, dedupe, freeze, currency, sheet…)"
            aria-label="Search commands"
          />
          <CloseBtn type="button" aria-label="Close command palette" onClick={onClose}>
            <PlatformIcon icon={X} size={14} />
          </CloseBtn>
        </SearchRow>

        <List role="listbox" aria-label="Matching commands">
          {filtered.length === 0 ? (
            <Empty>No matching commands</Empty>
          ) : (
            filtered.map((cmd, idx) => (
              <Row
                key={cmd.id}
                type="button"
                role="option"
                aria-selected={idx === clampedIndex}
                $active={idx === clampedIndex}
                onMouseEnter={() => setActiveIndex(idx)}
                onClick={() => execute(cmd)}
              >
                <RowLeft>
                  <Category>{cmd.category}</Category>
                  <Label>{cmd.label}</Label>
                </RowLeft>
                {cmd.shortcut ? <Shortcut>{cmd.shortcut}</Shortcut> : null}
              </Row>
            ))
          )}
        </List>
        <Footer>
          <span>↑↓ navigate</span>
          <span>↵ run</span>
          <span>esc close</span>
        </Footer>
      </PaletteCard>
    </Backdrop>
  )
}

const Backdrop = styled.div`
  position: fixed;
  inset: 0;
  z-index: 1500;
  display: flex;
  align-items: flex-start;
  justify-content: center;
  padding-top: 10vh;
  background: rgb(0 0 0 / 0.24);
`

const PaletteCard = styled.div`
  width: min(560px, calc(100vw - 32px));
  display: grid;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius, 3px);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-lg, 0 8px 24px rgb(0 0 0 / 0.18));
  overflow: hidden;
`

const SearchRow = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 12px;
  border-bottom: 1px solid var(--pure-chrome-line);
  color: var(--pure-chrome-soft);
`

const SearchInput = styled.input`
  flex: 1 1 auto;
  min-width: 0;
  border: 0;
  background: transparent;
  color: var(--sheets-ink);
  font: inherit;
  font-size: 13px;
  outline: none;
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
  &:hover {
    background: var(--pure-chrome-hover);
    color: var(--sheets-ink);
  }
`

const List = styled.div`
  max-height: 360px;
  overflow-y: auto;
  padding: 4px 0;
`

const Row = styled.button<{ $active?: boolean }>`
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 7px 12px;
  border: 0;
  background: ${({ $active }) =>
    $active ? 'var(--p-chosen, var(--pure-chrome-selection))' : 'transparent'};
  box-shadow: ${({ $active }) =>
    $active ? 'var(--p-chosen-edge, none)' : 'none'};
  color: var(--sheets-ink);
  font: inherit;
  font-size: 12.5px;
  text-align: left;
  cursor: pointer;
`

const RowLeft = styled.span`
  display: inline-flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
`

const Category = styled.span`
  flex: 0 0 68px;
  color: var(--pure-chrome-muted);
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.04em;
`

const Label = styled.span`
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-weight: 500;
`

const Shortcut = styled.kbd`
  flex: 0 0 auto;
  padding: 1px 5px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-well);
  color: var(--pure-chrome-soft);
  font-family: var(--sheets-mono);
  font-size: 10.5px;
`

const Empty = styled.div`
  padding: 20px 12px;
  color: var(--pure-chrome-muted);
  text-align: center;
  font-size: 12.5px;
`

const Footer = styled.div`
  display: flex;
  gap: 14px;
  padding: 6px 12px;
  border-top: 1px solid var(--pure-chrome-line);
  background: var(--n-panel, var(--pure-chrome-well));
  color: var(--pure-chrome-muted);
  font-family: var(--sheets-mono);
  font-size: 10.5px;
`
