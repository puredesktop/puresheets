import { useRef } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { Copy, Pencil, Plus, Trash2 } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import { IconButton, ToolIconButton } from './controls'
import type { WorkbookSheet } from '../types'
import { sheetAgentRefs } from '../agentRefs'

/**
 * Bottom bar: sheet tabs (add, select, rename inline) plus sheet actions and
 * the status line. Legacy agent-draft sheets are hidden here on purpose —
 * they are not user sheets.
 */
export function SheetTabsBar({
  sheets,
  activeSheetId,
  renamingSheetId,
  renameDraft,
  statusText,
  onAddSheet,
  onSelectSheet,
  onRenameSheet,
  onBeginRenameSheet,
  onDuplicateSheet,
  onDeleteSheet,
  onRenameDraftChange,
  onCommitSheetRename,
  onCancelSheetRename,
}: {
  sheets: WorkbookSheet[]
  activeSheetId: string
  renamingSheetId: string | null
  renameDraft: string
  statusText: string
  onAddSheet: () => void
  onSelectSheet: (sheetId: string) => void
  onRenameSheet: () => void
  /** Start renaming a specific tab (double-click on the tab). */
  onBeginRenameSheet: (sheetId: string) => void
  onDuplicateSheet: () => void
  onDeleteSheet: () => void
  onRenameDraftChange: (value: string) => void
  onCommitSheetRename: () => void
  onCancelSheetRename: () => void
}): React.ReactElement {
  const renameHandled = useRef(false)
  return (
    <BottomBar>
      <SheetTabs aria-label="Sheet tabs">
        <AddSheetButton
          title="Add sheet"
          aria-label="Add sheet"
          onClick={onAddSheet}
        >
          <PlatformIcon icon={Plus} size={15} strokeWidth={2} />
        </AddSheetButton>
        {sheets.map(sheet => (
          <SheetTab
            key={sheet.id}
            as={renamingSheetId === sheet.id ? 'div' : 'button'}
            {...sheetAgentRefs(sheet)}
            $active={sheet.id === activeSheetId}
            aria-current={sheet.id === activeSheetId ? 'true' : undefined}
            title={`${sheet.name} — double-click to rename`}
            onClick={() => onSelectSheet(sheet.id)}
            onDoubleClick={() => onBeginRenameSheet(sheet.id)}
          >
            {renamingSheetId === sheet.id ? (
              <SheetRenameInput
                value={renameDraft}
                autoFocus
                onFocus={() => {
                  renameHandled.current = false
                }}
                aria-label="Sheet name"
                onChange={event =>
                  onRenameDraftChange(event.currentTarget.value)
                }
                onClick={event => event.stopPropagation()}
                onBlur={() => {
                  if (!renameHandled.current) onCommitSheetRename()
                }}
                onKeyDown={event => {
                  if (event.key === 'Enter') {
                    event.preventDefault()
                    renameHandled.current = true
                    onCommitSheetRename()
                  } else if (event.key === 'Escape') {
                    event.preventDefault()
                    renameHandled.current = true
                    onCancelSheetRename()
                  }
                }}
              />
            ) : (
              sheet.name
            )}
          </SheetTab>
        ))}
        <ToolIconButton
          title="Rename sheet"
          onClick={onRenameSheet}
          icon={Pencil}
        />
        <ToolIconButton
          title="Duplicate sheet"
          onClick={onDuplicateSheet}
          icon={Copy}
        />
        <ToolIconButton
          title="Delete sheet"
          onClick={onDeleteSheet}
          icon={Trash2}
        />
      </SheetTabs>
      <Status title={statusText}>{statusText}</Status>
    </BottomBar>
  )
}

const BottomBar = styled.footer`
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 16px;
  align-items: center;
  border-top: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-bar);
  padding: 8px var(--pure-chrome-inset);

  @media (max-width: 720px) {
    grid-template-columns: minmax(0, 1fr);
    gap: 4px;
  }
`

const SheetTabs = styled.div`
  display: flex;
  gap: 6px;
  min-width: 0;
  overflow: auto;
`

const SheetTab = styled.button<{ $active?: boolean }>`
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 72px;
  flex-shrink: 0;
  height: 30px;
  border: 1px solid
    ${({ $active }) => ($active ? 'var(--pure-chrome-line)' : 'transparent')};
  border-radius: 7px;
  background: ${({ $active }) =>
    $active ? 'var(--sheets-accent-bg)' : 'transparent'};
  padding: 0 10px;
  color: ${({ $active }) =>
    $active ? 'var(--sheets-accent-text)' : 'var(--pure-chrome-soft)'};
  font-size: var(--pure-chrome-ui-size);
  font-weight: ${({ $active }) => ($active ? 600 : 500)};
  white-space: nowrap;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
  &:focus-visible {
    outline: 2px solid var(--sheets-accent);
    outline-offset: -2px;
  }
`

const SheetRenameInput = styled.input`
  width: 104px;
  min-width: 0;
  border: 0;
  background: transparent;
  color: inherit;
  font: inherit;
  text-align: center;
  outline: none;
`

const AddSheetButton = styled(IconButton)`
  flex: 0 0 auto;
`

/* Family status-line convention: meta, 11px mono in the muted colour. */
const Status = styled(MetaText)`
  min-width: 0;
  max-width: min(45vw, 520px);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;

  @media (max-width: 720px) {
    max-width: none;
  }
`
