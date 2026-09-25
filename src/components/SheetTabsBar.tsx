import type React from 'react'
import styled from 'styled-components'
import { Copy, Pencil, Plus, Trash2 } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import { IconButton, ToolIconButton } from './controls'
import type { WorkbookSheet } from '../types'

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
                aria-label="Sheet name"
                onChange={event =>
                  onRenameDraftChange(event.currentTarget.value)
                }
                onClick={event => event.stopPropagation()}
                onBlur={onCommitSheetRename}
                onKeyDown={event => {
                  if (event.key === 'Enter') {
                    onCommitSheetRename()
                  } else if (event.key === 'Escape') {
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
      <Status>{statusText}</Status>
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
  padding: 6px var(--pure-chrome-inset);
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
  /* The active tab is filled with the accent tint and gets an accent underline;
     inactive tabs stay quiet and muted so the current sheet is obvious. */
  height: var(--pure-chrome-control-height);
  border: 1px solid
    ${({ $active }) => ($active ? 'var(--sheets-accent)' : 'var(--pure-chrome-line)')};
  border-bottom-width: ${({ $active }) => ($active ? '3px' : '1px')};
  border-radius: 0;
  background: ${({ $active }) =>
    $active ? 'var(--sheets-accent-bg)' : 'var(--pure-chrome-surface)'};
  padding: 0 10px;
  color: ${({ $active }) =>
    $active ? 'var(--sheets-accent-text)' : 'var(--pure-chrome-soft)'};
  font-size: var(--pure-chrome-ui-size);
  font-weight: ${({ $active }) => ($active ? 600 : 500)};
  transition: background 160ms ease;
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
const Status = styled(MetaText)``
