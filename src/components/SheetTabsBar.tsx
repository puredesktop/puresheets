import { useEffect, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { ChevronLeft, ChevronRight, Copy, Pencil, Plus, Trash2 } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import { IconButton, ToolIconButton } from './controls'
import type { WorkbookSheet } from '../types'
import { sheetAgentRefs } from '../agentRefs'

/**
 * Bottom bar: sheet tabs (add, select, rename inline, reorder) plus sheet
 * actions and the status line. Legacy agent-draft sheets are hidden here on
 * purpose — they are not user sheets.
 */
export function SheetTabsBar({
  compact = false,
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
  onMoveSheet,
  onRenameDraftChange,
  onCommitSheetRename,
  onCancelSheetRename,
}: {
  /** Beside the Executive room: tabs and status on one line, no sheet actions. */
  compact?: boolean
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
  onMoveSheet?: (direction: 'left' | 'right') => void
  onRenameDraftChange: (value: string) => void
  onCommitSheetRename: () => void
  onCancelSheetRename: () => void
}): React.ReactElement {
  const renameHandled = useRef(false)
  const [confirmingDelete, setConfirmingDelete] = useState(false)
  useEffect(() => {
    setConfirmingDelete(false)
  }, [activeSheetId, sheets.length])
  const activeIndex = sheets.findIndex(sheet => sheet.id === activeSheetId)
  const activeSheet = sheets[activeIndex]
  const canDelete = sheets.length > 1
  const canMoveLeft = activeIndex > 0
  const canMoveRight = activeIndex >= 0 && activeIndex < sheets.length - 1
  return (
    <BottomBar $compact={compact}>
      <SheetTabs aria-label="Sheet tabs">
        <AddSheetButton
          type="button"
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
            type={renamingSheetId === sheet.id ? undefined : 'button'}
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
        {!compact && (
          <>
            {onMoveSheet && (
              <>
                <ToolIconButton
                  title="Move sheet left"
                  disabled={!canMoveLeft}
                  onClick={() => onMoveSheet('left')}
                  icon={ChevronLeft}
                />
                <ToolIconButton
                  title="Move sheet right"
                  disabled={!canMoveRight}
                  onClick={() => onMoveSheet('right')}
                  icon={ChevronRight}
                />
              </>
            )}
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
            {confirmingDelete && canDelete ? (
              <DeleteConfirm
                role="group"
                aria-label="Confirm delete sheet"
                onKeyDown={event => {
                  if (event.key === 'Escape') {
                    event.preventDefault()
                    setConfirmingDelete(false)
                  }
                }}
              >
                <DeleteConfirmLabel>
                  Delete “{activeSheet?.name ?? 'sheet'}”?
                </DeleteConfirmLabel>
                <ConfirmDangerButton
                  type="button"
                  onClick={() => {
                    setConfirmingDelete(false)
                    onDeleteSheet()
                  }}
                >
                  Confirm
                </ConfirmDangerButton>
                <ConfirmCancelButton
                  type="button"
                  onClick={() => setConfirmingDelete(false)}
                >
                  Cancel
                </ConfirmCancelButton>
              </DeleteConfirm>
            ) : (
              <ToolIconButton
                title="Delete sheet"
                disabled={!canDelete}
                onClick={() => {
                  if (canDelete) setConfirmingDelete(true)
                }}
                icon={Trash2}
              />
            )}
          </>
        )}
      </SheetTabs>
      <Status $compact={compact} title={statusText}>
        {statusText}
      </Status>
    </BottomBar>
  )
}

const BottomBar = styled.footer<{ $compact?: boolean }>`
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 16px;
  align-items: center;
  border-top: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-bar);
  padding: 8px var(--pure-chrome-inset);

  @media (max-width: 720px) {
    grid-template-columns: ${({ $compact }) =>
      $compact ? 'minmax(0, 1fr) auto' : 'minmax(0, 1fr)'};
    gap: ${({ $compact }) => ($compact ? '16px' : '4px')};
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
  border: 0;
  border-radius: 0;
  /* The open sheet has the selected look every app shares: a hairline in Sheets' deep colour around a faint ground, words ink. */
  background: ${({ $active }) => ($active ? 'var(--p-chosen)' : 'transparent')};
  box-shadow: ${({ $active }) => ($active ? 'var(--p-chosen-edge)' : 'none')};
  padding: 0 10px;
  color: ${({ $active }) =>
    $active ? 'var(--sheets-ink)' : 'var(--pure-chrome-soft)'};
  font-size: var(--pure-chrome-ui-size);
  font-weight: ${({ $active }) => ($active ? 600 : 500)};
  white-space: nowrap;
  cursor: pointer;
  &:hover {
    background: ${({ $active }) => ($active ? 'var(--p-chosen)' : 'var(--pure-chrome-hover)')};
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

const DeleteConfirm = styled.div`
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 0 6px;
  height: 30px;
  border: 1px solid var(--platform-colors-danger-border, #e3a29b);
  background: var(--pure-chrome-surface);
  white-space: nowrap;
  flex-shrink: 0;
`

const DeleteConfirmLabel = styled.span`
  color: var(--sheets-ink);
  font-size: 12px;
  font-weight: 500;
`

const ConfirmDangerButton = styled.button`
  height: 22px;
  padding: 0 8px;
  border: 0;
  background: var(--platform-colors-danger-text, #b3261e);
  color: #ffffff;
  font-size: 11px;
  font-weight: 600;
  cursor: pointer;
`

const ConfirmCancelButton = styled.button`
  height: 22px;
  padding: 0 8px;
  border: 1px solid var(--pure-chrome-line);
  background: transparent;
  color: var(--sheets-ink);
  font-size: 11px;
  font-weight: 500;
  cursor: pointer;
`

/* Family status-line convention: meta, 11px mono in the muted colour. */
const Status = styled(MetaText)<{ $compact?: boolean }>`
  min-width: 0;
  max-width: min(45vw, 520px);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;

  @media (max-width: 720px) {
    max-width: ${({ $compact }) => ($compact ? '38vw' : 'none')};
  }
`
