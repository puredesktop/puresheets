import { useEffect, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import {
  AlertTriangle,
  ChevronLeft,
  ChevronRight,
  Copy,
  EyeOff,
  List,
  Pencil,
  Plus,
  Trash2,
} from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import { IconButton, ToolIconButton } from './controls'
import type { WorkbookSheet } from '../types'
import { sheetAgentRefs } from '../agentRefs'

const TAB_COLORS = [
  { hex: '', label: 'None' },
  { hex: '#2f6ea2', label: 'Blue' },
  { hex: '#377d5a', label: 'Green' },
  { hex: '#b06b1f', label: 'Amber' },
  { hex: '#a8385a', label: 'Red' },
  { hex: '#7b52ab', label: 'Violet' },
]

const ZOOM_LEVELS = [0.75, 0.9, 1, 1.15, 1.25, 1.5]

export interface StatPill {
  label: string
  display: string
  raw: string
}

/**
 * Bottom bar: sheet tabs (add, select, rename inline, color dot, reorder,
 * right-click menu, sheet switcher, two-step delete) + interactive status pills & zoom.
 */
export function SheetTabsBar({
  compact = false,
  sheets,
  hiddenSheets = [],
  activeSheetId,
  renamingSheetId,
  renameDraft,
  statusText,
  statPills = [],
  invalidCount = 0,
  zoom = 1,
  onCopyStat,
  onOpenInvalidAudit,
  onZoomChange,
  onAddSheet,
  onSelectSheet,
  onRenameSheet,
  onBeginRenameSheet,
  onDuplicateSheet,
  onDeleteSheet,
  onRenameDraftChange,
  onCommitSheetRename,
  onCancelSheetRename,
  onMoveSheet,
  onSetSheetTabColor,
  onHideSheet,
  onUnhideSheet,
}: {
  compact?: boolean
  sheets: WorkbookSheet[]
  hiddenSheets?: WorkbookSheet[]
  activeSheetId: string
  renamingSheetId: string | null
  renameDraft: string
  statusText: string
  statPills?: StatPill[]
  invalidCount?: number
  zoom?: number
  onCopyStat?: (pill: StatPill) => void
  onOpenInvalidAudit?: () => void
  onZoomChange?: (ratio: number) => void
  onAddSheet: () => void
  onSelectSheet: (sheetId: string) => void
  onRenameSheet: () => void
  onBeginRenameSheet: (sheetId: string) => void
  onDuplicateSheet: () => void
  onDeleteSheet: () => void
  onRenameDraftChange: (value: string) => void
  onCommitSheetRename: () => void
  onCancelSheetRename: () => void
  onMoveSheet?: (sheetId: string, delta: -1 | 1) => void
  onSetSheetTabColor?: (sheetId: string, color: string | undefined) => void
  onHideSheet?: (sheetId: string) => void
  onUnhideSheet?: (sheetId: string) => void
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
  const [contextMenu, setContextMenu] = useState<{
    sheetId: string
    x: number
    y: number
  } | null>(null)
  const [switcherOpen, setSwitcherOpen] = useState(false)

  useEffect(() => {
    if (!contextMenu && !switcherOpen) return
    const close = (): void => {
      setContextMenu(null)
      setSwitcherOpen(false)
    }
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') close()
    }
    document.addEventListener('mousedown', close)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', close)
      document.removeEventListener('keydown', onKey)
    }
  }, [contextMenu, switcherOpen])
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

        {(sheets.length > 3 || hiddenSheets.length > 0) && !compact && (
          <SwitcherWrap onMouseDown={e => e.stopPropagation()}>
            <IconButton
              type="button"
              title="All sheets"
              aria-label="All sheets"
              onClick={() => setSwitcherOpen(o => !o)}
            >
              <PlatformIcon icon={List} size={14} />
            </IconButton>
            {switcherOpen && (
              <SheetMenu role="menu" aria-label="All worksheets">
                {sheets.map(s => (
                  <SheetMenuItem
                    key={s.id}
                    type="button"
                    $active={s.id === activeSheetId}
                    onClick={() => {
                      onSelectSheet(s.id)
                      setSwitcherOpen(false)
                    }}
                  >
                    {s.tabColor && <ColorDot $color={s.tabColor} />}
                    <span>{s.name}</span>
                  </SheetMenuItem>
                ))}
                {hiddenSheets.length > 0 && onUnhideSheet && (
                  <>
                    <MenuDivider />
                    <MenuHeading>Hidden sheets (click to show)</MenuHeading>
                    {hiddenSheets.map(s => (
                      <SheetMenuItem
                        key={s.id}
                        type="button"
                        onClick={() => {
                          onUnhideSheet(s.id)
                          setSwitcherOpen(false)
                        }}
                      >
                        <PlatformIcon icon={EyeOff} size={11} />
                        <span>{s.name}</span>
                      </SheetMenuItem>
                    ))}
                  </>
                )}
              </SheetMenu>
            )}
          </SwitcherWrap>
        )}

        {sheets.map(sheet => (
          <SheetTab
            key={sheet.id}
            as={renamingSheetId === sheet.id ? 'div' : 'button'}
            type={renamingSheetId === sheet.id ? undefined : 'button'}
            {...sheetAgentRefs(sheet)}
            $active={sheet.id === activeSheetId}
            $tabColor={sheet.tabColor}
            aria-current={sheet.id === activeSheetId ? 'true' : undefined}
            title={`${sheet.name} — double-click to rename, right-click for options`}
            onClick={() => onSelectSheet(sheet.id)}
            onDoubleClick={() => onBeginRenameSheet(sheet.id)}
            onContextMenu={(event: React.MouseEvent) => {
              if (compact) return
              event.preventDefault()
              onSelectSheet(sheet.id)
              setContextMenu({
                sheetId: sheet.id,
                x: Math.min(event.clientX, window.innerWidth - 200),
                y: Math.max(8, event.clientY - 220),
              })
            }}
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
              <>
                {sheet.pivot && <PivotTag>P</PivotTag>}
                {sheet.name}
              </>
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
                  onClick={() => onMoveSheet(activeSheetId, -1)}
                  icon={ChevronLeft}
                />
                <ToolIconButton
                  title="Move sheet right"
                  disabled={!canMoveRight}
                  onClick={() => onMoveSheet(activeSheetId, 1)}
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

      <RightCluster>
        {!compact && statPills.length > 0 && (
          <StatStrip aria-label="Selection statistics">
            {statPills.map(pill => (
              <StatChip
                key={pill.label}
                type="button"
                title={`Click to copy ${pill.label} (${pill.raw})`}
                onClick={() => onCopyStat?.(pill)}
              >
                <span>{pill.label}</span>
                <strong>{pill.display}</strong>
              </StatChip>
            ))}
          </StatStrip>
        )}

        {!compact && invalidCount > 0 && onOpenInvalidAudit && (
          <WarnChip
            type="button"
            title={`${invalidCount} validation issues — click to inspect`}
            onClick={onOpenInvalidAudit}
          >
            <PlatformIcon icon={AlertTriangle} size={11} />
            {invalidCount}
          </WarnChip>
        )}

        <Status $compact={compact} title={statusText}>
          {statusText}
        </Status>

        {!compact && onZoomChange && (
          <ZoomBtn
            type="button"
            title="Cycle sheet zoom"
            aria-label="Zoom level"
            onClick={() => {
              const idx = ZOOM_LEVELS.findIndex(z => Math.abs(z - zoom) < 0.01)
              const next = ZOOM_LEVELS[(idx + 1) % ZOOM_LEVELS.length]
              onZoomChange(next)
            }}
          >
            {Math.round(zoom * 100)}%
          </ZoomBtn>
        )}
      </RightCluster>

      {contextMenu && (
        <TabContextMenu
          style={{ left: contextMenu.x, top: contextMenu.y }}
          onMouseDown={e => e.stopPropagation()}
        >
          <SheetMenuItem
            type="button"
            onClick={() => {
              onBeginRenameSheet(contextMenu.sheetId)
              setContextMenu(null)
            }}
          >
            Rename
          </SheetMenuItem>
          <SheetMenuItem
            type="button"
            onClick={() => {
              onDuplicateSheet()
              setContextMenu(null)
            }}
          >
            Duplicate
          </SheetMenuItem>
          {onMoveSheet && (
            <>
              <SheetMenuItem
                type="button"
                onClick={() => {
                  onMoveSheet(contextMenu.sheetId, -1)
                  setContextMenu(null)
                }}
              >
                Move left
              </SheetMenuItem>
              <SheetMenuItem
                type="button"
                onClick={() => {
                  onMoveSheet(contextMenu.sheetId, 1)
                  setContextMenu(null)
                }}
              >
                Move right
              </SheetMenuItem>
            </>
          )}
          {onHideSheet && sheets.length > 1 && (
            <SheetMenuItem
              type="button"
              onClick={() => {
                onHideSheet(contextMenu.sheetId)
                setContextMenu(null)
              }}
            >
              Hide sheet
            </SheetMenuItem>
          )}
          {onSetSheetTabColor && (
            <>
              <MenuDivider />
              <ColorRow>
                {TAB_COLORS.map(c => (
                  <ColorSwatch
                    key={c.label}
                    type="button"
                    title={`Tab color: ${c.label}`}
                    $color={c.hex || 'transparent'}
                    onClick={() => {
                      onSetSheetTabColor(
                        contextMenu.sheetId,
                        c.hex || undefined,
                      )
                      setContextMenu(null)
                    }}
                  />
                ))}
              </ColorRow>
            </>
          )}
          <MenuDivider />
          <SheetMenuItem
            type="button"
            onClick={() => {
              onDeleteSheet()
              setContextMenu(null)
            }}
          >
            Delete
          </SheetMenuItem>
        </TabContextMenu>
      )}
    </BottomBar>
  )
}

const BottomBar = styled.footer<{ $compact?: boolean }>`
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 12px;
  align-items: center;
  border-top: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-bar);
  padding: 6px var(--pure-chrome-inset);

  @media (max-width: 720px) {
    grid-template-columns: ${({ $compact }) =>
      $compact ? 'minmax(0, 1fr) auto' : 'minmax(0, 1fr)'};
    gap: ${({ $compact }) => ($compact ? '12px' : '4px')};
  }
`

const SheetTabs = styled.div`
  display: flex;
  gap: 5px;
  min-width: 0;
  overflow: auto;
  align-items: center;
`

const SwitcherWrap = styled.div`
  position: relative;
  display: inline-flex;
`

const SheetTab = styled.button<{ $active?: boolean; $tabColor?: string }>`
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 5px;
  min-width: 72px;
  flex-shrink: 0;
  height: 30px;
  border: 0;
  border-bottom: ${({ $tabColor }) =>
    $tabColor ? `2px solid ${$tabColor}` : '2px solid transparent'};
  border-radius: 0;
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
    background: ${({ $active }) =>
      $active ? 'var(--p-chosen)' : 'var(--pure-chrome-hover)'};
  }
  &:focus-visible {
    outline: 2px solid var(--sheets-accent);
    outline-offset: -2px;
  }
`

const PivotTag = styled.span`
  display: inline-grid;
  place-items: center;
  width: 14px;
  height: 14px;
  border-radius: 2px;
  background: var(--sheets-accent);
  color: var(--pure-chrome-on-accent);
  font-family: var(--sheets-mono);
  font-size: 9px;
  font-weight: 700;
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

const RightCluster = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
`

const StatStrip = styled.div`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  @media (max-width: 980px) {
    display: none;
  }
`

const StatChip = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  height: 22px;
  padding: 0 6px;
  border: 1px solid var(--pure-chrome-fence, var(--pure-chrome-line));
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--pure-chrome-soft);
  font-family: var(--sheets-mono);
  font-size: 10.5px;
  cursor: pointer;
  strong {
    color: var(--sheets-ink);
  }
  &:hover {
    border-color: var(--pure-chrome-line);
    background: var(--pure-chrome-hover);
  }
`

const WarnChip = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 3px;
  height: 22px;
  padding: 0 6px;
  border: 1px solid var(--platform-colors-danger-text);
  border-radius: 2px;
  background: color-mix(in srgb, var(--platform-colors-danger-text) 10%, transparent);
  color: var(--platform-colors-danger-text);
  font-family: var(--sheets-mono);
  font-size: 10.5px;
  font-weight: 700;
  cursor: pointer;
`

const ZoomBtn = styled.button`
  height: 22px;
  padding: 0 6px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--pure-chrome-soft);
  font-family: var(--sheets-mono);
  font-size: 10.5px;
  font-weight: 600;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
    color: var(--sheets-ink);
  }
`

/* Family status-line convention: meta, 11px mono in the muted colour. */
const Status = styled(MetaText)<{ $compact?: boolean }>`
  min-width: 0;
  max-width: min(36vw, 420px);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;

  @media (max-width: 720px) {
    max-width: ${({ $compact }) => ($compact ? '38vw' : 'none')};
  }
`

const SheetMenu = styled.div`
  position: absolute;
  bottom: calc(100% + 6px);
  left: 0;
  z-index: 80;
  width: 196px;
  max-height: 260px;
  overflow-y: auto;
  padding: 4px;
  border: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
  display: grid;
  gap: 2px;
`

const TabContextMenu = styled.div`
  position: fixed;
  z-index: 100;
  width: 176px;
  padding: 4px;
  border: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
  display: grid;
  gap: 2px;
`

const SheetMenuItem = styled.button<{ $active?: boolean }>`
  display: flex;
  align-items: center;
  gap: 6px;
  width: 100%;
  padding: 5px 8px;
  border: 0;
  background: ${({ $active }) =>
    $active ? 'var(--p-chosen, var(--pure-chrome-selection))' : 'transparent'};
  color: var(--sheets-ink);
  font-size: 12px;
  text-align: left;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const ColorDot = styled.span<{ $color: string }>`
  width: 8px;
  height: 8px;
  border-radius: 1px;
  background: ${({ $color }) => $color};
`

const MenuDivider = styled.div`
  height: 1px;
  margin: 3px 0;
  background: var(--pure-chrome-fence, var(--pure-chrome-line));
`

const MenuHeading = styled.div`
  padding: 3px 8px;
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  color: var(--pure-chrome-muted);
`

const ColorRow = styled.div`
  display: flex;
  gap: 5px;
  padding: 4px 8px;
`

const ColorSwatch = styled.button<{ $color: string }>`
  width: 18px;
  height: 18px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: ${({ $color }) => $color};
  cursor: pointer;
`
