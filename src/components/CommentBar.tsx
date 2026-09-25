import { useEffect, useMemo, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'

const CELL_KEY = /^([A-Za-z]+)(\d+)$/
// Reading order: top-to-bottom, then left-to-right, with numeric row compare so
// A2 sorts before A10.
function compareCells(a: string, b: string): number {
  const ma = CELL_KEY.exec(a)
  const mb = CELL_KEY.exec(b)
  if (!ma || !mb) return a.localeCompare(b)
  const rowA = Number.parseInt(ma[2], 10)
  const rowB = Number.parseInt(mb[2], 10)
  if (rowA !== rowB) return rowA - rowB
  return ma[1].localeCompare(mb[1])
}

/**
 * Slim inline editor for the selected cell's comment/note (Phase S3), plus a
 * roll-up list of every note on the sheet so they can be found and jumped to
 * without hunting cell by cell. Rendered under the formula bar when comments
 * are toggled on — no modal.
 */
export function CommentBar({
  selectedCell,
  comment,
  rangeLabel,
  comments = {},
  onSaveComment,
  onSelectCell,
  onClose,
}: {
  selectedCell: string
  comment: string
  /** When the selection spans more than one cell, its range token (e.g.
   *  "A1:C3") — saving then applies the comment to every cell in the range. */
  rangeLabel?: string
  /** Every comment on the active sheet, keyed by cell — powers the list. */
  comments?: Record<string, string>
  onSaveComment: (text: string) => void
  /** Jump the selection to a cell (clicking a row in the comments list). */
  onSelectCell?: (cell: string) => void
  onClose: () => void
}): React.ReactElement {
  const [draft, setDraft] = useState(comment)
  const [listOpen, setListOpen] = useState(false)

  // Track selection changes: editing a different cell starts from its comment.
  useEffect(() => {
    setDraft(comment)
  }, [selectedCell, comment])

  // Dirty while the field differs from what's saved; once saved (draft matches
  // the stored comment) the button flips to a disabled "Saved" so it's obvious
  // the comment landed.
  const dirty = draft !== comment

  const all = useMemo(
    () =>
      Object.entries(comments)
        .filter(([, text]) => text.trim().length > 0)
        .sort(([a], [b]) => compareCells(a, b)),
    [comments],
  )

  return (
    <Wrap>
      <Bar aria-label="Cell comment editor">
        <CellLabel>{rangeLabel ?? selectedCell}</CellLabel>
        <CommentInput
          value={draft}
          onChange={event => setDraft(event.currentTarget.value)}
          onKeyDown={event => {
            if (event.key === 'Enter') onSaveComment(draft)
            if (event.key === 'Escape') onClose()
          }}
          placeholder={
            rangeLabel
              ? `Add a comment for ${rangeLabel}`
              : 'Add a comment for this cell'
          }
          aria-label="Cell comment"
        />
        {rangeLabel ? (
          <RangeHint aria-live="polite">applies to {rangeLabel}</RangeHint>
        ) : null}
        <BarButton
          type="button"
          disabled={!dirty}
          onClick={() => onSaveComment(draft)}
        >
          {!dirty && comment ? 'Saved' : 'Save comment'}
        </BarButton>
        <BarButton
          type="button"
          disabled={!comment}
          onClick={() => onSaveComment('')}
        >
          Remove
        </BarButton>
        {all.length > 0 ? (
          <NotesToggle
            type="button"
            aria-expanded={listOpen}
            onClick={() => setListOpen(open => !open)}
          >
            <NotesCount>{all.length}</NotesCount>
            {all.length === 1 ? 'comment' : 'comments'}
            <Caret $open={listOpen}>▾</Caret>
          </NotesToggle>
        ) : null}
        <BarButton type="button" onClick={onClose}>
          Close
        </BarButton>
      </Bar>
      {listOpen && all.length > 0 ? (
        <NotesPanel role="listbox" aria-label="All comments on this sheet">
          {all.map(([cell, text]) => (
            <NoteRow
              key={cell}
              type="button"
              role="option"
              aria-selected={cell === selectedCell}
              $active={cell === selectedCell}
              onClick={() => onSelectCell?.(cell)}
            >
              <NoteCell>{cell}</NoteCell>
              <NotePreview>{text}</NotePreview>
            </NoteRow>
          ))}
        </NotesPanel>
      ) : null}
    </Wrap>
  )
}

/** Typed loosely on purpose: styled-components' attrs rejects data-* literals. */
const chrome = (kind: string): Record<string, string> => ({ 'data-chrome': kind })

const Wrap = styled.div`
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-well);
`

const Bar = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px var(--pure-chrome-inset);
`

const RangeHint = styled(MetaText)`
  font-weight: 600;
`

const CellLabel = styled(MetaText)`
  flex: 0 0 44px;
  color: var(--sheets-ink);
  font-weight: 700;
  text-align: center;
`

const CommentInput = styled.input.attrs(chrome('field'))`
  flex: 1 1 auto;
  min-width: 0;
  border-radius: 0;
`

const BarButton = styled.button.attrs(chrome('toolbar-select'))`
  flex: 0 0 auto;
  cursor: pointer;

  &:hover:not(:disabled) {
    background: var(--pure-chrome-hover);
  }

  &:disabled {
    opacity: 0.45;
  }
`

const NotesToggle = styled.button.attrs(chrome('toolbar-select'))`
  flex: 0 0 auto;
  gap: 5px;
  color: var(--pure-chrome-soft);
  cursor: pointer;

  &:hover {
    background: var(--pure-chrome-hover);
    color: var(--platform-colors-text);
  }
`

const NotesCount = styled.span`
  display: inline-grid;
  place-items: center;
  min-width: 16px;
  height: 16px;
  padding: 0 4px;
  border-radius: 999px;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-family: var(--platform-typography-font-family-mono);
  font-size: 10px;
  font-weight: 700;
`

const Caret = styled.span<{ $open?: boolean }>`
  font-size: 9px;
  transform: rotate(${props => (props.$open ? '180deg' : '0deg')});
  transition: transform 120ms ease;
`

const NotesPanel = styled.div`
  max-height: 168px;
  overflow-y: auto;
  padding: 4px 8px 8px;
  display: flex;
  flex-direction: column;
  gap: 2px;
`

// One note in the list: a platform list row (36px, hairline below).
const NoteRow = styled.button.attrs(chrome('list-row'))<{ $active?: boolean }>`
  width: 100%;
  border: 0;
  border-bottom: 1px solid var(--pure-chrome-line);
  text-align: left;
  background: ${props =>
    props.$active ? 'var(--pure-chrome-selection)' : 'transparent'};
  cursor: pointer;
`

const NoteCell = styled(MetaText)`
  flex: 0 0 auto;
  color: var(--sheets-ink);
  font-weight: 700;
`

const NotePreview = styled.span`
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--pure-chrome-soft);
  font-size: var(--pure-chrome-ui-size);
`
