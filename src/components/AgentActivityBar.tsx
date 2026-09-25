import type React from 'react'
import styled from 'styled-components'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import type { AgentLogEntry } from '../types'

const MAX_RENDERED_ENTRIES = 50

function formatTime(iso: string): string {
  const parsed = new Date(iso)
  if (Number.isNaN(parsed.getTime())) return iso
  return parsed.toLocaleString(undefined, {
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}

/**
 * Agent activity — the read surface for the workbook's agent log. Every write
 * a drawer-agent tool makes (cell changes, charts, sheets, comments,
 * formatting) lands in `workbook.agentLog` with before/after; this panel
 * lists those entries (time, agent, tool, summary, range) so agent work on
 * the workbook is inspectable without any tool call. It is a log viewer,
 * not a review UI — the same store the `listSheetChanges` agent tool reads.
 * Rendered under the formula bar when toggled on, like the comment bar.
 */
export function AgentActivityBar({
  entries,
  onClose,
}: {
  /** The workbook's agent log, newest first. */
  entries: AgentLogEntry[]
  onClose: () => void
}): React.ReactElement {
  const rendered = entries.slice(0, MAX_RENDERED_ENTRIES)
  return (
    <Wrap aria-label="Agent activity">
      <Bar>
        <Title>Agent activity</Title>
        <Meta>
          {entries.length === 0
            ? 'No agent changes recorded in this workbook yet.'
            : `${entries.length} recorded change${
                entries.length === 1 ? '' : 's'
              }, newest first`}
        </Meta>
        <CloseButton type="button" onClick={onClose}>
          Close
        </CloseButton>
      </Bar>
      {rendered.length > 0 ? (
        <EntryList aria-label="Agent log entries">
          {rendered.map(entry => (
            <EntryRow key={entry.id}>
              <EntryTime dateTime={entry.at}>{formatTime(entry.at)}</EntryTime>
              <EntryTool>
                {entry.agentName} · {entry.tool}
              </EntryTool>
              <EntrySummary title={entry.summary}>{entry.summary}</EntrySummary>
              <EntryScope>
                {entry.sheetName}
                {entry.range ? ` · ${entry.range}` : ''}
                {entry.cells?.length
                  ? ` · ${entry.cells.length} cell${
                      entry.cells.length === 1 ? '' : 's'
                    }`
                  : ''}
              </EntryScope>
            </EntryRow>
          ))}
          {entries.length > rendered.length ? (
            <MoreNote>
              Showing the latest {rendered.length} of {entries.length} entries.
            </MoreNote>
          ) : null}
        </EntryList>
      ) : null}
    </Wrap>
  )
}

const Wrap = styled.section`
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-well);
`

const Bar = styled.div`
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 6px 14px;
`

const Title = styled.strong`
  flex: 0 0 auto;
  color: var(--sheets-ink);
  font-size: var(--pure-chrome-ui-size);
  font-weight: 600;
`

const Meta = styled(MetaText)`
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
`

/** Typed loosely on purpose: styled-components' attrs rejects data-* literals. */
const selectChrome: Record<string, string> = { 'data-chrome': 'toolbar-select' }

const CloseButton = styled.button.attrs(selectChrome)`
  flex: 0 0 auto;
  cursor: pointer;

  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const EntryList = styled.div`
  display: flex;
  flex-direction: column;
  gap: 2px;
  max-height: 200px;
  padding: 2px 8px 8px;
  overflow-y: auto;
`

const EntryRow = styled.div`
  display: flex;
  gap: 10px;
  align-items: baseline;
  padding: 4px 8px;
  border-radius: 5px;

  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const EntryTime = styled(MetaText).attrs({ as: 'time' })`
  flex: 0 0 auto;
`

const EntryTool = styled.span`
  flex: 0 0 auto;
  color: var(--sheets-ink);
  font-size: var(--pure-chrome-ui-size);
  font-weight: 600;
  white-space: nowrap;
`

const EntrySummary = styled.span`
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--platform-colors-text);
  font-size: var(--pure-chrome-ui-size);
`

const EntryScope = styled(MetaText)`
  flex: 0 0 auto;
`

const MoreNote = styled(MetaText)`
  display: block;
  padding: 4px 8px;
`
