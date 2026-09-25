# PureSheets Agent Safe-Duplicate QA Plan

## Summary

PureSheets agents must never directly mutate the source-of-truth sheet. When a user asks an agent to transform a selected range, the app creates an agent draft sheet, copies the source data there, runs the operation on the draft, and reports what changed. The user can then approve individual changes or approve all changes, which reconciles the accepted operations back into the original sheet.

The original sheet stays authoritative. The duplicate sheet is an operation workspace. Accepted changes become part of a readable linear history.

## Key Changes

- Original sheets are always the source of truth.
- Agent transforms always happen on an agent-created duplicate sheet.
- Selecting a range can show a small prompt icon for asking the agent to transform that range.
- If the prompt starts from the original sheet, PureSheets tells the user it is creating a draft sheet before applying the transform.
- Agent changes and QA highlights appear on the draft sheet.
- Users can approve one change, selected changes, or all changes.
- Approved changes are applied back to the source sheet.
- Approved and discarded operations become readable history.

## Implementation Changes

- Add an agent mutation rule:
  - user edits may change source sheets,
  - agent edits must target an agent draft sheet,
  - approval is the only path from draft back to source.
- On agent transform from a source sheet:
  - capture selected range and prompt,
  - duplicate the relevant sheet,
  - name it clearly, for example `Forecast - Agent draft`,
  - copy source data and metadata,
  - store source/draft links,
  - run the agent operation on the draft,
  - create QA ledger entries comparing draft cells with source cells.
- On agent transform from an existing draft:
  - keep changes inside that draft,
  - append a new operation to the same draft history unless the user chooses a new draft.
- Add metadata to draft sheets:
  - `sourceSheetId`,
  - `sourceSheetName`,
  - `agentRunId`,
  - `agentName`,
  - `createdAt`,
  - `selectedSourceRange`,
  - `prompt`,
  - `qaStatus`.
- Add operation history entries:
  - operation id,
  - prompt,
  - affected source range,
  - affected draft range,
  - before/after values,
  - explanation,
  - approval status,
  - timestamp.

## Prompt Interaction

- When a range is selected, show a small prompt icon near the selection.
- Clicking the prompt icon opens a compact prompt panel:
  ```text
  Ask agent to transform this range...
  ```
- On submit from a source sheet:
  - show a clear message: "Creating an agent draft so your original sheet stays unchanged."
  - create the draft,
  - switch to the draft or offer to open it,
  - run the transform there.
- On submit from a draft sheet:
  - run the transform inside the draft,
  - update QA and history.
- The agent must report:
  - what it changed,
  - which cells/ranges changed,
  - why,
  - any cells it skipped or could not process.

## QA And Approval UI

- QA tab groups changes by draft sheet and operation:
  ```text
  Forecast - Agent draft
  Source: Forecast
  Prompt: Normalize Q1 revenue labels
  Changed: 14 cells
  ```
- Each QA row includes:
  - source cell/range,
  - draft cell/range,
  - original value/formula/style,
  - proposed value/formula/style,
  - explanation,
  - status,
  - approve action.
- Draft sheet highlights changed cells.
- Hovering a highlighted draft cell shows the source/draft diff.
- Approval actions:
  - approve single change,
  - approve selected changes,
  - approve all changes in operation,
  - approve all changes in draft,
  - discard operation,
  - discard draft.
- When approved:
  - copy the proposed value/formula/style from draft to source,
  - mark the QA entry approved,
  - write a linear history record,
  - preserve who/what/when/why.
- If the source cell changed after draft creation:
  - mark the QA entry stale,
  - block approval for that entry,
  - offer compare/re-run/manual resolve.

## Linear History

- Accepted and rejected operations should be reviewable later in a history view.
- History should be readable:
  ```text
  10:42 Finance Assistant proposed 14 changes to Forecast from prompt "Normalize Q1 revenue labels".
  10:46 Adam approved 12 changes and rejected 2.
  ```
- History should link back to:
  - source sheet,
  - draft sheet if still retained,
  - QA operation,
  - affected ranges.
- Draft sheets may be archived or deleted later, but approved history remains.

## Test Plan

- Prompt icon appears when a valid range is selected.
- Agent prompt from source sheet creates a draft before transforming.
- Source sheet remains unchanged until approval.
- Agent prompt from draft sheet modifies only the draft.
- QA entries compare source values to draft values.
- Highlights appear on draft changed cells.
- Approving one change updates only the matching source cell.
- Approving all applies all non-stale proposed changes.
- Rejected changes never modify the source sheet.
- Source edits after draft creation make affected QA entries stale.
- History records approved, rejected, and discarded operations.
- Saving and reopening preserves source/draft links, QA entries, and history.

## Assumptions

- V1 duplicates a whole sheet for each source-sheet transform.
- Workbook-wide transforms may create one draft per affected sheet.
- The original sheet remains the only source of truth.
- Agent draft sheets are operation workspaces, not final data unless reconciled.
- QA history is retained even if the draft sheet is later discarded.
