# PureSheets Pivot Tables — Design Plan (Phase S4 deliverable, design only)

Status: **planning** — no implementation in this pass, per
`MAIL_SHEETS_MASTER_PLAN.md` §4 (S4: "Pivot tables: planning/design doc only").

## Grounding: what exists today

- **Document model** (`src/types.ts`, `src/lib/workbookModel.ts`):
  `PureSheetsDocument` → `WorkbookSheet` with a sparse `cells` record,
  formulas evaluated by Univer live (legacy shim for fallback paths),
  positional metadata (comments, validation, protection, conditional
  formats) that shifts with structural edits, and `charts` as
  PureSheets-owned derived artifacts.
- **Snapshot bridge** (`src/lib/univerAdapter.ts`): everything the Univer
  snapshot does not natively carry persists in `custom.puresheets` and is
  restored by the shared normalizers (`src/lib/sheetMetadataNormalize.ts`).
- **Derived-artifact precedent**: charts already model the pattern a pivot
  needs — a *source range* plus a *derived, re-computable projection* owned
  by PureSheets, rendered outside the Univer canvas, persisted in the sheet.
- **Agent invariant substrate** (S3/S5): protected ranges are enforced in the
  model mutation paths; agent edits go to draft sheets and reach sources only
  through QA approval.
- **Univer capabilities**: the installed preset (`@univerjs/preset-sheets-core`
  0.25) has **no pivot plugin**. Univer's pivot offering is part of its
  commercial/pro tier, not the OSS preset. A native-canvas pivot is therefore
  not on the table without a licensing decision.

## Design decision (proposed)

Model pivots as **derived sheets**, not floating widgets:

- A pivot is a `PivotDefinition` owned by a *destination sheet* (like
  `agentDraft` metadata marks draft sheets today):

  ```ts
  interface PivotDefinition {
    id: string
    sourceSheetId: string
    sourceRange: string            // header row + data rows
    rows: PivotFieldRef[]          // group-by fields (column indexes)
    columns: PivotFieldRef[]      // optional cross-tab fields
    values: PivotValueRef[]       // { field, aggregate: sum|count|avg|min|max }
    filters?: PivotFilterRef[]    // simple equals/contains per field
    refreshedAt: string
  }

  interface WorkbookSheet {
    // ...existing fields
    pivot?: PivotDefinition       // marks this sheet as a pivot projection
  }
  ```

- **Computation is a pure model function** —
  `computePivot(source: WorkbookSheet, def: PivotDefinition): WorkbookSheet`
  — producing ordinary cells (values + a table-style preset for header/totals
  rows). No new rendering surface: the result is a normal sheet on both the
  Univer canvas and the fallback grid, and it round-trips through the
  existing hard gate with zero adapter work beyond one `custom.puresheets`
  field (`pivot`, normalized like `validations`).

- **Refresh, not live linkage**: the pivot sheet is a snapshot projection
  with an explicit "Refresh" action (and an optional refresh-on-open prompt
  when `refreshedAt` predates the source's `updatedAt`). This matches the
  existing chart semantics (charts also hold materialized labels/values) and
  avoids reactive dependency tracking in v1.

- **Protection synergy**: the generated cells are wrapped in a locked range
  (`protection.lockedRanges = [usedRange]`) so users edit the *definition*,
  not the projection — reusing the S3 enforcement instead of new rules.

- **Agent/QA fit (S5)**: "pivot this data" becomes an agent transform that
  proposes a `PivotDefinition` through the normal QA draft flow; approval
  materializes the pivot sheet. The invariant (agents never mutate source
  sheets) holds for free because pivots only ever *read* the source range.

## v1 scope (when implementation is scheduled)

1. Model: `PivotDefinition`, `computePivot` (row grouping + one value
   aggregate: sum/count/avg/min/max), normalizer, snapshot persistence.
2. UI: "Insert pivot sheet" from the selected range (header row inferred),
   quiet editor panel on the pivot sheet (fields as option buttons — same
   pattern as the S3/S4 rule popovers), Refresh action in the sheet tab bar.
3. Structural safety: pivot `sourceRange` shifts with insert/delete via the
   existing `shiftRangeTokenForAxis`; a deleted source sheet marks the pivot
   stale (banner, keep last values) rather than deleting data.
4. Tests: `computePivot` goldens (grouping, aggregates, blanks, mixed
   types), round-trip fixture with a pivot sheet, structural-shift tests,
   stale-source behavior.

## v2+ (explicitly out of v1)

- Column (cross-tab) fields and multi-value aggregates.
- Live auto-refresh on source edits (needs a dependency graph — evaluate
  Univer's formula dependency service before building our own).
- Pivot charts (feed `chartFromRange` from the pivot projection).
- Univer Pro pivot plugin adoption if licensing changes (would replace the
  projection renderer but keep `PivotDefinition` as the portable format).

## Risks

- Aggregation over the legacy shim vs Univer engine could disagree on edge
  values — pivot computation must use `computedCellValue` (one code path)
  and be covered by the engine-agnostic golden harness pattern from S2.
- Large sources: v1 caps source ranges (e.g. 50k cells) with a clear message,
  mirroring the status-summary cap.
