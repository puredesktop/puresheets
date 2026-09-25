# PureSheets UI + Legacy Spreadsheet Function Completion Plan

## Summary

Rework PureSheets into a professional spreadsheet app while preserving the emerging PureDesktop enterprise look. Keep the clean, restrained product language already established in PureSheets and the broader suite, but replace the clunky custom spreadsheet controls with familiar spreadsheet behavior and polished icon-based command groups.

Current finding: PureSheets already depends on `@univerjs/core` and `@univerjs/preset-sheets-core`, but the visible app is mostly a custom React grid. The next build should use Univer more directly for spreadsheet behavior while keeping PureDesktop's enterprise shell, file model, and AI hooks.

## Key Changes

- Preserve the PureDesktop enterprise aesthetic:
  - quiet neutral chrome,
  - compact controls,
  - subtle borders,
  - restrained color,
  - dense but readable spacing,
  - no playful or consumer-style spreadsheet skin.
- Replace toolbar buttons like `S`, `U`, `R`, `FR`, `FC`, `W+`, `H+` with professional icon controls and tooltips.
- Use a familiar spreadsheet structure:
  - app header / save state,
  - compact command bar,
  - formula/name box row,
  - grid with row and column headers,
  - bottom sheet tabs/status bar.
- Use Univer for spreadsheet editing/rendering/formulas where possible.
- Keep PureDesktop responsible for:
  - file routing,
  - `.sheets` / `.sheets.html` persistence,
  - agent context,
  - app-level actions.

## Implementation Changes

- Add a `UniverSpreadsheetSurface` that mounts the Univer spreadsheet editor inside the existing PureSheets shell.
- Keep the current PureSheets document format and embedded Univer snapshot as the persistence bridge.
- Promote Univer snapshot import/export to the main editing bridge instead of using Univer only for validation.
- Retire duplicated custom spreadsheet logic gradually:
  - first replace the visible grid/editing surface,
  - then remove custom behavior once parity is verified.
- Replace the hand-written formula evaluator for live editing with Univer's formula engine.
- Keep a compatibility layer for opening old PureSheets files and extracting agent context.

## Enterprise UI Direction

- Command bar should feel closer to OnlyOffice / Excel web / enterprise SaaS than a demo grid.
- Toolbar groups:

  ```text
  [Save] [Undo] [Redo] | [Format] [Font size]
  [B] [I] [U] [Text color] [Fill] [Borders]
  [Align] [Number format] | [Sort] [Filter] [Freeze] [Chart] | [Find]
  ```

- Bottom bar:

  ```text
  [+ Sheet] [Sheet 1] [Sheet 2] ... [autosave/status/selection summary]
  ```

- Use icons from the suite's existing icon system where available.
- Use text only for menus, dropdown values, sheet names, and status.
- Keep controls compact and aligned; avoid oversized buttons, floating cards, or loud color blocks.

## Required Spreadsheet Functions

Expose the 80/20 spreadsheet surface through Univer where available:

- Cell editing, multi-cell selection, copy/cut/paste, delete, clear.
- Formatting: bold, italic, underline, font size, text color, fill color, borders, alignment.
- Number formats: text, number, currency, percent, date.
- Structure: insert/delete rows and columns, resize rows/columns, freeze row/column.
- Sheets: add, rename, duplicate, delete.
- Data tools: sort, filter, find, find/replace if cleanly supported.
- Formula compatibility: preserve support for existing formulas such as `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `ROUND`, `CONCAT`, `LEFT`, `RIGHT`, `VLOOKUP`, `TODAY`, and `NOW`.

## Test Plan

- Existing `.sheets` and `.sheets.html` files still open and save.
- Existing autosave, corrupt-file, undo/redo, freeze, chart, and selected-range behavior is preserved or replaced with equivalent Univer tests.
- Formula compatibility test confirms current formulas still calculate.
- UI tests confirm:
  - no mystery text buttons remain,
  - formula bar is visible,
  - row/column headers are visible,
  - sheet tabs are visible,
  - toolbar controls have accessible labels/tooltips.
- Visual checks confirm the result matches the established PureDesktop enterprise style on desktop and narrow layouts.

## Assumptions

- "Legacy spreadsheet functions" means familiar spreadsheet behavior expected from enterprise spreadsheet tools, not a separate older PureSheets codebase.
- Univer should become the primary spreadsheet behavior layer.
- PureSheets should look like part of PureDesktop, not like a generic embedded spreadsheet demo.
