# PureSheets reliability audit — 2026-10-03

The editable source of truth is `PureSheetsDocument` in `PureSheetsShell`.
Univer owns the live grid; the app captures its committed snapshot before a save,
export, completion receipt, or workbook replacement. The unified document lifecycle
owns filesystem writes. No shared shell or SDK source changes were needed.

## Changes

- Pin Univer core and both presets to 0.25.1. Mixed 0.25.2 core / 0.25.1
  plugins produced duplicate service identities and injector failures. Declare the
  headless formula test plugins directly at that same version.
- Use one generation-controlled Open/New/XLSX replacement boundary. Commit pending
  edits and save the previous workbook successfully before installing the new
  document, format, and file binding. Failed reads remain retryable; only successful
  matching host requests are consumed. Edits and agent writes cannot cross an open.
- Commit both live in-cell edits and unblurred formula-bar typing before Save.
  Await any formula-bar cell update already in progress. A failed snapshot or write
  cannot report save success. Failed opens do not create empty drafts.
- Preserve package, legacy flat JSON, and HTML formats when saving the predecessor.
  Recover legacy backups after unreadable main files as well as corrupt JSON/HTML.
  Preserve package permission/timeout/parse errors instead of guessing another format.
- Invalidate asynchronous editor exports on unit reload or disposal, and reject
  delayed agent mutations from a previous workbook context.
- Bound selection-summary and single-cell-reference checks before allocating range
  keys, avoiding full-sheet allocations merely to decide to skip a summary.
- Give formula input styling precedence over shared field styling, so formula
  highlighting does not draw the text twice. Keep the formula bar and sheet tabs
  usable at narrow widths; constrain long status messages.
- Declare the compatible SWC compiler/plugin pair and happy-dom test environment.
  Run tests with a separate Vitest configuration, avoiding build-only transforms.

## Verification

- `npm test`: 296 tests pass in 23 files, including all 71 headless formula checks.
- `npm run typecheck` and `npm run build`: pass. The build still reports large
  existing Univer/ExcelJS chunks; this audit does not claim to remove that cost.
- Twenty regression cases were added. Seventeen persistence/editor cases were
  rerun against the original source and failed. The original full-sheet allocation
  was deliberately not executed: materializing billions of keys is unsafe.
- Browser: the real Univer grid loaded and calculated `SUM(10,20)` as 30. In an
  iframe harness with a disposable mock bridge, Command-S saved formula-bar typing
  without blurring the field; the saved workbook JSON contained the formula.
  Formula text rendered once, and the formula bar and sheet tabs remained visible
  at 390px. A browser MutationObserver diagnostic appeared without an app stack;
  the grid and mock save remained functional.

The iframe harness was a test host, not the installed Electron host. Real host
filesystem writes, native import/export dialogs, crash recovery and simultaneous
external writers were not exercised. No real workbooks, assistant requests, or
installer builds were used. Temporary harness files were removed.
