# Drawer-owned Sheets reasoning

The drawer is now the only owner of spreadsheet reasoning. Removed `runSheetsTransform`, the low/high model invocations, private OpenCode session creation/polling/deletion, response schemas and obsolete runner types. Existing workbook read/write tools remain the execution surface. No shell code changed.

`getExchangeRate(from,to)` replaces the hidden currency prefetch with a deterministic public-data lookup. It returns a published reference rate, date, retrieval timestamp and source URL. It validates the requested base, amount, positive finite rate and source date, times out after ten seconds, and never fabricates a fallback or writes cells. Other external facts require the drawer's available live retrieval tools and citations; if unavailable, leave dependent cells unchanged.

`getSheetsContext` and `readRange` return an SHA-256 version covering the workbook and file binding. `applyCellChanges` requires an explicit sheet, target range and expectedVersion; it checks again inside the synchronous write callback. All proposed cells must lie in the target range. Malformed batches are rejected before mutation. Reads are bounded before expanding a range. Normal sanitization, locked-cell handling, before/after logs and undo remain.

The document ref is published synchronously before a write receipt returns. `completeWorkbook` retains ownership of flush/save/readback evidence; cell-write success alone does not imply disk persistence.

Validation: 66 focused tests passed (drawer tools, exchange rates, agent invariants, data integrity, 30 component tests and five completion receipt tests). Typecheck and production build passed. A live read-only Frankfurter lookup returned HTTP 200 with a dated USD/EUR rate. Build emitted existing bundle-size/dynamic-import warnings.

No end-to-end Electron/model mission performed: Electron was running a Canvas task and was left undisturbed. Follow-up smoke scenarios: source data → formulas → save/readback; dated currency conversion with source cells; protected range; concurrent user edit; document switch; unavailable external evidence; failure then repair using current version. Parent suite gitlink publication is separate from this submodule merge.
