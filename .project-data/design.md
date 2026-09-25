# PureSheets Product Plan

## Product Goal

PureSheets is the suite spreadsheet workspace for editing workbook packages and portable spreadsheet documents with a desktop-style grid, formula, chart, validation, and sheet navigation surface.

## Current Product Surface

- Univer-backed spreadsheet editing surface.
- Command bar, formula bar, sheet tabs, chart panel, comments, compatibility grid, and data rule controls.
- Workbook model and file IO for `.sheets` and `.sheets.html` workflows.
- CSV/XLSX import-export and formula/data integrity test coverage.
- Agent context tool for active file, selected sheet, selected cell, dirty state, and visible workbook summary.

## Development Notes

Spreadsheet model and IO code lives under `src/lib`; the grid shell and controls live under `src/components`; bridge access is isolated in `src/bridge/platformBridge.ts`. Future changes should keep spreadsheet behavior in model/adapter modules instead of burying logic inside UI controls.
