export const PURESHEETS_APP_SLUG = 'sheets'

export const DEFAULT_ROW_COUNT = 40
export const DEFAULT_COLUMN_COUNT = 18

// Size of the grid the Univer canvas actually renders. The model grows cells
// on demand (a sparse map), but Univer only paints within the snapshot's
// row/column bounds — a small floor left everything past ~row 40 / col 18 as
// grey void (#272). These give the canvas a full spreadsheet-sized grid
// (Google Sheets ships 1000×26) regardless of how much data the sheet holds;
// the model/import/compat-grid defaults above stay small on purpose.
export const UNIVER_SURFACE_ROW_COUNT = 1000
export const UNIVER_SURFACE_COLUMN_COUNT = 26

// Cell border color for the "toggle all borders" style. Must stay clearly
// darker than the canvas gridlines (`--sheets-line`, #d9dde3) — the previous
// #b8c0b2 sat at almost the same lightness, so toggled borders were
// imperceptible (#270). Matches the platform text-muted tone.
export const CELL_BORDER_COLOR = '#687064'
