import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import {
  cellKey,
  columnLabel,
  computedCellValue,
  createBlankSheet,
  createBlankWorkbook,
  parseCellKey,
} from './workbookModel'
import { withFreshUniverSnapshot } from './univerAdapter'
import { patternOfStyle, styleFromPattern } from './numberFormats'
import { DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT } from '../constants'
import type {
  CellStyle,
  PureSheetsDocument,
  SheetCell,
  WorkbookSheet,
} from '../types'

/**
 * Excel (.xlsx) import and export, built on exceljs.
 *
 * Import turns a workbook into an ordinary PureSheetsDocument: every sheet and
 * its name, values and formulas (cross-sheet ones included), number formats,
 * merged cells, column widths and row heights, bold/italic/underline, font
 * size and colour, fill, borders by side, alignment and wrap, frozen panes and
 * cell notes. What the model cannot hold is counted and named in plain
 * warnings, never dropped silently. Export writes the same things back, so a
 * round trip keeps values and formulas identical. `.sheets` stays the native
 * format; Excel is a door in and out.
 */

export interface XlsxImportResult {
  document: PureSheetsDocument
  warnings: string[]
}

export interface XlsxExportResult {
  bytes: Uint8Array
  warnings: string[]
}

const MAX_IMPORTED_CELLS = 250_000

export async function importXlsxWorkbook(
  data: ArrayBuffer | Uint8Array,
  title: string,
): Promise<XlsxImportResult> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
  const workbook = await loadWorkbook(bytes)
  const palette = themePalette(
    (workbook as unknown as { _themes?: Record<string, string> })._themes?.theme1,
  )
  const losses = new LossCounter()
  const sheets: WorkbookSheet[] = []
  let importedCells = 0

  workbook.eachSheet(worksheet => {
    const name = worksheet.name || `Sheet ${sheets.length + 1}`
    const sheetLosses = new LossCounter()
    const cells: Record<string, SheetCell> = {}
    const comments: Record<string, string> = {}
    let maxRow = 0
    let maxColumn = 0

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
        if (++importedCells > MAX_IMPORTED_CELLS)
          throw new Error(
            'This workbook has more than 250,000 filled cells, more than PureSheets imports at once. Split it by sheet and import the parts.',
          )
        const key = cellKey(rowNumber - 1, columnNumber - 1)
        const note = noteText(cell.note)
        if (note) comments[key] = note
        const mapped = mapXlsxCell(cell, worksheet, palette, sheetLosses)
        if (!mapped) return
        cells[key] = mapped
        maxRow = Math.max(maxRow, rowNumber)
        maxColumn = Math.max(maxColumn, columnNumber)
      })
    })

    const columnWidths: Record<string, number> = {}
    let hiddenColumns = 0
    worksheet.columns?.forEach((column, index) => {
      if (column?.hidden) hiddenColumns += 1
      if (typeof column?.width === 'number')
        columnWidths[String(index)] = excelWidthToPixels(column.width)
    })
    const defaultRowHeight = worksheet.properties.defaultRowHeight ?? 15
    const rowHeights: Record<string, number> = {}
    let hiddenRows = 0
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (row.hidden) hiddenRows += 1
      // A height equal to the sheet default is Excel's, not the author's.
      if (typeof row.height === 'number' && Math.abs(row.height - defaultRowHeight) > 0.01)
        rowHeights[String(rowNumber - 1)] = pointsToPixels(row.height)
    })

    const view = worksheet.views?.[0]
    const frozen =
      view?.state === 'frozen'
        ? { rows: view.ySplit ?? 0, columns: view.xSplit ?? 0 }
        : { rows: 0, columns: 0 }
    if (view?.state === 'split') sheetLosses.add('split panes (shown unsplit)')

    const merges = ((worksheet.model as { merges?: string[] }).merges ?? []).filter(merge =>
      /^[A-Z]+\d+:[A-Z]+\d+$/i.test(merge),
    )
    for (const merge of merges) {
      const end = parseCellKey(merge.split(':')[1])
      if (end) {
        maxRow = Math.max(maxRow, end.row + 1)
        maxColumn = Math.max(maxColumn, end.column + 1)
      }
    }

    if (worksheet.state && worksheet.state !== 'visible')
      sheetLosses.add('hidden sheet (shown as an ordinary sheet)')
    if (hiddenRows) sheetLosses.add('hidden rows (shown)', hiddenRows)
    if (hiddenColumns) sheetLosses.add('hidden columns (shown)', hiddenColumns)
    const images = worksheet.getImages().length
    if (images) sheetLosses.add('pictures', images)
    const validations = Object.keys(
      (worksheet as unknown as { dataValidations?: { model?: Record<string, unknown> } })
        .dataValidations?.model ?? {},
    ).length
    if (validations)
      sheetLosses.add('data-validation rules (add them again with Data validation)', validations)
    const conditional = (
      worksheet as unknown as { conditionalFormattings?: unknown[] }
    ).conditionalFormattings?.length ?? 0
    if (conditional)
      sheetLosses.add('conditional-formatting rules (add them again with Conditional format)', conditional)
    const tables = Object.keys(
      (worksheet as unknown as { tables?: Record<string, unknown> }).tables ?? {},
    ).length
    if (tables)
      sheetLosses.add('Excel tables (their cells came in; formulas that name a table column will not calculate)', tables)
    if ((worksheet as unknown as { autoFilter?: unknown }).autoFilter)
      sheetLosses.add('filter buttons (turn the filter on again)')
    if ((worksheet as unknown as { sheetProtection?: unknown }).sheetProtection) sheetLosses.add('sheet protection (lock ranges again with Protect)')

    losses.addSection(name, sheetLosses)
    sheets.push({
      ...createBlankSheet(`sheet-xlsx-${sheets.length + 1}`, name),
      rowCount: Math.max(DEFAULT_ROW_COUNT, maxRow),
      columnCount: Math.max(DEFAULT_COLUMN_COUNT, maxColumn),
      cells,
      columnWidths: Object.keys(columnWidths).length ? columnWidths : undefined,
      rowHeights: Object.keys(rowHeights).length ? rowHeights : undefined,
      frozenRows: frozen.rows || undefined,
      frozenColumns: frozen.columns || undefined,
      merges: merges.length ? merges : undefined,
      comments: Object.keys(comments).length ? comments : undefined,
    })
  })

  const names = (workbook.definedNames.model ?? []).filter(
    (entry: { name?: string }) => entry.name && !entry.name.startsWith('_xlnm.'),
  ).length
  if (names)
    losses.add('named ranges (formulas that use a name will show #NAME?)', names)
  for (const part of unsupportedParts(listZipEntries(bytes))) losses.add(part)

  const warnings = losses.lines()
  if (!sheets.length) return { document: createBlankWorkbook(title), warnings }
  const now = new Date().toISOString()
  const document = withFreshUniverSnapshot({
    app: 'PureSheets',
    version: 1,
    engine: { name: 'univer' },
    metadata: { title, createdAt: now, updatedAt: now },
    workbook: { activeSheetId: sheets[0].id, sheets },
  })
  return { document, warnings }
}

async function loadWorkbook(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook()
  try {
    await workbook.xlsx.load(bytes.slice().buffer as ArrayBuffer)
    return workbook
  } catch (firstError) {
    // Files written by .NET tools prefix every element (<x:workbook>) and point at parts by absolute path,
    // neither of which exceljs reads.
    const unprefixed = await normalizedForExcelJs(bytes).catch(() => null)
    if (unprefixed) {
      const retry = new ExcelJS.Workbook()
      try {
        await retry.xlsx.load(unprefixed)
        return retry
      } catch {
        // fall through to the plain message
      }
    }
    throw new Error(
      `This file could not be read as an Excel workbook${
        firstError instanceof Error && firstError.message ? ` (${firstError.message})` : ''
      }. If it opens in Excel, save it again as .xlsx and import that copy.`,
    )
  }
}

const UNPREFIXED_NAMESPACES = [
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://schemas.openxmlformats.org/package/2006/relationships',
  'http://schemas.openxmlformats.org/package/2006/content-types',
]

/**
 * Rewrite each XML part so the main spreadsheet namespace is the default one
 * (`<x:sheet>` → `<sheet>`), and each relationship target relative to its part
 * (`/xl/styles.xml` → `styles.xml`). Returns null when nothing needed it.
 */
export async function normalizedForExcelJs(bytes: Uint8Array): Promise<ArrayBuffer | null> {
  const zip = await JSZip.loadAsync(bytes)
  let changed = false
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || !/\.(xml|rels)$/i.test(entry.name)) continue
    const xml = await entry.async('string')
    let next = xml
    for (const namespace of UNPREFIXED_NAMESPACES) {
      const declared = new RegExp(`xmlns:([A-Za-z_][\\w.-]*)="${namespace.replace(/[.]/g, '\\.')}"`).exec(next)
      if (!declared) continue
      const prefix = declared[1]
      next = next
        .replace(new RegExp(`<(/?)${prefix}:`, 'g'), '<$1')
        .replace(declared[0], /\sxmlns="/.test(next) ? '' : `xmlns="${namespace}"`)
    }
    const relsOf = /^(.*?)_rels\/[^/]+\.rels$/.exec(entry.name)
    if (relsOf) {
      const base = relsOf[1]
      next = next.replace(/(\sTarget=")\/([^"]+)"/g, (whole, lead: string, target: string) =>
        /TargetMode="External"/.test(whole) ? whole : `${lead}${relativePath(base, target)}"`,
      )
    }
    if (next !== xml) {
      zip.file(entry.name, next)
      changed = true
    }
  }
  return changed ? zip.generateAsync({ type: 'arraybuffer' }) : null
}

/** `target` (from the zip root) as a path relative to the folder `base` ('xl/', 'xl/worksheets/', ''). */
function relativePath(base: string, target: string): string {
  const from = base.split('/').filter(Boolean)
  const to = target.split('/').filter(Boolean)
  let shared = 0
  while (shared < from.length && shared < to.length - 1 && from[shared] === to[shared]) shared += 1
  return [...from.slice(shared).map(() => '..'), ...to.slice(shared)].join('/')
}

/**
 * Tallies what an import or export could not carry, by sheet, and words it
 * as one plain line per sheet ("Budget: 2 pictures, 3 hidden rows (shown)").
 */
class LossCounter {
  private counts = new Map<string, number>()
  private sections: string[] = []

  add(label: string, count = 1): void {
    this.counts.set(label, (this.counts.get(label) ?? 0) + count)
  }

  addSection(name: string, inner: LossCounter): void {
    const parts = inner.parts()
    if (parts.length) this.sections.push(`${name}: ${parts.join('; ')}.`)
  }

  lines(): string[] {
    const parts = this.parts()
    return [...this.sections, ...(parts.length ? [`Workbook: ${parts.join('; ')}.`] : [])]
  }

  private parts(): string[] {
    return [...this.counts].map(([label, count]) => (count > 1 ? `${count} ${label}` : label))
  }
}

function mapXlsxCell(
  cell: ExcelJS.Cell,
  worksheet: ExcelJS.Worksheet,
  palette: string[],
  losses: LossCounter,
): SheetCell | null {
  // The hidden parts of a merge report the top-left value; only its style is theirs.
  const mergedAway = cell.isMerged && cell.master !== cell
  let value = ''
  let kind: SheetCell['kind'] = 'text'
  const raw = mergedAway ? null : cell.value
  if (raw === null || raw === undefined) {
    value = ''
  } else if (typeof raw === 'number') {
    value = String(raw)
    kind = 'number'
  } else if (typeof raw === 'boolean') {
    value = raw ? 'TRUE' : 'FALSE'
  } else if (raw instanceof Date) {
    value = isoFromDate(raw)
  } else if (typeof raw === 'object') {
    const objectValue = raw as {
      formula?: string
      sharedFormula?: string
      shareType?: string
      result?: unknown
      richText?: Array<{ text: string }>
      text?: unknown
      hyperlink?: string
      error?: string
    }
    if (objectValue.formula) {
      value = `=${formulaFromExcel(objectValue.formula)}`
      kind = 'formula'
      if (objectValue.shareType === 'array')
        losses.add('array formulas (came in as ordinary formulas)')
    } else if (objectValue.sharedFormula) {
      const master = worksheet.getCell(objectValue.sharedFormula)
      const masterFormula = (master.value as { formula?: string } | null)?.formula
      if (masterFormula) {
        value = `=${formulaFromExcel(
          translateSharedFormula(
            masterFormula,
            Number(cell.row) - Number(master.row),
            Number(cell.col) - Number(master.col),
          ),
        )}`
        kind = 'formula'
      } else {
        value = plainText(objectValue.result)
        losses.add('formulas without their source (kept as their last value)')
      }
    } else if (objectValue.richText) {
      value = objectValue.richText.map(part => part.text).join('')
      losses.add('cells of mixed text formatting (kept as plain text)')
    } else if (objectValue.hyperlink) {
      value = plainText(objectValue.text) || objectValue.hyperlink
      losses.add('links (kept as their text)')
    } else if (objectValue.error) {
      value = objectValue.error
    } else {
      value = plainText(raw)
    }
  } else {
    value = String(raw)
  }
  const style = mapXlsxStyle(cell, palette, losses)
  if (!value && !style) return null
  if (!value) kind = 'blank'
  return { value, kind, ...(style ? { style } : {}) }
}

/** Text from whatever exceljs holds: a string, a number, a date, or runs of rich text. */
function plainText(input: unknown): string {
  if (input === null || input === undefined) return ''
  if (input instanceof Date) return isoFromDate(input)
  if (typeof input === 'object') {
    const object = input as { richText?: Array<{ text?: string }>; error?: string; text?: unknown }
    if (Array.isArray(object.richText)) return object.richText.map(part => part.text ?? '').join('')
    if (typeof object.error === 'string') return object.error
    if (object.text !== undefined) return plainText(object.text)
    return ''
  }
  return String(input)
}

/** Dates come back from exceljs as UTC instants; a whole day reads as yyyy-mm-dd. */
function isoFromDate(date: Date): string {
  if (Number.isNaN(date.getTime())) return ''
  const iso = date.toISOString()
  return iso.endsWith('T00:00:00.000Z') ? iso.slice(0, 10) : iso
}

function noteText(note: ExcelJS.Cell['note']): string {
  if (!note) return ''
  if (typeof note === 'string') return note.trim()
  return (note.texts ?? []).map(part => part.text).join('').trim()
}

function mapXlsxStyle(
  cell: ExcelJS.Cell,
  palette: string[],
  losses: LossCounter,
): CellStyle | undefined {
  const style: CellStyle = {}
  const font = cell.font
  if (font?.bold) style.bold = true
  if (font?.italic) style.italic = true
  if (font?.underline) style.underline = true
  if (font?.strike) losses.add('struck-through cells (shown without the line)')
  if (font?.name && !DEFAULT_FONTS.has(font.name.toLowerCase()))
    losses.add(`cells in ${font.name} (shown in the PureSheets font)`)
  if (typeof font?.size === 'number' && font.size !== 11)
    style.fontSize = clamp(Math.round(font.size), 6, 96)
  const textColor = resolveColor(font?.color, palette)
  // Black is Excel's automatic ink; leave it to the PureSheets default.
  if (textColor && textColor !== '#000000') style.textColor = textColor

  const fill = cell.fill as
    | { type?: string; pattern?: string; fgColor?: Partial<ExcelJS.Color> }
    | undefined
  if (fill?.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
    const fillColor = resolveColor(fill.fgColor, palette)
    if (fillColor && fillColor !== '#ffffff') style.fillColor = fillColor
    if (fill.pattern !== 'solid') losses.add('patterned fills (shown solid)')
  } else if (fill?.type === 'gradient') {
    losses.add('gradient fills (left out)')
  }

  const border = cell.border
  const sides = (['top', 'right', 'bottom', 'left'] as const).filter(
    side => border?.[side]?.style,
  )
  if (sides.length === 4) style.border = true
  else if (sides.length) style.borderSides = [...sides]
  if (
    sides.some(side => {
      const line = border?.[side]?.style
      return line && line !== 'thin' && line !== 'hair'
    })
  )
    losses.add('thick, dashed or double borders (shown thin)')

  const alignment = cell.alignment
  const horizontal = alignment?.horizontal
  if (horizontal === 'left' || horizontal === 'center' || horizontal === 'right') {
    style.align = horizontal
  } else if (horizontal === 'centerContinuous') {
    style.align = 'center'
  } else if (horizontal && (horizontal as string) !== 'general') {
    losses.add(`${horizontal} alignment (shown left)`)
  }
  const vertical = alignment?.vertical
  if (vertical === 'top' || vertical === 'bottom') style.verticalAlign = vertical
  else if (vertical === 'middle') style.verticalAlign = 'middle'
  else if (vertical) losses.add(`${vertical} vertical alignment (shown centred)`)
  if (alignment?.wrapText) style.wrap = 'wrap'
  if (alignment?.textRotation) losses.add('rotated text (shown level)')
  if (alignment?.indent) losses.add('indented cells (shown without the indent)')

  const numberStyle = styleFromPattern(cell.numFmt)
  if (numberStyle) Object.assign(style, numberStyle)
  return Object.keys(style).length ? style : undefined
}

const DEFAULT_FONTS = new Set(['calibri', 'calibri light', 'aptos', 'aptos narrow', 'aptos display', 'arial'])

// ---------------------------------------------------------------------------
// Colours: Excel cells name colours as ARGB, as a theme slot with a tint, or as
// an index into the legacy 64-colour palette.
// ---------------------------------------------------------------------------

/** Office 2013–2022 theme, used when a file carries no theme of its own. */
const OFFICE_THEME = [
  '#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31',
  '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#0563c1', '#954f72',
]

const INDEXED_COLORS = [
  '000000', 'ffffff', 'ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff',
  '000000', 'ffffff', 'ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff',
  '800000', '008000', '000080', '808000', '800080', '008080', 'c0c0c0', '808080',
  '9999ff', '993366', 'ffffcc', 'ccffff', '660066', 'ff8080', '0066cc', 'ccccff',
  '000080', 'ff00ff', 'ffff00', '00ffff', '800080', '800000', '008080', '0000ff',
  '00ccff', 'ccffff', 'ccffcc', 'ffff99', '99ccff', 'ff99cc', 'cc99ff', 'ffcc99',
  '3366ff', '33cccc', '99cc00', 'ffcc00', 'ff9900', 'ff6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
  '000000', 'ffffff',
]

/**
 * Theme slots in Excel's order (light 1, dark 1, light 2, dark 2, accents,
 * links) read from the file's theme XML, which lists dark before light.
 */
export function themePalette(themeXml: string | undefined): string[] {
  if (!themeXml) return OFFICE_THEME
  const scheme = /<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(themeXml)?.[1]
  if (!scheme) return OFFICE_THEME
  const slot = (name: string): string | undefined => {
    const body = new RegExp(`<a:${name}>([\\s\\S]*?)</a:${name}>`).exec(scheme)?.[1]
    const hex =
      /<a:srgbClr val="([0-9a-f]{6})"/i.exec(body ?? '')?.[1] ??
      /lastClr="([0-9a-f]{6})"/i.exec(body ?? '')?.[1]
    return hex ? `#${hex.toLowerCase()}` : undefined
  }
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
  return order.map((name, index) => slot(name) ?? OFFICE_THEME[index])
}

export function resolveColor(
  color: Partial<ExcelJS.Color> & { indexed?: number; tint?: number } | undefined,
  palette: string[],
): string | undefined {
  if (!color) return undefined
  let hex: string | undefined
  if (typeof color.argb === 'string' && /^([0-9a-f]{2})?[0-9a-f]{6}$/i.test(color.argb)) {
    hex = `#${color.argb.slice(-6).toLowerCase()}`
  } else if (typeof color.theme === 'number') {
    hex = palette[color.theme]
  } else if (typeof color.indexed === 'number' && INDEXED_COLORS[color.indexed]) {
    hex = `#${INDEXED_COLORS[color.indexed]}`
  }
  if (!hex) return undefined
  return typeof color.tint === 'number' && color.tint !== 0 ? applyTint(hex, color.tint) : hex
}

/** Excel's tint: lighten (positive) or darken (negative) in HSL lightness. */
function applyTint(hex: string, tint: number): string {
  const [r, g, b] = [1, 3, 5].map(index => parseInt(hex.slice(index, index + 2), 16) / 255)
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let lightness = (max + min) / 2
  const delta = max - min
  const saturation = delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1))
  let hue = 0
  if (delta !== 0) {
    if (max === r) hue = ((g - b) / delta) % 6
    else if (max === g) hue = (b - r) / delta + 2
    else hue = (r - g) / delta + 4
    hue *= 60
    if (hue < 0) hue += 360
  }
  lightness = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation
  const x = chroma * (1 - Math.abs(((hue / 60) % 2) - 1))
  const m = lightness - chroma / 2
  const [r1, g1, b1] =
    hue < 60 ? [chroma, x, 0] : hue < 120 ? [x, chroma, 0] : hue < 180 ? [0, chroma, x]
      : hue < 240 ? [0, x, chroma] : hue < 300 ? [x, 0, chroma] : [chroma, 0, x]
  return `#${[r1, g1, b1]
    .map(channel => Math.round(clamp((channel + m) * 255, 0, 255)).toString(16).padStart(2, '0'))
    .join('')}`
}

// ---------------------------------------------------------------------------
// Formulas
// ---------------------------------------------------------------------------

/**
 * Excel stores functions newer than 2007 with a `_xlfn.` prefix (and FILTER
 * and SORT with `_xlfn._xlws.`). PureSheets shows and evaluates them bare;
 * export puts the prefix back, so the file Excel reads is the one it wrote.
 */
const FUTURE_FUNCTIONS = new Set([
  'AGGREGATE', 'ARRAYTOTEXT', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.INV', 'BITAND', 'BITLSHIFT',
  'BITOR', 'BITRSHIFT', 'BITXOR', 'BYCOL', 'BYROW', 'CEILING.MATH', 'CEILING.PRECISE', 'CHISQ.DIST',
  'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'CHOOSECOLS', 'CHOOSEROWS', 'CONCAT',
  'CONFIDENCE.NORM', 'CONFIDENCE.T', 'COVARIANCE.P', 'COVARIANCE.S', 'DAYS', 'DROP', 'ERF.PRECISE',
  'ERFC.PRECISE', 'EXPAND', 'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST', 'FILTER',
  'FLOOR.MATH', 'FLOOR.PRECISE', 'FORECAST.LINEAR', 'FORMULATEXT', 'GAMMA', 'GAMMA.DIST', 'GAMMA.INV',
  'GAMMALN.PRECISE', 'HSTACK', 'HYPGEOM.DIST', 'IFNA', 'IFS', 'IMAGE', 'ISFORMULA', 'ISOMITTED',
  'ISOWEEKNUM', 'LAMBDA', 'LET', 'LOGNORM.DIST', 'LOGNORM.INV', 'MAKEARRAY', 'MAP', 'MAXIFS', 'MINIFS',
  'MODE.MULT', 'MODE.SNGL', 'NEGBINOM.DIST', 'NETWORKDAYS.INTL', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST',
  'NORM.S.INV', 'NUMBERVALUE', 'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTRANK.EXC', 'PERCENTRANK.INC',
  'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC', 'RANDARRAY', 'RANK.AVG', 'RANK.EQ', 'REDUCE', 'SCAN',
  'SEQUENCE', 'SHEET', 'SHEETS', 'SORT', 'SORTBY', 'STDEV.P', 'STDEV.S', 'SWITCH', 'T.DIST', 'T.DIST.2T',
  'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'TAKE', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTJOIN', 'TEXTSPLIT',
  'TOCOL', 'TOROW', 'UNICHAR', 'UNICODE', 'UNIQUE', 'VALUETOTEXT', 'VAR.P', 'VAR.S', 'VSTACK',
  'WEIBULL.DIST', 'WORKDAY.INTL', 'WRAPCOLS', 'WRAPROWS', 'XLOOKUP', 'XMATCH', 'XOR', 'Z.TEST',
])
const WORKSHEET_FUNCTIONS = new Set(['FILTER', 'SORT'])

export function formulaFromExcel(formula: string): string {
  return mapOutsideLiterals(formula, segment =>
    segment.replace(/_xlfn\.(?:_xlws\.)?([A-Z][A-Z0-9.]*)(?=\()/gi, (token, name: string) =>
      FUTURE_FUNCTIONS.has(name.toUpperCase()) ? name : token,
    ),
  )
}

export function formulaToExcel(formula: string): string {
  return mapOutsideLiterals(formula, segment =>
    segment.replace(/(?<![A-Za-z0-9_.])([A-Z][A-Z0-9.]*)(?=\()/gi, (token, name: string) => {
      const upper = name.toUpperCase()
      if (!FUTURE_FUNCTIONS.has(upper)) return token
      return `${WORKSHEET_FUNCTIONS.has(upper) ? '_xlfn._xlws.' : '_xlfn.'}${name}`
    }),
  )
}

/** Apply `transform` to the parts of a formula outside "strings" and 'sheet names'. */
function mapOutsideLiterals(formula: string, transform: (segment: string) => string): string {
  let result = ''
  let segment = ''
  let index = 0
  while (index < formula.length) {
    const char = formula[index]
    if (char === '"' || char === "'") {
      result += transform(segment)
      segment = ''
      let end = index + 1
      while (end < formula.length) {
        if (formula[end] === char) {
          if (formula[end + 1] === char) end += 2 // doubled quote is an escaped quote
          else break
        } else end += 1
      }
      result += formula.slice(index, end + 1)
      index = end + 1
    } else {
      segment += char
      index += 1
    }
  }
  return result + transform(segment)
}

function translateSharedFormula(formula: string, rowDelta: number, columnDelta: number): string {
  return mapOutsideLiterals(formula, segment =>
    segment.replace(
      /(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(!.])/gi,
      (token, absoluteColumn: string, column: string, absoluteRow: string, row: string) => {
        const position = parseCellKey(`${column}${row}`)
        if (!position) return token
        const nextRow = position.row + (absoluteRow ? 0 : rowDelta)
        const nextColumn = position.column + (absoluteColumn ? 0 : columnDelta)
        if (nextRow < 0 || nextColumn < 0) return '#REF!'
        return `${absoluteColumn}${columnLabel(nextColumn)}${absoluteRow}${nextRow + 1}`
      },
    ),
  )
}

// ---------------------------------------------------------------------------
// Sizes: Excel widths are in characters of the default font (7px wide at
// 100%, plus 5px of padding); heights are in points.
// ---------------------------------------------------------------------------

function excelWidthToPixels(width: number): number {
  return clamp(Math.round(width * 7 + 5), 16, 2000)
}

function pixelsToExcelWidth(pixels: number): number {
  return Math.max(0.5, Math.round(((pixels - 5) / 7) * 100) / 100)
}

function pointsToPixels(points: number): number {
  return clamp(Math.round((points * 4) / 3), 8, 800)
}

function pixelsToPoints(pixels: number): number {
  return Math.round(pixels * 0.75 * 100) / 100
}

// ---------------------------------------------------------------------------
// Parts exceljs reads past: charts, pivot tables, macros and the like are
// found by name in the zip's directory, so the notice says what was there.
// ---------------------------------------------------------------------------

export function listZipEntries(bytes: Uint8Array): string[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const floor = Math.max(0, bytes.length - 65_557)
  let end = -1
  for (let offset = bytes.length - 22; offset >= floor; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      end = offset
      break
    }
  }
  if (end < 0) return []
  const count = view.getUint16(end + 10, true)
  let offset = view.getUint32(end + 16, true)
  const decoder = new TextDecoder()
  const names: string[] = []
  for (let index = 0; index < count && offset + 46 <= bytes.length; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) break
    const nameLength = view.getUint16(offset + 28, true)
    const extraLength = view.getUint16(offset + 30, true)
    const commentLength = view.getUint16(offset + 32, true)
    names.push(decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength)))
    offset += 46 + nameLength + extraLength + commentLength
  }
  return names
}

export function unsupportedParts(entries: string[]): string[] {
  const count = (pattern: RegExp) => entries.filter(entry => pattern.test(entry)).length
  const parts: string[] = []
  const charts = count(/^xl\/charts\/chart\d+\.xml$/)
  if (charts) parts.push(`${charts} chart${charts === 1 ? '' : 's'} (left out; make them again with Chart)`)
  const pivots = count(/^xl\/pivotTables\/pivotTable\d+\.xml$/)
  if (pivots)
    parts.push(`${pivots} pivot table${pivots === 1 ? ' (its' : 's (their'} last values came in as plain cells)`)
  if (count(/^xl\/vbaProject\.bin$/)) parts.push('macros (left out)')
  const slicers = count(/^xl\/(slicers|timelines)\//)
  if (slicers) parts.push('slicers and timelines (left out)')
  const links = count(/^xl\/externalLinks\/externalLink\d+\.xml$/)
  if (links) parts.push(`links to ${links} other workbook${links === 1 ? '' : 's'} (formulas that use them will show #REF!)`)
  if (count(/^xl\/drawings\/drawing\d+\.xml$/) && !charts)
    parts.push('drawn shapes and text boxes (left out)')
  return parts
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------

const ERROR_LITERALS = new Set(['#N/A', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!'])
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d{1,3})?)?Z?)?$/

/** Write the whole workbook as .xlsx bytes, with a plain list of what Excel will not get. */
export async function exportXlsxWorkbook(
  document: PureSheetsDocument,
): Promise<XlsxExportResult> {
  const workbook = new ExcelJS.Workbook()
  workbook.created = new Date()
  // Excel recalculates on open, so a function PureSheets cannot evaluate still shows Excel's answer.
  workbook.calcProperties = { fullCalcOnLoad: true }
  const losses = new LossCounter()
  // Agent draft sheets are internal review state, not export content.
  const exported = document.workbook.sheets.filter(sheet => !sheet.agentDraft)
  const names = excelSheetNames(exported.map(sheet => sheet.name))
  for (const [from, to] of names)
    if (from !== to) losses.add(`sheet "${from}" is named "${to}" (Excel does not allow that name; formulas follow it)`)
  for (const sheet of exported) {
    const worksheet = workbook.addWorksheet(names.get(sheet.name) ?? sheet.name)
    const sheetLosses = new LossCounter()
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (!parseCellKey(key)) continue
      const excelCell = worksheet.getCell(key)
      excelCell.value = excelValue(sheet, key, cell, document, names)
      applyStyleToExcelCell(excelCell, cell.style)
    }
    for (const [key, text] of Object.entries(sheet.comments ?? {})) {
      if (parseCellKey(key) && text) worksheet.getCell(key).note = text
    }
    for (const merge of sheet.merges ?? []) {
      try {
        worksheet.mergeCellsWithoutStyle(merge)
      } catch {
        sheetLosses.add(`overlapping merge ${merge} (left unmerged)`)
      }
    }
    for (const [column, width] of Object.entries(sheet.columnWidths ?? {})) {
      worksheet.getColumn(Number(column) + 1).width = pixelsToExcelWidth(width)
    }
    for (const [row, height] of Object.entries(sheet.rowHeights ?? {})) {
      worksheet.getRow(Number(row) + 1).height = pixelsToPoints(height)
    }
    if (sheet.frozenRows || sheet.frozenColumns) {
      worksheet.views = [
        { state: 'frozen', ySplit: sheet.frozenRows ?? 0, xSplit: sheet.frozenColumns ?? 0 },
      ]
    }
    if (sheet.charts?.length) sheetLosses.add('charts (Excel will not get them)', sheet.charts.length)
    if (sheet.validations?.length) sheetLosses.add('data-validation rules', sheet.validations.length)
    if (sheet.conditionalFormats?.length)
      sheetLosses.add('conditional-formatting rules', sheet.conditionalFormats.length)
    if (sheet.protection?.lockedRanges.length) sheetLosses.add('locked ranges (cells export unlocked)')
    if (sheet.filterRow !== undefined) sheetLosses.add('the filter')
    losses.addSection(sheet.name, sheetLosses)
  }
  const buffer = await workbook.xlsx.writeBuffer()
  return { bytes: new Uint8Array(buffer as ArrayBuffer), warnings: losses.lines() }
}

/**
 * Excel's rules for sheet names: 31 characters, none of \ / ? * : [ ], no
 * leading or trailing apostrophe, not "History", unique ignoring case. Each
 * name maps to itself when it already fits.
 */
export function excelSheetNames(names: string[]): Map<string, string> {
  const result = new Map<string, string>()
  const taken = new Set<string>()
  for (const name of names) {
    const base =
      (name || 'Sheet').replace(/[\\/?*:[\]]/g, '-').replace(/^'+|'+$/g, '').slice(0, 31).trim() || 'Sheet'
    let candidate = base.toLowerCase() === 'history' ? `${base} sheet` : base
    for (let suffix = 2; taken.has(candidate.toLowerCase()); suffix += 1)
      candidate = `${base.slice(0, 31 - String(suffix).length - 1)} ${suffix}`
    taken.add(candidate.toLowerCase())
    result.set(name, candidate)
  }
  return result
}

/** Point sheet-qualified references (Costs!A1, 'Q1/Q2'!A1) at the sheet's Excel name. */
function renameSheetReferences(formula: string, names: Map<string, string>): string {
  const renamed = new Map(
    [...names].filter(([from, to]) => from !== to).map(([from, to]) => [from.toLowerCase(), to]),
  )
  if (!renamed.size) return formula
  const quote = (name: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, index) =>
      index % 2 === 1
        ? part
        : part.replace(/'((?:[^']|'')+)'!|(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)!/g, (whole, quoted?: string, bare?: string) => {
            const name = quoted !== undefined ? quoted.replace(/''/g, "'") : bare ?? ''
            const target = renamed.get(name.toLowerCase())
            return target ? `${quote(target)}!` : whole
          }),
    )
    .join('')
}

function excelValue(
  sheet: WorkbookSheet,
  key: string,
  cell: SheetCell,
  document: PureSheetsDocument,
  names: Map<string, string>,
): ExcelJS.CellValue {
  const value = cell.value
  if (cell.kind === 'formula') {
    return {
      formula: formulaToExcel(renameSheetReferences(value.replace(/^=/, ''), names)),
      result: cachedResult(computedCellValue(sheet, key, document.workbook)),
    } as ExcelJS.CellFormulaValue
  }
  if (!value) return null
  if (
    cell.kind === 'number' &&
    Number.isFinite(Number(value)) &&
    (!Number.isInteger(Number(value)) || Number.isSafeInteger(Number(value)))
  )
    return Number(value)
  if (cell.style?.numberFormat === 'text') return value
  if (value === 'TRUE' || value === 'FALSE') return value === 'TRUE'
  if (ERROR_LITERALS.has(value)) return { error: value } as ExcelJS.CellErrorValue
  if (cell.style?.numberFormat === 'date' && ISO_DATE.test(value)) {
    const date = new Date(value.length === 10 ? `${value}T00:00:00.000Z` : value.endsWith('Z') ? value : `${value}Z`)
    if (!Number.isNaN(date.getTime())) return date
  }
  return value
}

function cachedResult(value: string): ExcelJS.CellFormulaValue['result'] {
  if (ERROR_LITERALS.has(value)) return { error: value } as ExcelJS.CellErrorValue
  const numeric = Number(value)
  if (value.trim() !== '' && Number.isFinite(numeric)) return numeric
  if (value === 'TRUE' || value === 'FALSE') return value === 'TRUE'
  return value
}

function applyStyleToExcelCell(excelCell: ExcelJS.Cell, style: CellStyle | undefined): void {
  if (!style) return
  const font: Partial<ExcelJS.Font> = {}
  if (style.bold) font.bold = true
  if (style.italic) font.italic = true
  if (style.underline) font.underline = true
  if (style.fontSize) font.size = style.fontSize
  if (style.textColor) font.color = { argb: hexToArgb(style.textColor) }
  if (Object.keys(font).length) excelCell.font = font as ExcelJS.Font
  if (style.fillColor) {
    excelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(style.fillColor) } }
  }
  const sides = style.border ? (['top', 'right', 'bottom', 'left'] as const) : style.borderSides ?? []
  if (sides.length) {
    const line: Partial<ExcelJS.Border> = { style: 'thin' }
    excelCell.border = Object.fromEntries(sides.map(side => [side, line])) as Partial<ExcelJS.Borders>
  }
  const alignment: Partial<ExcelJS.Alignment> = {}
  if (style.align) alignment.horizontal = style.align
  if (style.verticalAlign) alignment.vertical = style.verticalAlign
  if (style.wrap === 'wrap') alignment.wrapText = true
  if (Object.keys(alignment).length) excelCell.alignment = alignment
  const pattern = patternOfStyle(style)
  // PureSheets' plain number (digits as entered) is Excel's General.
  if (pattern && !(style.numberFormat === 'number' && !style.numberPattern && style.decimals === undefined))
    excelCell.numFmt = pattern
}

function hexToArgb(hex: string): string {
  return `FF${hex.replace('#', '').toUpperCase()}`
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

// ---------------------------------------------------------------------------
// Base64 helpers for the binary bridge (renderer-safe, chunked).
// ---------------------------------------------------------------------------

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunkSize = 0x8000
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize))
  }
  return btoa(binary)
}
