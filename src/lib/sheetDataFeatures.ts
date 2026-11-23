import {
  boundsIntersect,
  cellsInRange,
  computedCellValue,
  evaluateFormula,
  normalizeRangeToken,
  parseCellKey,
  parseRangeBounds,
  type FormulaWorkbook,
} from './workbookModel'
import type {
  CellPosition,
  CellStyle,
  ConditionalFormatRule,
  ValidationRule,
  WorkbookSheet,
} from '../types'

function positionInBounds(
  position: CellPosition,
  bounds: { start: CellPosition; end: CellPosition },
): boolean {
  return (
    position.row >= bounds.start.row &&
    position.row <= bounds.end.row &&
    position.column >= bounds.start.column &&
    position.column <= bounds.end.column
  )
}

function cellInRange(rangeToken: string, position: CellPosition): boolean {
  const bounds = parseRangeBounds(rangeToken)
  return Boolean(bounds && positionInBounds(position, bounds))
}

function rangesIntersect(aToken: string, bToken: string): boolean {
  const a = parseRangeBounds(aToken)
  const b = parseRangeBounds(bToken)
  return Boolean(a && b && boundsIntersect(a, b))
}

/**
 * Data validation and conditional formatting (Phase S3).
 *
 * Both are model-layer rule sets stored on the sheet and persisted through
 * the snapshot bridge (`custom.puresheets`). Validation FLAGS invalid entries
 * — it never rejects them (per plan: "invalid entries flagged, not silently
 * rejected"). Conditional formats are evaluated against computed (formula
 * result) values at render time and merged over the cell's own style.
 */

let ruleIdCounter = 1
function nextRuleId(prefix: string): string {
  ruleIdCounter += 1
  return `${prefix}-${Date.now()}-${ruleIdCounter}`
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

export function addValidationRule(
  sheet: WorkbookSheet,
  rule: Omit<ValidationRule, 'id'>,
): WorkbookSheet {
  const range = normalizeRangeToken(rule.range)
  if (!range) return sheet
  // One rule per range: replace rules with the identical range.
  const remaining = (sheet.validations ?? []).filter(
    existing => existing.range !== range,
  )
  return {
    ...sheet,
    validations: [
      ...remaining,
      { ...rule, range, id: nextRuleId('validation') },
    ],
  }
}

/** Remove every validation rule whose range intersects the given range. */
export function clearValidationInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  if (!parseRangeBounds(rangeToken)) return sheet
  const remaining = (sheet.validations ?? []).filter(
    rule => !rangesIntersect(rule.range, rangeToken),
  )
  if (remaining.length === (sheet.validations ?? []).length) return sheet
  return {
    ...sheet,
    validations: remaining.length ? remaining : undefined,
  }
}

export function validationRuleForCell(
  sheet: WorkbookSheet,
  key: string,
): ValidationRule | undefined {
  if (!sheet.validations?.length) return undefined
  const position = parseCellKey(key)
  if (!position) return undefined
  return sheet.validations.find(rule => cellInRange(rule.range, position))
}

/** Remove a single validation rule by ID. */
export function removeValidationRule(
  sheet: WorkbookSheet,
  ruleId: string,
): WorkbookSheet {
  const remaining = (sheet.validations ?? []).filter(r => r.id !== ruleId)
  if (remaining.length === (sheet.validations ?? []).length) return sheet
  return {
    ...sheet,
    validations: remaining.length ? remaining : undefined,
  }
}

/** Resolve allowed options for a list validation rule (explicit options or range reference). */
export function resolveValidationOptions(
  sheet: WorkbookSheet,
  rule: ValidationRule,
  workbook?: FormulaWorkbook,
): string[] {
  if (rule.kind !== 'list') return []
  if (rule.sourceRange?.trim()) {
    const raw = rule.sourceRange.trim().replace(/^=/, '')
    let targetSheet = sheet
    let rangePart = raw
    const bangIdx = raw.indexOf('!')
    if (bangIdx > 0 && workbook) {
      const sheetName = raw.slice(0, bangIdx).replace(/^'|'$/g, '').trim()
      const found = workbook.sheets.find(
        s => s.name.toLowerCase() === sheetName.toLowerCase(),
      )
      if (found) targetSheet = found
      rangePart = raw.slice(bangIdx + 1)
    }
    const cleanRange = normalizeRangeToken(rangePart.replace(/\$/g, ''))
    if (cleanRange) {
      const vals = cellsInRange(cleanRange, 1000)
        .map(k => computedCellValue(targetSheet, k, workbook).trim())
        .filter(Boolean)
      return [...new Set(vals)]
    }
  }
  return rule.options ?? []
}

export interface ValidationResult {
  valid: boolean
  message?: string
}

/** Validate one raw entry against a rule. Blank values are always valid. */
export function validateValue(
  rule: ValidationRule,
  value: string,
  sheet?: WorkbookSheet,
  workbook?: FormulaWorkbook,
): ValidationResult {
  const trimmed = value.trim()
  if (!trimmed) return { valid: true }
  if (rule.kind === 'list') {
    const options = sheet
      ? resolveValidationOptions(sheet, rule, workbook)
      : rule.options ?? []
    if (options.includes(trimmed)) return { valid: true }
    return {
      valid: false,
      message: `Value is not in the allowed list (${options
        .slice(0, 5)
        .join(', ')}${options.length > 5 ? ', …' : ''})`,
    }
  }
  if (rule.kind === 'number-range') {
    if (rule.whole && !Number.isInteger(Number(value))) return { valid: false, message: 'Enter a whole number.' }
    const numeric = Number(trimmed)
    if (!Number.isFinite(numeric)) {
      return { valid: false, message: 'Value must be a number' }
    }
    if (rule.min !== undefined && numeric < rule.min) {
      return { valid: false, message: `Value must be at least ${rule.min}` }
    }
    if (rule.max !== undefined && numeric > rule.max) {
      return { valid: false, message: `Value must be at most ${rule.max}` }
    }
    return { valid: true }
  }
  if (rule.kind === 'text-length') {
    const len = trimmed.length
    if (rule.min !== undefined && len < rule.min) {
      return { valid: false, message: `Length must be at least ${rule.min}` }
    }
    if (rule.max !== undefined && len > rule.max) {
      return { valid: false, message: `Length must be at most ${rule.max}` }
    }
    return { valid: true }
  }
  if (rule.kind === 'text-contains') {
    const q = (rule.textQuery ?? '').trim().toLowerCase()
    if (!q || trimmed.toLowerCase().includes(q)) return { valid: true }
    return { valid: false, message: `Value must contain "${rule.textQuery}"` }
  }
  if (rule.kind === 'custom-formula') {
    if (!rule.formula?.trim() || !sheet) return { valid: true }
    const expr = rule.formula.trim().startsWith('=')
      ? rule.formula.trim()
      : `=${rule.formula.trim()}`
    const res = evaluateFormula(sheet, expr, workbook).trim().toUpperCase()
    const ok = res === 'TRUE' || (Number.isFinite(Number(res)) && Number(res) !== 0)
    return ok ? { valid: true } : { valid: false, message: `Failed formula ${expr}` }
  }
  // date: reject bare numbers like "1" or "2025" that V8's Date() accepts
  if (/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) {
    return { valid: false, message: 'Value must be a date' }
  }
  const isoMatch = /^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/.exec(trimmed)
  if (isoMatch) {
    const year = Number(isoMatch[1])
    const month = Number(isoMatch[2])
    const day = Number(isoMatch[3])
    const check = new Date(Date.UTC(year, month - 1, day))
    if (
      check.getUTCFullYear() !== year ||
      check.getUTCMonth() !== month - 1 ||
      check.getUTCDate() !== day
    ) {
      return { valid: false, message: 'Value must be a date' }
    }
    return { valid: true }
  }
  if (!/[-/,]|\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i.test(trimmed)) {
    return { valid: false, message: 'Value must be a date' }
  }
  const parsed = new Date(trimmed)
  if (Number.isNaN(parsed.getTime())) {
    return { valid: false, message: 'Value must be a date' }
  }
  return { valid: true }
}

/** Flag map for the whole sheet: cell key → validation message. */
export function invalidCellsForSheet(
  sheet: WorkbookSheet,
  workbook?: FormulaWorkbook,
): Map<string, string> {
  const flagged = new Map<string, string>()
  const rules = (sheet.validations ?? [])
    .map(rule => ({ rule, bounds: parseRangeBounds(rule.range) }))
    .filter(
      (entry): entry is { rule: ValidationRule; bounds: NonNullable<ReturnType<typeof parseRangeBounds>> } =>
        entry.bounds !== null,
    )
  if (!rules.length) return flagged
  for (const key of Object.keys(sheet.cells)) {
    const position = parseCellKey(key)
    if (!position) continue
    for (const { rule, bounds } of rules) {
      if (!positionInBounds(position, bounds)) continue
      const value = computedCellValue(sheet, key, workbook)
      const result = validateValue(rule, value, sheet, workbook)
      if (!result.valid) {
        flagged.set(key, result.message ?? 'Invalid value')
        break
      }
    }
  }
  return flagged
}

/**
 * One flagged cell for the live editing surface: the cell key, its current
 * computed value (the surface paints via exact-match display rules that stop
 * matching the moment the value changes), and the human-readable reason.
 */
export interface InvalidCellFlag {
  cell: string
  value: string
  message: string
}

/**
 * Production validation feedback: every currently invalid cell with its
 * computed value and message. This is what the Univer surface paints and
 * what the fallback grid titles its cells with.
 */
export function invalidCellFlagsForSheet(
  sheet: WorkbookSheet,
  workbook?: FormulaWorkbook,
): InvalidCellFlag[] {
  return [...invalidCellsForSheet(sheet, workbook)].map(([cell, message]) => ({
    cell,
    value: computedCellValue(sheet, cell, workbook),
    message,
  }))
}

// ---------------------------------------------------------------------------
// Conditional formatting
// ---------------------------------------------------------------------------

export function addConditionalFormatRule(
  sheet: WorkbookSheet,
  rule: Omit<ConditionalFormatRule, 'id'>,
): WorkbookSheet {
  const range = normalizeRangeToken(rule.range)
  if (!range) return sheet
  return {
    ...sheet,
    conditionalFormats: [
      ...(sheet.conditionalFormats ?? []),
      { ...rule, range, id: nextRuleId('conditional-format') },
    ],
  }
}

/** Remove every conditional-format rule intersecting the given range. */
export function clearConditionalFormatsInRange(
  sheet: WorkbookSheet,
  rangeToken: string,
): WorkbookSheet {
  if (!parseRangeBounds(rangeToken)) return sheet
  const remaining = (sheet.conditionalFormats ?? []).filter(
    rule => !rangesIntersect(rule.range, rangeToken),
  )
  if (remaining.length === (sheet.conditionalFormats ?? []).length) {
    return sheet
  }
  return {
    ...sheet,
    conditionalFormats: remaining.length ? remaining : undefined,
  }
}

/** Remove a single conditional-format rule by ID. */
export function removeConditionalFormatRule(
  sheet: WorkbookSheet,
  ruleId: string,
): WorkbookSheet {
  const remaining = (sheet.conditionalFormats ?? []).filter(r => r.id !== ruleId)
  if (remaining.length === (sheet.conditionalFormats ?? []).length) return sheet
  return {
    ...sheet,
    conditionalFormats: remaining.length ? remaining : undefined,
  }
}

/** Reorder a conditional-format rule up (-1) or down (+1). */
export function moveConditionalFormatRule(
  sheet: WorkbookSheet,
  ruleId: string,
  delta: -1 | 1,
): WorkbookSheet {
  const rules = [...(sheet.conditionalFormats ?? [])]
  const index = rules.findIndex(r => r.id === ruleId)
  if (index < 0) return sheet
  const target = index + delta
  if (target < 0 || target >= rules.length) return sheet
  const [moved] = rules.splice(index, 1)
  rules.splice(target, 0, moved)
  return { ...sheet, conditionalFormats: rules }
}

export function conditionMatches(
  condition: ConditionalFormatRule['condition'],
  value: string,
): boolean {
  const trimmed = value.trim()
  if (condition.kind === 'blank') return trimmed === ''
  if (condition.kind === 'not-blank') return trimmed !== ''
  if (!trimmed) return false
  if (condition.kind === 'text-contains') {
    const needle = condition.value.trim().toLowerCase()
    if (!needle) return false
    return trimmed.toLowerCase().includes(needle)
  }
  const numeric = Number(trimmed)
  const target = Number(condition.value)
  if (condition.kind === 'equal') {
    if (Number.isFinite(numeric) && Number.isFinite(target) && condition.value.trim() !== '') {
      return numeric === target
    }
    return trimmed.toLowerCase() === condition.value.trim().toLowerCase()
  }
  if (condition.kind === 'not-equal') {
    if (Number.isFinite(numeric) && Number.isFinite(target) && condition.value.trim() !== '') {
      return numeric !== target
    }
    return trimmed.toLowerCase() !== condition.value.trim().toLowerCase()
  }
  if (!Number.isFinite(numeric) || !Number.isFinite(target) || condition.value.trim() === '') return false
  if (condition.kind === 'greater') return numeric > target
  if (condition.kind === 'less') return numeric < target
  if (condition.kind === 'between') {
    const upper = Number(condition.value2)
    if (!condition.value2?.trim() || !Number.isFinite(upper)) return false
    const low = Math.min(target, upper)
    const high = Math.max(target, upper)
    return numeric >= low && numeric <= high
  }
  return false
}

function interpolateHex(a: string, b: string, t: number): string {
  const clamp01 = Math.max(0, Math.min(1, t))
  const parse = (hex: string) => [
    parseInt(hex.slice(1, 3), 16) || 0,
    parseInt(hex.slice(3, 5), 16) || 0,
    parseInt(hex.slice(5, 7), 16) || 0,
  ]
  const [r1, g1, b1] = parse(a)
  const [r2, g2, b2] = parse(b)
  const mix = (x: number, y: number) =>
    Math.round(x + (y - x) * clamp01)
      .toString(16)
      .padStart(2, '0')
  return `#${mix(r1, r2)}${mix(g1, g2)}${mix(b1, b2)}`
}

/**
 * The conditional style for one cell: matching rules merge in order, later
 * rules winning per property. Evaluated on the computed (formula-result)
 * value, not the formatted display string.
 */
export function conditionalStyleForCell(
  sheet: WorkbookSheet,
  key: string,
  workbook?: FormulaWorkbook,
): CellStyle | undefined {
  const rules = sheet.conditionalFormats
  if (!rules?.length) return undefined
  const position = parseCellKey(key)
  if (!position) return undefined
  let merged: CellStyle | undefined
  for (const rule of rules) {
    if (!cellInRange(rule.range, position)) continue
    const rangeKeys = cellsInRange(rule.range, 10_000)
    const value = computedCellValue(sheet, key, workbook)
    const kind = rule.condition.kind

    if (kind === 'duplicate') {
      const norm = value.trim().toLowerCase()
      if (!norm) continue
      let count = 0
      for (const rk of rangeKeys) {
        if (computedCellValue(sheet, rk, workbook).trim().toLowerCase() === norm) {
          count += 1
          if (count > 1) break
        }
      }
      if (count > 1) merged = { ...merged, ...rule.style }
      continue
    }

    if (kind === 'top-n' || kind === 'bottom-n') {
      const n = Math.max(1, Math.floor(Number(rule.condition.value) || 3))
      const num = Number(value.trim())
      if (value.trim() === '' || !Number.isFinite(num)) continue
      const allNums = rangeKeys
        .map(rk => Number(computedCellValue(sheet, rk, workbook).trim()))
        .filter(Number.isFinite)
        .sort((a, b) => (kind === 'top-n' ? b - a : a - b))
      const cutoff = allNums[Math.min(n - 1, allNums.length - 1)]
      if (
        cutoff !== undefined &&
        (kind === 'top-n' ? num >= cutoff : num <= cutoff)
      ) {
        merged = { ...merged, ...rule.style }
      }
      continue
    }

    if (kind === 'color-scale' && rule.colorScale) {
      const num = Number(value.trim())
      if (value.trim() === '' || !Number.isFinite(num)) continue
      const allNums = rangeKeys
        .map(rk => Number(computedCellValue(sheet, rk, workbook).trim()))
        .filter(Number.isFinite)
      if (!allNums.length) continue
      const min = Math.min(...allNums)
      const max = Math.max(...allNums)
      const t = max === min ? 0.5 : (num - min) / (max - min)
      const fillColor = rule.colorScale.midColor
        ? t <= 0.5
          ? interpolateHex(rule.colorScale.minColor, rule.colorScale.midColor, t * 2)
          : interpolateHex(rule.colorScale.midColor, rule.colorScale.maxColor, (t - 0.5) * 2)
        : interpolateHex(rule.colorScale.minColor, rule.colorScale.maxColor, t)
      merged = { ...merged, fillColor }
      continue
    }

    if (kind === 'custom-formula') {
      const expr = rule.condition.value.trim()
      if (!expr) continue
      const res = evaluateFormula(
        sheet,
        expr.startsWith('=') ? expr : `=${expr}`,
        workbook,
      )
        .trim()
        .toUpperCase()
      if (res === 'TRUE' || (Number.isFinite(Number(res)) && Number(res) !== 0)) {
        merged = { ...merged, ...rule.style }
      }
      continue
    }

    if (!conditionMatches(rule.condition, value)) continue
    merged = { ...merged, ...rule.style }
  }
  return merged
}
