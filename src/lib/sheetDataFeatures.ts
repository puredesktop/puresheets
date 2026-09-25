import {
  cellsInRange,
  computedCellValue,
  normalizeRangeToken,
  parseCellKey,
} from './workbookModel'
import type {
  CellStyle,
  ConditionalFormatRule,
  ValidationRule,
  WorkbookSheet,
} from '../types'

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
  const cells = new Set(cellsInRange(rangeToken))
  if (!cells.size) return sheet
  const remaining = (sheet.validations ?? []).filter(
    rule => !cellsInRange(rule.range).some(key => cells.has(key)),
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
  return sheet.validations.find(rule =>
    cellsInRange(rule.range).includes(key),
  )
}

export interface ValidationResult {
  valid: boolean
  message?: string
}

/** Validate one raw entry against a rule. Blank values are always valid. */
export function validateValue(
  rule: ValidationRule,
  value: string,
): ValidationResult {
  const trimmed = value.trim()
  if (!trimmed) return { valid: true }
  if (rule.kind === 'list') {
    const options = rule.options ?? []
    if (options.includes(trimmed)) return { valid: true }
    return {
      valid: false,
      message: `Value is not in the allowed list (${options
        .slice(0, 5)
        .join(', ')}${options.length > 5 ? ', …' : ''})`,
    }
  }
  if (rule.kind === 'number-range') {
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
  // date
  const parsed = new Date(trimmed)
  if (Number.isNaN(parsed.getTime())) {
    return { valid: false, message: 'Value must be a date' }
  }
  return { valid: true }
}

/** Flag map for the whole sheet: cell key → validation message. */
export function invalidCellsForSheet(
  sheet: WorkbookSheet,
): Map<string, string> {
  const flagged = new Map<string, string>()
  for (const rule of sheet.validations ?? []) {
    for (const key of cellsInRange(rule.range)) {
      if (flagged.has(key)) continue
      const value = computedCellValue(sheet, key)
      const result = validateValue(rule, value)
      if (!result.valid) {
        flagged.set(key, result.message ?? 'Invalid value')
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
): InvalidCellFlag[] {
  return [...invalidCellsForSheet(sheet)].map(([cell, message]) => ({
    cell,
    value: computedCellValue(sheet, cell),
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
  const cells = new Set(cellsInRange(rangeToken))
  if (!cells.size) return sheet
  const remaining = (sheet.conditionalFormats ?? []).filter(
    rule => !cellsInRange(rule.range).some(key => cells.has(key)),
  )
  if (remaining.length === (sheet.conditionalFormats ?? []).length) {
    return sheet
  }
  return {
    ...sheet,
    conditionalFormats: remaining.length ? remaining : undefined,
  }
}

export function conditionMatches(
  condition: ConditionalFormatRule['condition'],
  value: string,
): boolean {
  const trimmed = value.trim()
  if (!trimmed) return false
  if (condition.kind === 'text-contains') {
    return trimmed.toLowerCase().includes(condition.value.trim().toLowerCase())
  }
  const numeric = Number(trimmed)
  const target = Number(condition.value)
  if (condition.kind === 'equal') {
    if (Number.isFinite(numeric) && Number.isFinite(target)) {
      return numeric === target
    }
    return trimmed.toLowerCase() === condition.value.trim().toLowerCase()
  }
  if (!Number.isFinite(numeric) || !Number.isFinite(target)) return false
  if (condition.kind === 'greater') return numeric > target
  if (condition.kind === 'less') return numeric < target
  // between (inclusive)
  const upper = Number(condition.value2)
  if (!Number.isFinite(upper)) return false
  const low = Math.min(target, upper)
  const high = Math.max(target, upper)
  return numeric >= low && numeric <= high
}

/**
 * The conditional style for one cell: matching rules merge in order, later
 * rules winning per property. Evaluated on the computed (formula-result)
 * value, not the formatted display string.
 */
export function conditionalStyleForCell(
  sheet: WorkbookSheet,
  key: string,
): CellStyle | undefined {
  const rules = sheet.conditionalFormats
  if (!rules?.length) return undefined
  let merged: CellStyle | undefined
  for (const rule of rules) {
    if (!cellsInRange(rule.range).includes(key)) continue
    const value = computedCellValue(sheet, key)
    if (!conditionMatches(rule.condition, value)) continue
    merged = { ...merged, ...rule.style }
  }
  return merged
}
