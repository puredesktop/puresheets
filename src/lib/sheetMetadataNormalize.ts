import type {
  CellStyle,
  ColumnFilterRule,
  ConditionalFormatRule,
  NamedRange,
  PivotAggregate,
  PivotDefinition,
  SheetProtection,
  ValidationRule,
} from '../types'

const VALIDATION_KINDS = new Set<ValidationRule['kind']>([
  'list',
  'number-range',
  'date',
  'text-length',
  'text-contains',
  'custom-formula',
])

const CONDITION_KINDS = new Set<ConditionalFormatRule['condition']['kind']>([
  'greater',
  'less',
  'equal',
  'not-equal',
  'between',
  'text-contains',
  'blank',
  'not-blank',
  'duplicate',
  'top-n',
  'bottom-n',
  'color-scale',
  'data-bar',
  'custom-formula',
])

const PIVOT_AGGREGATES = new Set<PivotAggregate>([
  'sum',
  'count',
  'avg',
  'min',
  'max',
  'percent-of-total',
])

/**
 * Trust-boundary normalizers for the Phase S3 sheet metadata (comments,
 * validation rules, protection, conditional formats, filter query). Shared by
 * the legacy JSON loader (`sheetsDocument.ts`) and the Univer snapshot
 * restore path (`univerAdapter.ts`) so both produce identical shapes.
 */

export const MAX_SHEET_ROWS = 1_048_576
export const MAX_SHEET_COLUMNS = 16_384
export const MAX_PERSISTED_ROWS = 100_000

const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype'])

/**
 * Strict hex color validator/normalizer (`#RGB` → `#rrggbb`, `#RRGGBB` → `#rrggbb`).
 * Rejects CSS injection payloads, rgb()/var()/url(), named colors, etc.
 */
export function normalizeHexColor(input: unknown): string | undefined {
  if (typeof input !== 'string') return undefined
  const trimmed = input.trim()
  const match = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(trimmed)
  if (!match) return undefined
  const hex = match[1].toLowerCase()
  if (hex.length === 3) {
    return `#${hex[0]}${hex[0]}${hex[1]}${hex[1]}${hex[2]}${hex[2]}`
  }
  return `#${hex}`
}

export function clampFontSize(input: unknown): number | undefined {
  if (typeof input !== 'number' || !Number.isFinite(input)) return undefined
  return Math.min(96, Math.max(6, input))
}

export function isForbiddenKey(key: string): boolean {
  return FORBIDDEN_KEYS.has(key)
}

/** Validate an A1 cell key is syntactically valid and within Excel grid bounds. */
export function isSafeCellKey(input: string): boolean {
  const trimmed = input.trim()
  if (!trimmed || isForbiddenKey(trimmed)) return false
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/i.exec(trimmed)
  if (!match) return false
  const row = Number(match[2])
  if (!Number.isSafeInteger(row) || row < 1 || row > MAX_SHEET_ROWS) return false
  const col = match[1]
    .toUpperCase()
    .split('')
    .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0)
  return col >= 1 && col <= MAX_SHEET_COLUMNS
}

export function isSafeRangeToken(input: string): boolean {
  const trimmed = input.trim()
  if (!trimmed) return false
  const parts = trimmed.split(':')
  if (parts.length === 1) return isSafeCellKey(parts[0])
  if (parts.length === 2) return isSafeCellKey(parts[0]) && isSafeCellKey(parts[1])
  return false
}

export function normalizeFilterQuery(input: unknown): string | undefined {
  return typeof input === 'string' && input.trim() ? input.trim() : undefined
}

export function normalizeComments(
  input: unknown,
): Record<string, string> | undefined {
  if (!isRecord(input)) return undefined
  const comments: Record<string, string> = {}
  for (const [key, value] of Object.entries(input)) {
    if (typeof value !== 'string' || !value.trim()) continue
    const normalizedKey = key.trim().toUpperCase()
    if (!isSafeCellKey(normalizedKey)) continue
    comments[normalizedKey] = value
  }
  return Object.keys(comments).length ? comments : undefined
}

export function normalizeValidations(
  input: unknown,
): ValidationRule[] | undefined {
  if (!Array.isArray(input)) return undefined
  const rules = input.flatMap((item): ValidationRule[] => {
    if (!isRecord(item)) return []
    if (
      typeof item.id !== 'string' ||
      typeof item.range !== 'string' ||
      !isRangeToken(item.range) ||
      typeof item.kind !== 'string' ||
      !VALIDATION_KINDS.has(item.kind as ValidationRule['kind'])
    ) {
      return []
    }
    const rule: ValidationRule = {
      id: item.id,
      range: item.range,
      kind: item.kind as ValidationRule['kind'],
    }
    if (Array.isArray(item.options)) {
      rule.options = item.options.filter(
        (option): option is string => typeof option === 'string',
      )
    }
    if (typeof item.sourceRange === 'string' && item.sourceRange.trim()) {
      rule.sourceRange = item.sourceRange.trim()
    }
    if (typeof item.min === 'number' && Number.isFinite(item.min)) {
      rule.min = item.min
    }
    if (typeof item.max === 'number' && Number.isFinite(item.max)) {
      rule.max = item.max
    }
    if (item.whole === true) {
      rule.whole = true
    }
    if (typeof item.textQuery === 'string' && item.textQuery.trim()) {
      rule.textQuery = item.textQuery
    }
    if (typeof item.formula === 'string' && item.formula.trim()) {
      rule.formula = item.formula
    }
    if (typeof item.strict === 'boolean') {
      rule.strict = item.strict
    }
    return [rule]
  })
  return rules.length ? rules : undefined
}

export function normalizeProtection(
  input: unknown,
): SheetProtection | undefined {
  if (!isRecord(input) || !Array.isArray(input.lockedRanges)) return undefined
  const lockedRanges = input.lockedRanges.filter(
    (range): range is string => typeof range === 'string' && isRangeToken(range),
  )
  return lockedRanges.length ? { lockedRanges } : undefined
}

export function normalizeConditionalFormats(
  input: unknown,
): ConditionalFormatRule[] | undefined {
  if (!Array.isArray(input)) return undefined
  const rules = input.flatMap((item): ConditionalFormatRule[] => {
    if (!isRecord(item)) return []
    if (
      typeof item.id !== 'string' ||
      typeof item.range !== 'string' ||
      !isRangeToken(item.range) ||
      !isRecord(item.condition)
    ) {
      return []
    }
    const condition = item.condition
    if (
      typeof condition.kind !== 'string' ||
      !CONDITION_KINDS.has(
        condition.kind as ConditionalFormatRule['condition']['kind'],
      ) ||
      typeof condition.value !== 'string'
    ) {
      return []
    }
    const kind = condition.kind as ConditionalFormatRule['condition']['kind']
    const style = normalizeStyleObject(item.style) ?? {}
    if (kind !== 'color-scale' && kind !== 'data-bar' && !Object.keys(style).length) {
      return []
    }
    const rule: ConditionalFormatRule = {
      id: item.id,
      range: item.range,
      condition: {
        kind,
        value: condition.value,
        value2:
          typeof condition.value2 === 'string' ? condition.value2 : undefined,
      },
      style,
    }
    if (
      isRecord(item.colorScale) &&
      typeof item.colorScale.minColor === 'string' &&
      typeof item.colorScale.maxColor === 'string'
    ) {
      rule.colorScale = {
        minColor: item.colorScale.minColor,
        midColor:
          typeof item.colorScale.midColor === 'string'
            ? item.colorScale.midColor
            : undefined,
        maxColor: item.colorScale.maxColor,
      }
    }
    if (typeof item.dataBarColor === 'string' && item.dataBarColor.trim()) {
      rule.dataBarColor = item.dataBarColor
    }
    return [rule]
  })
  return rules.length ? rules : undefined
}

export function normalizeColumnFilters(
  input: unknown,
): Record<string, ColumnFilterRule> | undefined {
  if (!isRecord(input)) return undefined
  const out: Record<string, ColumnFilterRule> = {}
  for (const [col, raw] of Object.entries(input)) {
    if (!/^\d+$/.test(col) || !isRecord(raw)) continue
    const rule: ColumnFilterRule = {}
    if (Array.isArray(raw.selectedValues)) {
      rule.selectedValues = raw.selectedValues.filter(
        (v): v is string => typeof v === 'string',
      )
    }
    if (typeof raw.operator === 'string') {
      rule.operator = raw.operator as ColumnFilterRule['operator']
    }
    if (typeof raw.value === 'string') rule.value = raw.value
    if (typeof raw.value2 === 'string') rule.value2 = raw.value2
    if (Object.keys(rule).length) out[col] = rule
  }
  return Object.keys(out).length ? out : undefined
}

export function normalizePivotDefinition(
  input: unknown,
): PivotDefinition | undefined {
  if (!isRecord(input)) return undefined
  if (
    typeof input.id !== 'string' ||
    typeof input.sourceSheetId !== 'string' ||
    typeof input.sourceRange !== 'string' ||
    !Array.isArray(input.rows) ||
    !Array.isArray(input.values)
  ) {
    return undefined
  }
  const parseFields = (arr: unknown[]) =>
    arr.flatMap(item =>
      isRecord(item) &&
      typeof item.columnIndex === 'number' &&
      typeof item.label === 'string'
        ? [{ columnIndex: item.columnIndex, label: item.label }]
        : [],
    )
  const values = input.values.flatMap(item =>
    isRecord(item) &&
    typeof item.columnIndex === 'number' &&
    typeof item.label === 'string' &&
    typeof item.aggregate === 'string' &&
    PIVOT_AGGREGATES.has(item.aggregate as PivotAggregate)
      ? [
          {
            columnIndex: item.columnIndex,
            label: item.label,
            aggregate: item.aggregate as PivotAggregate,
          },
        ]
      : [],
  )
  const filters = Array.isArray(input.filters)
    ? input.filters.flatMap(item =>
        isRecord(item) &&
        typeof item.columnIndex === 'number' &&
        typeof item.label === 'string' &&
        Array.isArray(item.selectedValues)
          ? [
              {
                columnIndex: item.columnIndex,
                label: item.label,
                selectedValues: item.selectedValues.filter(
                  (v): v is string => typeof v === 'string',
                ),
              },
            ]
          : [],
      )
    : undefined
  return {
    id: input.id,
    sourceSheetId: input.sourceSheetId,
    sourceRange: input.sourceRange,
    rows: parseFields(input.rows),
    columns: Array.isArray(input.columns) ? parseFields(input.columns) : [],
    values,
    filters: filters?.length ? filters : undefined,
    refreshedAt:
      typeof input.refreshedAt === 'string'
        ? input.refreshedAt
        : new Date().toISOString(),
    sourceSignature:
      typeof input.sourceSignature === 'string'
        ? input.sourceSignature
        : undefined,
    output:
      typeof input.output === 'string' && isSafeCellKey(input.output)
        ? input.output.trim().toUpperCase()
        : undefined,
    outputRows:
      typeof input.outputRows === 'number' &&
      Number.isInteger(input.outputRows) &&
      input.outputRows >= 1 &&
      input.outputRows <= 25_000
        ? input.outputRows
        : undefined,
    outputColumns:
      typeof input.outputColumns === 'number' &&
      Number.isInteger(input.outputColumns) &&
      input.outputColumns >= 1 &&
      input.outputColumns <= 25_000
        ? input.outputColumns
        : undefined,
  }
}

export function normalizeNamedRanges(input: unknown): NamedRange[] | undefined {
  if (!Array.isArray(input)) return undefined
  const out = input.flatMap((item): NamedRange[] => {
    if (!isRecord(item)) return []
    if (
      typeof item.id !== 'string' ||
      typeof item.name !== 'string' ||
      !item.name.trim() ||
      typeof item.sheetId !== 'string' ||
      typeof item.range !== 'string' ||
      !isRangeToken(item.range)
    ) {
      return []
    }
    return [
      {
        id: item.id,
        name: item.name.trim(),
        sheetId: item.sheetId,
        range: item.range,
        comment: typeof item.comment === 'string' ? item.comment : undefined,
      },
    ]
  })
  return out.length ? out : undefined
}

export function normalizeIndexList(input: unknown): number[] | undefined {
  if (!Array.isArray(input)) return undefined
  const list = [...new Set(input.filter((n): n is number => typeof n === 'number' && Number.isInteger(n) && n >= 0))].sort((a, b) => a - b)
  return list.length ? list : undefined
}

/** Style validator used for conditional-format styles. */
export function normalizeStyleObject(input: unknown): CellStyle | undefined {
  if (!isRecord(input)) return undefined
  const style: CellStyle = {}
  if (typeof input.bold === 'boolean') style.bold = input.bold
  if (typeof input.italic === 'boolean') style.italic = input.italic
  if (typeof input.underline === 'boolean') style.underline = input.underline
  if (typeof input.strikethrough === 'boolean')
    style.strikethrough = input.strikethrough
  const textColor = normalizeHexColor(input.textColor)
  if (textColor) style.textColor = textColor
  const fillColor = normalizeHexColor(input.fillColor)
  if (fillColor) style.fillColor = fillColor
  const fontSize = clampFontSize(input.fontSize)
  if (fontSize !== undefined) style.fontSize = fontSize
  if (typeof input.border === 'boolean') style.border = input.border
  return Object.keys(style).length ? style : undefined
}

function isRangeToken(input: string): boolean {
  return isSafeRangeToken(input)
}

function isRecord(input: unknown): input is Record<string, unknown> {
  return typeof input === 'object' && input !== null && !Array.isArray(input)
}
