import type {
  CellStyle,
  ConditionalFormatRule,
  SheetProtection,
  ValidationRule,
} from '../types'

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
      (item.kind !== 'list' &&
        item.kind !== 'number-range' &&
        item.kind !== 'date')
    ) {
      return []
    }
    const rule: ValidationRule = {
      id: item.id,
      range: item.range,
      kind: item.kind,
    }
    if (Array.isArray(item.options)) {
      rule.options = item.options.filter(
        (option): option is string => typeof option === 'string',
      )
    }
    if (typeof item.min === 'number' && Number.isFinite(item.min)) {
      rule.min = item.min
    }
    if (typeof item.max === 'number' && Number.isFinite(item.max)) {
      rule.max = item.max
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
      !isRecord(item.condition) ||
      !isRecord(item.style)
    ) {
      return []
    }
    const condition = item.condition
    if (
      (condition.kind !== 'greater' &&
        condition.kind !== 'less' &&
        condition.kind !== 'equal' &&
        condition.kind !== 'between' &&
        condition.kind !== 'text-contains') ||
      typeof condition.value !== 'string'
    ) {
      return []
    }
    const style = normalizeStyleObject(item.style)
    if (!style) return []
    return [
      {
        id: item.id,
        range: item.range,
        condition: {
          kind: condition.kind,
          value: condition.value,
          value2:
            typeof condition.value2 === 'string' ? condition.value2 : undefined,
        },
        style,
      },
    ]
  })
  return rules.length ? rules : undefined
}

/** Style validator used for conditional-format styles. */
export function normalizeStyleObject(input: unknown): CellStyle | undefined {
  if (!isRecord(input)) return undefined
  const style: CellStyle = {}
  if (typeof input.bold === 'boolean') style.bold = input.bold
  if (typeof input.italic === 'boolean') style.italic = input.italic
  if (typeof input.underline === 'boolean') style.underline = input.underline
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
