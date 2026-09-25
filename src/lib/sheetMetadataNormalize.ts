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
    if (!isCellKey(key)) continue
    comments[key] = value
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
  if (typeof input.textColor === 'string') style.textColor = input.textColor
  if (typeof input.fillColor === 'string') style.fillColor = input.fillColor
  if (typeof input.border === 'boolean') style.border = input.border
  return Object.keys(style).length ? style : undefined
}

function isCellKey(input: string): boolean {
  return /^[A-Z]+[1-9][0-9]*$/i.test(input.trim())
}

function isRangeToken(input: string): boolean {
  return /^[A-Z]+[1-9][0-9]*(?::[A-Z]+[1-9][0-9]*)?$/i.test(input.trim())
}

function isRecord(input: unknown): input is Record<string, unknown> {
  return typeof input === 'object' && input !== null && !Array.isArray(input)
}
