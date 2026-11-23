import type ExcelJS from 'exceljs'
import type { CellStyle, ConditionalFormatRule, ValidationRule } from '../types'

const numericLiteral = (value: unknown): value is number | string =>
  (typeof value === 'number' || typeof value === 'string' && value.trim() !== '') && Number.isFinite(Number(value))

export function importValidation(type: string, operator?: string, first?: string, second?: string): Omit<ValidationRule, 'id' | 'range'> | undefined {
  if (type === 'list') {
    if (first?.startsWith('"') && first.endsWith('"')) return { kind: 'list', options: first.slice(1, -1).split(',') }
    if (first?.trim()) return { kind: 'list', sourceRange: first.trim().replace(/^=/, '') }
    return
  }
  if (type === 'date') {
    if ((!operator || operator === 'between') && (!first || [1, 2].includes(Number(first))) && (!second || Number(second) === 2958465)) return { kind: 'date' }
    return
  }
  if (type === 'textLength') {
    if ((!operator || operator === 'between') && numericLiteral(first) && numericLiteral(second))
      return { kind: 'text-length', min: Number(first), max: Number(second) }
    if ((operator === 'greaterThanOrEqual' || operator === 'greaterThan') && numericLiteral(first)) return { kind: 'text-length', min: Number(first) }
    if ((operator === 'lessThanOrEqual' || operator === 'lessThan') && numericLiteral(first)) return { kind: 'text-length', max: Number(first) }
    return
  }
  if (type === 'custom' && first?.trim()) {
    return { kind: 'custom-formula', formula: first.trim() }
  }
  if (type !== 'decimal' && type !== 'whole') return
  const whole = type === 'whole' || undefined
  if ((!operator || operator === 'between') && numericLiteral(first) && numericLiteral(second))
    return { kind: 'number-range', min: Number(first), max: Number(second), whole }
  if ((operator === 'greaterThanOrEqual' || operator === 'greaterThan') && numericLiteral(first)) return { kind: 'number-range', min: Number(first), whole }
  if ((operator === 'lessThanOrEqual' || operator === 'lessThan') && numericLiteral(first)) return { kind: 'number-range', max: Number(first), whole }
}

export function exportValidation(rule: ValidationRule): ExcelJS.DataValidation | undefined {
  const common = { allowBlank: true, showErrorMessage: true, errorStyle: 'warning' }
  if (rule.kind === 'list') {
    if (rule.sourceRange?.trim()) {
      return { ...common, type: 'list', formulae: [rule.sourceRange.trim()] }
    }
    const list = rule.options?.join(',') ?? ''
    if (!list || list.length > 255 || rule.options?.some(option => /[",\r\n]/.test(option))) return
    return { ...common, type: 'list', formulae: [`"${list}"`] }
  }
  if (rule.kind === 'date') return {
    ...common, type: 'date', operator: 'between', formulae: [new Date(Date.UTC(1900, 0, 1)), new Date(Date.UTC(9999, 11, 31))],
  }
  if (rule.kind === 'text-length') return {
    ...common, type: 'textLength', operator: 'between', formulae: [rule.min ?? 0, rule.max ?? 10_000],
  }
  if (rule.kind === 'custom-formula') {
    if (!rule.formula?.trim()) return
    return { ...common, type: 'custom', formulae: [rule.formula.trim().replace(/^=/, '')] }
  }
  if (rule.kind !== 'number-range') return
  const type = rule.whole ? 'whole' : 'decimal'
  if (rule.min !== undefined && rule.max !== undefined) return { ...common, type, operator: 'between', formulae: [rule.min, rule.max] }
  if (rule.min !== undefined) return { ...common, type, operator: 'greaterThanOrEqual', formulae: [rule.min] }
  if (rule.max !== undefined) return { ...common, type, operator: 'lessThanOrEqual', formulae: [rule.max] }
}

export function importConditional(rule: ExcelJS.ConditionalFormattingRule, color: (value?: Partial<ExcelJS.Color>) => string | undefined): Omit<ConditionalFormatRule, 'id' | 'range'> | undefined {
  const font = rule.style?.font
  const fill = rule.style?.fill
  const style: CellStyle = {
    ...(font?.bold ? { bold: true } : {}),
    ...(font?.italic ? { italic: true } : {}),
    ...(font?.underline ? { underline: true } : {}),
    textColor: color(font?.color),
    fillColor: fill?.type === 'pattern' ? color(fill.fgColor ?? fill.bgColor) : undefined,
  }
  if (rule.type === 'cellIs') {
    const op = rule.operator as string | undefined
    const kind = op === 'greaterThan' ? 'greater' : op === 'lessThan' ? 'less' : op === 'notEqual' ? 'not-equal' : op
    const [first, second] = rule.formulae ?? []
    if (!kind || !['greater', 'less', 'equal', 'not-equal', 'between'].includes(kind) || !numericLiteral(first) || (kind === 'between' && !numericLiteral(second))) return
    return { style, condition: { kind: kind as ConditionalFormatRule['condition']['kind'], value: String(first), value2: kind === 'between' ? String(second) : undefined } }
  }
  if ((rule as { type?: string }).type === 'duplicateValues') {
    return { style, condition: { kind: 'duplicate', value: '' } }
  }
  if (rule.type === 'containsText' || rule.type === 'expression') {
    const formula = (rule as { formulae?: unknown[] }).formulae?.[0]
    const value = rule.type === 'containsText' && rule.text !== undefined ? rule.text :
      typeof formula === 'string' ? /^NOT\(ISERROR\(SEARCH\("((?:[^"]|"")*)",\$?[A-Z]+\$?\d+\)\)\)$/i.exec(formula)?.[1]?.replace(/""/g, '"') : undefined
    if (value !== undefined) return { style, condition: { kind: 'text-contains', value } }
    if (rule.type === 'expression' && typeof formula === 'string' && formula.trim()) {
      return { style, condition: { kind: 'custom-formula', value: formula.trim() } }
    }
  }
}

const argb = (hex: string) => 'FF' + hex.replace('#', '').toUpperCase()

export function exportConditional(rule: ConditionalFormatRule, range: string, priority: number): ExcelJS.ConditionalFormattingRule | undefined {
  const style: Partial<ExcelJS.Style> = {}
  const { fillColor, textColor, bold, italic, underline } = rule.style
  if (fillColor) style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(fillColor) }, bgColor: { argb: argb(fillColor) } }
  if (textColor || bold || italic || underline) style.font = { bold, italic, underline, ...(textColor ? { color: { argb: argb(textColor) } } : {}) } as ExcelJS.Font
  const { kind, value, value2 } = rule.condition
  if (kind === 'text-contains') return {
    type: 'expression', priority, style,
    formulae: [`NOT(ISERROR(SEARCH("${value.replace(/"/g, '""')}",${range.split(':')[0]})))`],
  }
  if (kind === 'custom-formula') {
    if (!value.trim()) return
    return { type: 'expression', priority, style, formulae: [value.trim().replace(/^=/, '')] }
  }
  if (!['greater', 'less', 'equal', 'not-equal', 'between'].includes(kind)) return
  if (!numericLiteral(value) || (kind === 'between' && !numericLiteral(value2))) return
  const operator = kind === 'greater' ? 'greaterThan' : kind === 'less' ? 'lessThan' : kind === 'not-equal' ? 'notEqual' : kind
  return { type: 'cellIs', priority, style, operator: operator as 'greaterThan' | 'lessThan' | 'equal' | 'between', formulae: kind === 'between' ? [Number(value), Number(value2)] : [Number(value)] }
}
