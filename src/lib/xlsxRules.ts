import type ExcelJS from 'exceljs'
import type { CellStyle, ConditionalFormatRule, ValidationRule } from '../types'

const numericLiteral = (value: unknown): value is number | string =>
  (typeof value === 'number' || typeof value === 'string' && value.trim() !== '') && Number.isFinite(Number(value))

export function importValidation(type: string, operator?: string, first?: string, second?: string): Omit<ValidationRule, 'id' | 'range'> | undefined {
  if (type === 'list' && first?.startsWith('"') && first.endsWith('"')) return { kind: 'list', options: first.slice(1, -1).split(',') }
  if (type === 'date' && (!operator || operator === 'between') && [1, 2].includes(Number(first)) && Number(second) === 2958465) return { kind: 'date' }
  if (type !== 'decimal' && type !== 'whole') return
  const whole = type === 'whole' || undefined
  if ((!operator || operator === 'between') && numericLiteral(first) && numericLiteral(second))
    return { kind: 'number-range', min: Number(first), max: Number(second), whole }
  if (operator === 'greaterThanOrEqual' && numericLiteral(first)) return { kind: 'number-range', min: Number(first), whole }
  if (operator === 'lessThanOrEqual' && numericLiteral(first)) return { kind: 'number-range', max: Number(first), whole }
}

export function exportValidation(rule: ValidationRule): ExcelJS.DataValidation | undefined {
  const common = { allowBlank: true, showErrorMessage: true, errorStyle: 'warning' }
  if (rule.kind === 'list') {
    const list = rule.options?.join(',') ?? ''
    if (!list || list.length > 255 || rule.options?.some(option => /[",\r\n]/.test(option))) return
    return { ...common, type: 'list', formulae: [`"${list}"`] }
  }
  if (rule.kind === 'date') return {
    ...common, type: 'date', operator: 'between', formulae: [new Date(Date.UTC(1900, 0, 1)), new Date(Date.UTC(9999, 11, 31))],
  }
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
    const kind = rule.operator === 'greaterThan' ? 'greater' : rule.operator === 'lessThan' ? 'less' : rule.operator
    const [first, second] = rule.formulae ?? []
    if (!kind || !['greater', 'less', 'equal', 'between'].includes(kind) || !numericLiteral(first) || (kind === 'between' && !numericLiteral(second))) return
    return { style, condition: { kind, value: String(first), value2: kind === 'between' ? String(second) : undefined } }
  }
  if (rule.type === 'containsText' || rule.type === 'expression') {
    const formula = (rule as { formulae?: unknown[] }).formulae?.[0]
    const value = rule.type === 'containsText' && rule.text !== undefined ? rule.text :
      typeof formula === 'string' ? /^NOT\(ISERROR\(SEARCH\("((?:[^"]|"")*)",\$?[A-Z]+\$?\d+\)\)\)$/i.exec(formula)?.[1]?.replace(/""/g, '"') : undefined
    if (value !== undefined) return { style, condition: { kind: 'text-contains', value } }
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
  if (!numericLiteral(value) || (kind === 'between' && !numericLiteral(value2))) return
  const operator = kind === 'greater' ? 'greaterThan' : kind === 'less' ? 'lessThan' : kind
  return { type: 'cellIs', priority, style, operator, formulae: kind === 'between' ? [Number(value), Number(value2)] : [Number(value)] }
}
