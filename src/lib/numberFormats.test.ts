import { describe, expect, it } from 'vitest'
import { familyOfPattern, patternOfStyle, styleFromPattern } from './numberFormats'

describe('number format codes', () => {
  it('names the family without being misled by colours, quotes or locales', () => {
    expect(familyOfPattern('General')).toBeNull()
    expect(familyOfPattern('#,##0;[Red]-#,##0')).toBe('number')
    expect(familyOfPattern('0.0" days"')).toBe('number')
    expect(familyOfPattern('#,##0.00_);(#,##0.00)')).toBe('number')
    expect(familyOfPattern('[$€-407] #,##0.00')).toBe('currency')
    expect(familyOfPattern('_("$"* #,##0.00_)')).toBe('currency')
    expect(familyOfPattern('0%')).toBe('percent')
    expect(familyOfPattern('d mmm yyyy')).toBe('date')
    expect(familyOfPattern('[h]:mm:ss')).toBe('date')
    expect(familyOfPattern('@')).toBe('text')
  })

  it('keeps the code only when the family and decimals would not write it back', () => {
    expect(styleFromPattern('$#,##0.00')).toEqual({ numberFormat: 'currency' })
    expect(styleFromPattern('0.0%')).toEqual({ numberFormat: 'percent', decimals: 1 })
    expect(styleFromPattern('#,##0.0;[Red]-#,##0.0')).toEqual({
      numberFormat: 'number',
      decimals: 1,
      numberPattern: '#,##0.0;[Red]-#,##0.0',
    })
    expect(styleFromPattern('General')).toBeNull()
  })

  it('shows the kept code before the family', () => {
    expect(patternOfStyle({ numberFormat: 'date', numberPattern: 'd mmm yyyy' })).toBe('d mmm yyyy')
    expect(patternOfStyle({ numberFormat: 'currency', decimals: 0 })).toBe('$#,##0')
    expect(patternOfStyle({ bold: true })).toBeUndefined()
  })
})
