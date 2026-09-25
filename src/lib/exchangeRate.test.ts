import { describe, expect, it, vi } from 'vitest'
import { getExchangeRate } from './exchangeRate'

describe('dated exchange-rate lookup', () => {
  it('returns source date and rate without inference or writes', async () => {
    const fetcher = vi.fn().mockResolvedValue({ ok: true, json: async () => ({ base: 'USD', amount: 1, date: '2026-09-09', rates: { EUR: 0.85 } }) })
    const result = await getExchangeRate('usd', 'eur', fetcher)
    expect(result).toMatchObject({ from: 'USD', to: 'EUR', rate: 0.85, asOf: '2026-09-09', source: 'https://api.frankfurter.app/latest?from=USD&to=EUR' })
    expect(fetcher).toHaveBeenCalledWith(result.source, expect.objectContaining({ signal: expect.any(AbortSignal) }))
  })
  it.each(['USD&token=secret', '', 'US'])('rejects invalid currency %s before fetching', async from => {
    const fetcher = vi.fn()
    await expect(getExchangeRate(from, 'EUR', fetcher)).rejects.toThrow('currency codes')
    expect(fetcher).not.toHaveBeenCalled()
  })
  it.each([
    { base: 'USD', amount: 1, rates: { EUR: 0.85 } },
    { base: 'USD', amount: 1, date: '2026-09-09', rates: { EUR: -1 } },
    { base: 'GBP', amount: 1, date: '2026-09-09', rates: { EUR: 0.85 } },
  ])('refuses unusable source data without guessing a replacement', async data => {
    await expect(getExchangeRate('USD', 'EUR', vi.fn().mockResolvedValue({ ok: true, json: async () => data }))).rejects.toThrow('invalid or undated')
  })
  it('reports provider failure without fallback', async () => {
    await expect(getExchangeRate('USD', 'EUR', vi.fn().mockResolvedValue({ ok: false, status: 503 }))).rejects.toThrow('HTTP 503')
  })
})
