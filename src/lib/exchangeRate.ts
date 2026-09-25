/** Dated source data only. No classification, model invocation or cell writes. */
export async function getExchangeRate(from: string, to: string, fetcher: typeof fetch = fetch) {
  from = from.trim().toUpperCase()
  to = to.trim().toUpperCase()
  if (!/^[A-Z]{3}$/.test(from) || !/^[A-Z]{3}$/.test(to) || from === to) {
    throw new Error('Provide two distinct three-letter currency codes, e.g. USD and EUR.')
  }
  const source = `https://api.frankfurter.app/latest?from=${from}&to=${to}`
  const response = await fetcher(source, { signal: AbortSignal.timeout(10_000) })
  if (!response.ok) throw new Error(`Exchange-rate source failed (HTTP ${response.status}). No rate is available; do not guess.`)
  const data = await response.json() as { base?: string; amount?: number; date?: string; rates?: Record<string, number> }
  const rate = data.rates?.[to]
  if (data.base !== from || data.amount !== 1 || typeof rate !== 'number' || !Number.isFinite(rate) || rate <= 0 || typeof data.date !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(data.date) || !Number.isFinite(Date.parse(data.date)) || data.date > new Date().toISOString().slice(0, 10)) {
    throw new Error('Exchange-rate source returned invalid or undated data. No rate is available; do not guess.')
  }
  return { from, to, rate, asOf: data.date, retrievedAt: new Date().toISOString(), source,
    note: 'Published reference rate for the stated date, not a real-time trading quote. Cite the date/source in the workbook. No cells changed.' }
}
