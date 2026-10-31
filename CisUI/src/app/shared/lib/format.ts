const ISO_DATE = /^\d{4}-\d{2}-\d{2}T/

/** ISO date-time from the API → dd/mm/yyyy (the desktop's display format). */
export function formatDate(value: unknown): string {
  if (typeof value !== 'string' || !ISO_DATE.test(value)) return value == null ? '' : String(value)
  const [y, m, d] = value.slice(0, 10).split('-')
  return `${d}/${m}/${y}`
}

/** RM amounts: thousands separators, two decimals. */
export function formatMoney(value: unknown): string {
  if (value == null || value === '') return ''
  const n = Number(value)
  return Number.isNaN(n) ? String(value) : n.toLocaleString('en-MY', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

/** Generic cell text: Yes / No for booleans, dd/mm/yyyy for ISO dates. */
export function formatValue(value: unknown): string {
  if (value == null) return ''
  if (typeof value === 'boolean') return value ? 'Yes' : 'No'
  if (typeof value === 'string' && ISO_DATE.test(value)) return formatDate(value)
  return String(value)
}
