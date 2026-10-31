import { csvCell, fileNameFromDisposition, toCsv } from './download'
import { formatDate, formatMoney, formatValue } from './format'
import { digitsOnly, dmyToIso, isoToDmy, maskDate, maskTime, parseDmy } from './masks'

describe('masks (MaskEdBox / KeyNumber parity)', () => {
  it('masks dates as dd/mm/yyyy while typing', () => {
    expect(maskDate('3')).toBe('3')
    expect(maskDate('3110')).toBe('31/10')
    expect(maskDate('31102026')).toBe('31/10/2026')
    expect(maskDate('31/10/2026 extra 99')).toBe('31/10/2026')
  })

  it('accepts only real calendar dates', () => {
    expect(parseDmy('29/02/2024')).not.toBeNull()
    expect(parseDmy('29/02/2026')).toBeNull()
    expect(parseDmy('31/04/2026')).toBeNull()
    expect(parseDmy('1/1/2026')).toBeNull()
  })

  it('converts between dd/mm/yyyy and ISO', () => {
    expect(dmyToIso('31/10/2026')).toBe('2026-10-31')
    expect(dmyToIso('31/02/2026')).toBe('')
    expect(isoToDmy('2026-10-31T00:00:00')).toBe('31/10/2026')
    expect(isoToDmy(null)).toBe('')
  })

  it('masks times and keeps digits with one decimal point', () => {
    expect(maskTime('0930')).toBe('09:30')
    expect(digitsOnly('1a2.3.4')).toBe('12.34')
    expect(digitsOnly('RM 1,000')).toBe('1000')
  })
})

describe('format', () => {
  it('shows ISO dates as dd/mm/yyyy and money with two decimals', () => {
    expect(formatDate('2026-10-31T21:10:35')).toBe('31/10/2026')
    expect(formatMoney(1234.5)).toBe('1,234.50')
    expect(formatMoney('')).toBe('')
    expect(formatValue(true)).toBe('Yes')
    expect(formatValue(null)).toBe('')
  })
})

describe('download / csv', () => {
  it('quotes cells and neutralises spreadsheet formulas', () => {
    expect(csvCell('a,b')).toBe('"a,b"')
    expect(csvCell('say "hi"')).toBe('"say ""hi"""')
    expect(csvCell('=SUM(A1)')).toBe("'=SUM(A1)")
    expect(csvCell('-12.5')).toBe('-12.5')
    expect(toCsv([['A', 'B'], [1, null]])).toBe('\ufeffA,B\r\n1,')
  })

  it('reads the file name from Content-Disposition', () => {
    expect(fileNameFromDisposition('attachment; filename="Claim Report.xlsx"', 'x')).toBe('Claim Report.xlsx')
    expect(fileNameFromDisposition("attachment; filename*=UTF-8''Bordx%20123.pdf", 'x')).toBe('Bordx 123.pdf')
    expect(fileNameFromDisposition(null, 'fallback.csv')).toBe('fallback.csv')
  })
})
