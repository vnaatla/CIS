const FILE_NAME_PATTERN = /filename\*?=(?:UTF-8'')?"?([^";]+)"?/i

/** File name from a Content-Disposition header, or the fallback. */
export function fileNameFromDisposition(header: string | null | undefined, fallback: string): string {
  const match = FILE_NAME_PATTERN.exec(header ?? '')
  return match ? decodeURIComponent(match[1]) : fallback
}

/** Saves a blob through a temporary link (the browser's download prompt). */
export function saveBlob(blob: Blob, fileName: string): void {
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = fileName
  link.click()
  URL.revokeObjectURL(link.href)
}

/** Opens a blob (PDF) in a new tab and releases it once the tab has had time to load it. */
export function openBlob(blob: Blob): void {
  const url = URL.createObjectURL(blob)
  window.open(url, '_blank', 'noopener')
  setTimeout(() => URL.revokeObjectURL(url), 60_000)
}

/** A CSV cell: quoted when needed; leading = + - @ are prefixed so spreadsheets do not run them as formulas. */
export function csvCell(value: string | number | null | undefined): string {
  let text = value == null ? '' : String(value)
  if (/^[=+\-@]/.test(text) && Number.isNaN(Number(text))) text = `'${text}`
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

/** CSV text (header row, then one line per row) with a BOM so Excel opens it as UTF-8. */
export function toCsv(rows: (string | number | null | undefined)[][]): string {
  return '\ufeff' + rows.map((r) => r.map(csvCell).join(',')).join('\r\n')
}

/** Client-side export of a table (replaces ModListviewtoExcel.ExportListViewtoExcel for on-screen grids). */
export function saveCsv(rows: (string | number | null | undefined)[][], fileName: string): void {
  saveBlob(new Blob([toCsv(rows)], { type: 'text/csv;charset=utf-8' }), fileName)
}
