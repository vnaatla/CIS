import { signal } from '@angular/core'

/** What the shared grid asks for in server mode (`pageRequest`). */
export interface PageRequest {
  page: number
  pageSize: number
  sortKey: string | null
  sortDir: 'asc' | 'desc'
}

/** `{ rows, total }` of a paged API search. */
export interface Paged<T> {
  rows: T[]
  total: number
}

/**
 * State of one server-paged grid: bind `[serverTotal]="paging.total()" [serverPage]="paging.page()"`, call `change()` from
 * `(pageRequest)` and send `query(page)` with the search; `loaded()` after the answer, `reset()` on Clear.
 */
export class ServerPaging {
  readonly total = signal(0)
  readonly page = signal(0)
  private size = 50
  private sortKey = ''
  private sortDesc = false

  /** Takes the grid's size / sort; returns the page to load. */
  change(r: PageRequest): number {
    this.size = r.pageSize
    this.sortKey = r.sortKey ?? ''
    this.sortDesc = r.sortDir === 'desc'
    return r.page
  }

  /** Query-string values for page `page`. */
  query(page: number): Record<string, string | number | boolean> {
    return { page, pageSize: this.size, sortKey: this.sortKey, sortDesc: this.sortDesc }
  }

  loaded(page: number, total: number): void {
    this.page.set(page)
    this.total.set(total)
  }

  reset(): void {
    this.page.set(0)
    this.total.set(0)
  }
}
