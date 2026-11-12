import { ChangeDetectionStrategy, Component, computed, input, linkedSignal, output, signal } from '@angular/core'

export interface GridColumn<T> {
  key: string
  label: string
  /** Cell text; defaults to `String(row[key])`. */
  value?: (row: T) => string | number | null | undefined
  /** Value used for sorting when it differs from the display text (dates, money). */
  sortBy?: (row: T) => string | number
  align?: 'left' | 'right' | 'center'
  sortable?: boolean
}

type Sort = { key: string; dir: 'asc' | 'desc' } | null

const compare = (a: string | number, b: string | number) =>
  typeof a === 'number' && typeof b === 'number' ? a - b : String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' })

/** Sortable, optionally selectable / checkable results table (GENESIS DataGrid; replaces the VB6 ListView + ColumnClick sort). */
@Component({
  selector: 'app-data-grid',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="table-wrap" [class.tall]="tall()">
      <table class="grid" [attr.aria-label]="caption()">
        <thead>
          <tr>
            @if (checkable()) {
              <th scope="col" class="check-cell">
                <input
                  type="checkbox"
                  aria-label="Select all rows"
                  [checked]="allChecked()"
                  [indeterminate]="someChecked()"
                  [disabled]="rows().length === 0"
                  (change)="toggleAll()"
                />
              </th>
            }
            @for (c of columns(); track c.key) {
              <th scope="col" [class]="cls(c)" [attr.aria-sort]="ariaSort(c)">
                @if (c.sortable === false) {
                  {{ c.label }}
                } @else {
                  <button type="button" class="th-sort" (click)="toggleSort(c.key)" [title]="'Sort by ' + c.label">
                    {{ c.label }}<span aria-hidden="true" class="sort-mark">{{ mark(c) }}</span>
                  </button>
                }
              </th>
            }
          </tr>
        </thead>
        <tbody>
          @for (row of paged(); track rowKey()(row)) {
            <tr
              [class.clickable]="selectable()"
              [class.selected]="selectedKey() !== null && rowKey()(row) === selectedKey()"
              [attr.tabindex]="selectable() ? 0 : null"
              [attr.aria-selected]="selectable() ? rowKey()(row) === selectedKey() : null"
              (click)="selectable() && rowSelected.emit(row)"
              (keydown.enter)="selectable() && rowSelected.emit(row)"
              (dblclick)="rowActivated.emit(row)"
            >
              @if (checkable()) {
                <td class="check-cell">
                  <input type="checkbox" [attr.aria-label]="'Select row ' + rowKey()(row)" [checked]="checked().has(rowKey()(row))" (click)="$event.stopPropagation()" (change)="toggleRow(rowKey()(row))" />
                </td>
              }
              @for (c of columns(); track c.key) {
                <td [class]="cls(c)">{{ text(c, row) }}</td>
              }
            </tr>
          } @empty {
            <tr class="empty-row">
              <td [attr.colspan]="columns().length + (checkable() ? 1 : 0)">{{ empty() }}</td>
            </tr>
          }
        </tbody>
      </table>
    </div>
    @if (pageSize() > 0 && total() > size()) {
      <nav class="grid-pager" [attr.aria-label]="caption() + ' pages'">
        <span>Rows {{ first() + 1 }}–{{ last() }} of {{ total() }}</span>
        <ul class="pagination">
          <li [class.disabled]="page() === 0"><button type="button" class="page-btn" [disabled]="page() === 0" (click)="goto(0)" aria-label="First page">First</button></li>
          <li [class.disabled]="page() === 0"><button type="button" class="page-btn" [disabled]="page() === 0" (click)="goto(page() - 1)">Previous</button></li>
          @for (n of pageList(); track $index) {
            @if (n < 0) {
              <li class="disabled"><span class="page-btn" aria-hidden="true">…</span></li>
            } @else {
              <li [class.active]="n === page()">
                <button type="button" class="page-btn" [attr.aria-current]="n === page() ? 'page' : null" (click)="goto(n)">{{ n + 1 }}</button>
              </li>
            }
          }
          <li [class.disabled]="page() >= pages() - 1"><button type="button" class="page-btn" [disabled]="page() >= pages() - 1" (click)="goto(page() + 1)">Next</button></li>
          <li [class.disabled]="page() >= pages() - 1"><button type="button" class="page-btn" [disabled]="page() >= pages() - 1" (click)="goto(pages() - 1)" aria-label="Last page">Last</button></li>
        </ul>
        <label>Rows per page
          <select [value]="size()" (change)="setSize($any($event.target).value)">
            @for (n of sizes; track n) {
              <option [value]="n" [selected]="n === size()">{{ n }}</option>
            }
          </select>
        </label>
      </nav>
    }
  `,
})
export class DataGridComponent<T extends object> {
  /** Accessible table name. */
  readonly caption = input.required<string>()
  readonly columns = input.required<readonly GridColumn<T>[]>()
  readonly rows = input.required<readonly T[]>()
  readonly rowKey = input.required<(row: T) => string>()
  readonly empty = input('No records found.')
  readonly selectable = input(false)
  readonly selectedKey = input<string | null>(null)
  /** Adds a leading checkbox column with a select-all header (VB6 checkbox ListView). */
  readonly checkable = input(false)
  readonly checked = input<ReadonlySet<string>>(new Set())
  readonly tall = input(false)
  /** Rows per page (0 = no paging). Sorting, TOTAL RECORDS and exports still use every row. */
  readonly pageSize = input(50)
  /** Server paging: rows() is one page from the API, serverTotal() all matching rows; page / sort changes go out through pageRequest. */
  readonly serverTotal = input<number | null>(null)
  /** Server paging: the page the screen loaded (0-based); a new search sets it back to 0. */
  readonly serverPage = input<number | null>(null)
  readonly pageRequest = output<{ page: number; pageSize: number; sortKey: string | null; sortDir: 'asc' | 'desc' }>()

  readonly rowSelected = output<T>()
  /** A row double-clicked (screens that open the row elsewhere, e.g. Invoice Tracking). */
  readonly rowActivated = output<T>()
  readonly checkedChange = output<Set<string>>()

  private readonly sort = signal<Sort>(null)

  protected readonly sorted = computed(() => {
    const s = this.sort()
    const col = s && this.columns().find((c) => c.key === s.key)
    if (!s || !col || this.serverTotal() !== null) return this.rows()
    const by = (r: T) => (col.sortBy ? col.sortBy(r) : (this.raw(col, r) ?? ''))
    return [...this.rows()].sort((a, b) => (s.dir === 'asc' ? 1 : -1) * compare(by(a), by(b)))
  })

  protected readonly sizes = [25, 50, 100, 200]
  private readonly chosenSize = signal<number | null>(null)
  protected readonly size = computed(() => this.chosenSize() ?? this.pageSize())
  /** Back to the first page whenever the rows or the sort change. */
  protected readonly page = linkedSignal<unknown, number>({
    source: () => (this.serverTotal() !== null ? [this.sort(), this.size(), this.serverPage()] : [this.rows(), this.sort(), this.size()]),
    computation: () => (this.serverTotal() !== null ? (this.serverPage() ?? 0) : 0),
  })
  /** Matching rows: the API's total in server mode, else the rows given. */
  protected readonly total = computed(() => this.serverTotal() ?? this.rows().length)
  protected readonly pages = computed(() => Math.max(1, Math.ceil(this.total() / Math.max(1, this.size()))))
  protected readonly first = computed(() => this.page() * this.size())
  protected readonly last = computed(() => Math.min(this.total(), this.first() + this.size()))
  protected readonly paged = computed(() =>
    this.serverTotal() !== null || this.pageSize() <= 0 ? this.sorted() : this.sorted().slice(this.first(), this.last()),
  )

  /** Shows page n; in server mode asks the screen to load it. */
  protected goto(n: number): void {
    this.page.set(n)
    this.request()
  }

  private request(): void {
    if (this.serverTotal() === null) return
    const s = this.sort()
    this.pageRequest.emit({ page: this.page(), pageSize: this.size(), sortKey: s?.key ?? null, sortDir: s?.dir ?? 'asc' })
  }

  /** Page numbers to show: first, last and two either side of the current one; -1 = a gap (…). */
  protected readonly pageList = computed(() => {
    const total = this.pages()
    const cur = this.page()
    const keep = [...new Set([0, total - 1, cur - 2, cur - 1, cur, cur + 1, cur + 2])].filter((n) => n >= 0 && n < total).sort((a, b) => a - b)
    return keep.flatMap((n, i) => (i > 0 && n - keep[i - 1] > 1 ? [-1, n] : [n]))
  })

  protected setSize(value: string): void {
    this.chosenSize.set(Number(value))
    this.page.set(0)
    this.request()
  }

  private readonly checkedCount = computed(() => this.rows().filter((r) => this.checked().has(this.rowKey()(r))).length)
  protected readonly allChecked = computed(() => this.rows().length > 0 && this.checkedCount() === this.rows().length)
  protected readonly someChecked = computed(() => this.checkedCount() > 0 && this.checkedCount() < this.rows().length)

  protected toggleSort(key: string): void {
    this.sort.update((s) => (s?.key !== key ? { key, dir: 'asc' } : s.dir === 'asc' ? { key, dir: 'desc' } : null))
    this.page.set(0)
    this.request()
  }

  protected toggleRow(key: string): void {
    const next = new Set(this.checked())
    if (!next.delete(key)) next.add(key)
    this.checkedChange.emit(next)
  }

  protected toggleAll(): void {
    this.checkedChange.emit(this.allChecked() ? new Set() : new Set(this.rows().map((r) => this.rowKey()(r))))
  }

  protected cls(c: GridColumn<T>): string {
    return c.align === 'right' ? 'num' : c.align === 'center' ? 'center' : ''
  }

  protected ariaSort(c: GridColumn<T>): string | null {
    const s = this.sort()
    return s?.key === c.key ? (s.dir === 'asc' ? 'ascending' : 'descending') : null
  }

  protected mark(c: GridColumn<T>): string {
    const s = this.sort()
    return s?.key === c.key ? (s.dir === 'asc' ? '▲' : '▼') : ''
  }

  protected text(c: GridColumn<T>, row: T): string {
    return String(this.raw(c, row) ?? '')
  }

  private raw(c: GridColumn<T>, row: T): string | number | null | undefined {
    return c.value ? c.value(row) : ((row as Record<string, unknown>)[c.key] as string | number | null | undefined)
  }
}
