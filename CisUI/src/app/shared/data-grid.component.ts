import { ChangeDetectionStrategy, Component, computed, input, output, signal } from '@angular/core'

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
          @for (row of sorted(); track rowKey()(row)) {
            <tr
              [class.clickable]="selectable()"
              [class.selected]="selectedKey() !== null && rowKey()(row) === selectedKey()"
              [attr.tabindex]="selectable() ? 0 : null"
              [attr.aria-selected]="selectable() ? rowKey()(row) === selectedKey() : null"
              (click)="selectable() && rowSelected.emit(row)"
              (keydown.enter)="selectable() && rowSelected.emit(row)"
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

  readonly rowSelected = output<T>()
  readonly checkedChange = output<Set<string>>()

  private readonly sort = signal<Sort>(null)

  protected readonly sorted = computed(() => {
    const s = this.sort()
    const col = s && this.columns().find((c) => c.key === s.key)
    if (!s || !col) return this.rows()
    const by = (r: T) => (col.sortBy ? col.sortBy(r) : (this.raw(col, r) ?? ''))
    return [...this.rows()].sort((a, b) => (s.dir === 'asc' ? 1 : -1) * compare(by(a), by(b)))
  })

  private readonly checkedCount = computed(() => this.rows().filter((r) => this.checked().has(this.rowKey()(r))).length)
  protected readonly allChecked = computed(() => this.rows().length > 0 && this.checkedCount() === this.rows().length)
  protected readonly someChecked = computed(() => this.checkedCount() > 0 && this.checkedCount() < this.rows().length)

  protected toggleSort(key: string): void {
    this.sort.update((s) => (s?.key !== key ? { key, dir: 'asc' } : s.dir === 'asc' ? { key, dir: 'desc' } : null))
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
