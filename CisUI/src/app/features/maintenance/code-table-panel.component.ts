import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, input, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { digitsOnly } from '../../shared/lib/masks'
import { openBlob } from '../../shared/lib/download'
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

export interface CodeFieldView {
  column: string
  label: string
  maxLength: number
  required: boolean
  kind: 'Text' | 'Number' | 'MoneyText' | 'ChoiceFirstChar'
  searchable: boolean
  isKey: boolean
  hasChoices: boolean
}

export interface CodeTableView {
  key: string
  legacyForm: string
  tab: string
  listReportKey: string | null
  fields: CodeFieldView[]
  identityKey: boolean
  allowDelete: boolean
}

type Row = Record<string, unknown> & { _id: string }
/** search = entries usable for Search / Add (desktop after Clear) · selected = a list row shown read-only · add / edit = entering a record. */
type Mode = 'search' | 'selected' | 'add' | 'edit'

/**
 * One maintenance code-table tab (frmCheckingItems and the tabs of frmMaintenanceCal / frmMaintenanceOthers / frmMaintenanceBNFDental /
 * frmMaintenanceBenefits): entry fields, Search / Add / Edit / Delete / Save / Clear / Print, the list and TOTAL RECORDS. Field captions,
 * lengths, mandatory marks and messages come from the API (Cis.Application CodeTableCatalog).
 */
@Component({
  selector: 'app-code-table-panel',
  imports: [FormsModule, FieldComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    @if (view(); as v) {
      <div class="screen-form">
        <fieldset [attr.aria-label]="v.tab">
          <legend>{{ v.tab }}</legend>
          @for (f of v.fields; track f.column) {
            <app-field [label]="f.label" [required]="f.required" [hint]="f.searchable ? 'Searchable' : undefined">
              @if (f.hasChoices) {
                <select [name]="f.column" [disabled]="!enabled(f)" [ngModel]="values()[f.column] ?? ''" (ngModelChange)="set(f, $event)">
                  <option value=""></option>
                  @for (o of choices()[f.column] ?? []; track o.code) {
                    <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
                  }
                </select>
              } @else {
                <input
                  [name]="f.column"
                  [attr.maxlength]="f.maxLength"
                  [attr.inputmode]="f.kind === 'Text' ? null : 'decimal'"
                  [disabled]="!enabled(f)"
                  [ngModel]="values()[f.column] ?? ''"
                  (ngModelChange)="set(f, $event)"
                />
              }
            </app-field>
          }
          <p class="muted wide">Searchable fields are marked; * = mandatory.</p>
        </fieldset>

        <app-action-bar [label]="v.tab + ' actions'">
          <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
          <button appButton (click)="add()">Add</button>
          <button appButton [disabled]="mode() !== 'selected'" (click)="edit()">Edit</button>
          @if (v.allowDelete) {
            <button appButton variant="danger" [disabled]="mode() !== 'selected'" (click)="remove()">Delete</button>
          }
          <button appButton [disabled]="mode() !== 'add' && mode() !== 'edit'" [busy]="busy()" (click)="save()">Save</button>
          <button appButton variant="ghost" (click)="clear()">Clear</button>
          @if (v.listReportKey) {
            <button appButton variant="ghost" (click)="print()">Print</button>
          }
          <span class="spacer"></span>
          <app-stat label="TOTAL RECORDS" [value]="rows().length" />
        </app-action-bar>

        <app-data-grid
          [caption]="v.tab + ' list'"
          [columns]="columns()"
          [rows]="rows()"
          [rowKey]="rowKey"
          [selectable]="true"
          [selectedKey]="selected()?._id ?? null"
          (rowSelected)="select($event)"
          [tall]="true"
        />
      </div>
    } @else if (loadError()) {
      <p class="error" role="alert">{{ loadError() }}</p>
    }
  `,
})
export class CodeTablePanelComponent implements OnInit {
  /** Catalog key, e.g. 'check-items'. */
  readonly tableKey = input.required<string>()

  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly view = signal<CodeTableView | null>(null)
  protected readonly loadError = signal('')
  protected readonly choices = signal<Record<string, LookupOption[]>>({})
  protected readonly rows = signal<Row[]>([])
  protected readonly selected = signal<Row | null>(null)
  protected readonly values = signal<Record<string, string>>({})
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  private lastFilters: Record<string, string> = {}

  protected readonly rowKey = (r: Row) => r._id
  protected readonly columns = computed<GridColumn<Row>[]>(() =>
    (this.view()?.fields ?? []).map((f) => ({ key: f.column, label: f.label, align: f.kind === 'Text' || f.kind === 'ChoiceFirstChar' ? 'left' : 'right' })),
  )

  private get base(): string {
    return `/api/maintenance/code-tables/${this.tableKey()}`
  }

  async ngOnInit(): Promise<void> {
    try {
      const view = await firstValueFrom(this.http.get<CodeTableView>(this.base))
      const lists: Record<string, LookupOption[]> = {}
      for (const f of view.fields.filter((x) => x.hasChoices)) lists[f.column] = await firstValueFrom(this.http.get<LookupOption[]>(`${this.base}/choices/${f.column}`))
      this.choices.set(lists)
      this.view.set(view)
    } catch (e) {
      this.loadError.set(errorMessage(e))
    }
  }

  /** Key read-only while editing; everything read-only when a list row is shown. */
  protected enabled(f: CodeFieldView): boolean {
    const m = this.mode()
    if (f.isKey && this.view()?.identityKey) return false
    return m === 'search' || m === 'add' || (m === 'edit' && !f.isKey)
  }

  protected set(f: CodeFieldView, value: string): void {
    this.values.update((v) => ({ ...v, [f.column]: f.kind === 'Number' || f.kind === 'MoneyText' ? digitsOnly(value) : value }))
  }

  /** cmdSearch_Click: the searchable fields that have a value; none = whole table. "Record not found!" when empty. */
  protected async search(): Promise<void> {
    const filters: Record<string, string> = {}
    for (const f of this.view()?.fields ?? []) if (f.searchable && this.values()[f.column]?.trim()) filters[f.column] = this.values()[f.column].trim()
    if (await this.load(filters)) {
      this.values.set({})
      this.mode.set('search')
    }
  }

  private async load(filters: Record<string, string>): Promise<boolean> {
    this.busy.set(true)
    this.rows.set([])
    this.selected.set(null)
    try {
      const rows =
        (await firstValueFrom(this.http.get<Record<string, unknown>[] | null>(`${this.base}/rows`, { params: new HttpParams({ fromObject: filters }) }))) ?? []
      this.lastFilters = filters
      this.rows.set(rows.map((r, i) => ({ ...r, _id: String(i) })))
      if (rows.length === 0) this.notify.notify('Record not found!', 'info')
      return true
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
      return false
    } finally {
      this.busy.set(false)
    }
  }

  protected select(row: Row): void {
    this.selected.set(row)
    const v: Record<string, string> = {}
    for (const f of this.view()?.fields ?? []) v[f.column] = row[f.column] == null ? '' : String(row[f.column])
    this.values.set(v)
    this.mode.set('selected')
  }

  protected add(): void {
    this.selected.set(null)
    this.values.set({})
    this.mode.set('add')
  }

  protected edit(): void {
    if (this.selected()) this.mode.set('edit')
  }

  protected clear(): void {
    this.values.set({})
    this.rows.set([])
    this.selected.set(null)
    this.mode.set('search')
  }

  /** cmdSave_Click: the API applies the desktop's checks and answers "Record Saved..." / "Record Updated...". */
  protected async save(): Promise<void> {
    const view = this.view()
    if (!view || !(await this.notify.ask('Do you want to save the record?'))) return
    const keyField = view.fields.find((f) => f.isKey)!
    this.busy.set(true)
    try {
      const body = { values: this.values() }
      const answer =
        this.mode() === 'edit'
          ? await firstValueFrom(this.http.put<{ message: string }>(`${this.base}/rows/${encodeURIComponent(String(this.selected()![keyField.column]))}`, body))
          : await firstValueFrom(this.http.post<{ message: string }>(`${this.base}/rows`, body))
      this.notify.notify(answer.message, 'success')
      this.values.set({})
      this.mode.set('search')
      await this.load(this.lastFilters)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdDelete_Click: needs a selected list row; deletes every row with its code, as the desktop did. */
  protected async remove(): Promise<void> {
    const view = this.view()
    const row = this.selected()
    if (!view || !row) {
      this.notify.notify('Please search the record', 'warn')
      return
    }
    if (!(await this.notify.ask('Do you want to delete this record?'))) return
    const keyField = view.fields.find((f) => f.isKey)!
    try {
      const answer = await firstValueFrom(this.http.delete<{ message: string }>(`${this.base}/rows/${encodeURIComponent(String(row[keyField.column]))}`))
      this.notify.notify(answer.message, 'success')
      this.values.set({})
      this.mode.set('search')
      await this.load(this.lastFilters)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdPrint_Click: the tab's list print-out (the desktop's DataReport) as PDF. */
  protected async print(): Promise<void> {
    const key = this.view()?.listReportKey
    if (!key) return
    try {
      openBlob(await firstValueFrom(this.http.get(`/api/maintenance/list-reports/${key}`, { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }
}
