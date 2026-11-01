import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { openBlob } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { parseDmy } from '../../shared/lib/masks'
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

const BASE = '/api/maintenance/plan/visit-limits'

type Key = 'plnCode' | 'hltCode' | 'insCode' | 'payCode' | 'grpCompany' | 'visitLimit' | 'suppStatus' | 'suppLimit' | 'effDate' | 'version'
type Limit = Record<Key, string>
const EMPTY: Limit = { plnCode: '', hltCode: '', insCode: '', payCode: '', grpCompany: '', visitLimit: '', suppStatus: 'A', suppLimit: '', effDate: '', version: '' }
type Row = Record<string, unknown> & { _id: string }
type Mode = 'search' | 'selected' | 'add' | 'edit'
interface Setup {
  plans: LookupOption[]
  insuredTypes: LookupOption[]
  healthTypes: LookupOption[]
  payors: LookupOption[]
  groupCompanies: LookupOption[]
}
/** lstVisitLimit columns. */
const LIST: [string, string, boolean?][] = [
  ['HLTCode', 'Health Type'], ['VisitIndex', 'Index'], ['PLNCode', 'Plan Code'], ['INSCode', 'Insured Type'], ['VisitLimit', 'Visit Limit', true],
  ['SuppVisitLimit', 'Supp. Visit Limit', true], ['VisitEffDate', 'Eff. Date'], ['PAYCode', 'Payor'], ['GRPCompany', 'Group Company'],
  ['VisitLimitVersion', 'Version'], ['SuppVisitLimitStatus', 'Supp. Status'],
]

/** frmMaintenancePlan tab 5 "Visit Limit" (CISDB VisitLimit / VIEW_VISIT_PAY_PLN). */
@Component({
  selector: 'app-plan-visit-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Visit Limit">
        <legend>Visit Limit</legend>
        <app-field label="Plan Code" [required]="true" hint="Searchable">
          <input name="plnCode" list="visit-plans" [disabled]="locked()" [(ngModel)]="v.plnCode" />
          <datalist id="visit-plans">
            @for (o of setup().plans; track o.code) {
              <option [value]="o.code + ' - ' + o.name"></option>
            }
          </datalist>
        </app-field>
        @for (c of combos; track c.key) {
          <app-field [label]="c.label" [required]="!!c.required" hint="Searchable">
            <select [name]="c.key" [attr.name]="c.key" [disabled]="locked() || (!!c.searchOnly && entering())" [(ngModel)]="v[c.key]">
              <option value=""></option>
              @for (o of setup()[c.list]; track o.code) {
                <option [value]="o.code">{{ o.code === o.name ? o.code : o.code + ' - ' + o.name }}</option>
              }
            </select>
          </app-field>
        }
        <app-field label="Visit Limit" [required]="true">
          <input class="num" name="visitLimit" inputmode="decimal" [disabled]="locked()" [(ngModel)]="v.visitLimit" />
        </app-field>
        <app-field label="Supp. Visit Lmt Status">
          <select name="suppStatus" [disabled]="locked()" [(ngModel)]="v.suppStatus">
            <option value="A">A - Share Limit For Family</option>
            <option value="B">B - Share Limit For Supplementary</option>
            <option value="C">C - Separate Limit For Supplementary</option>
          </select>
        </app-field>
        <app-field label="Supp. Visit Lmt" [required]="v.suppStatus !== 'A'">
          <input class="num" name="suppLimit" inputmode="decimal" [disabled]="locked() || v.suppStatus === 'A'" [(ngModel)]="v.suppLimit" />
        </app-field>
        <app-field label="Visit Limit Eff. Date" [required]="true">
          <app-date-input name="effDate" ariaLabel="Visit Limit Eff. Date" [disabled]="locked()" [(ngModel)]="v.effDate" (focusout)="checkDate()" />
        </app-field>
        <app-field label="Visit Limit Version" [required]="true">
          <input name="version" maxlength="9" inputmode="numeric" [disabled]="locked()" [(ngModel)]="v.version" />
        </app-field>
      </fieldset>

      <app-action-bar label="Visit Limit actions">
        <button appButton variant="primary" [busy]="busy()" [disabled]="entering()" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [disabled]="mode() !== 'selected'" (click)="mode.set('edit')">Edit</button>
        <button appButton variant="danger" [disabled]="mode() !== 'selected'" (click)="remove()">Delete</button>
        <button appButton [disabled]="!entering()" [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <button appButton variant="ghost" (click)="refresh()">Refresh</button>
        <button appButton variant="ghost" (click)="print()">Print</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="rows().length" />
      </app-action-bar>

      <app-data-grid caption="Visit limits" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
        [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class PlanVisitPanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly combos: { key: Key; label: string; list: keyof Setup; required?: boolean; searchOnly?: boolean }[] = [
    { key: 'hltCode', label: 'Health Type', list: 'healthTypes', required: true },
    { key: 'insCode', label: 'Insured Type', list: 'insuredTypes', required: true },
    { key: 'payCode', label: 'Payor Code', list: 'payors', searchOnly: true },
    { key: 'grpCompany', label: 'Group Company', list: 'groupCompanies', searchOnly: true },
  ]
  protected readonly columns: GridColumn<Row>[] = LIST.map(([key, label, isMoney]) => ({
    key, label, align: isMoney ? 'right' : undefined,
    value: (r: Row) => (isMoney ? formatMoney(r[key]) : key === 'VisitEffDate' ? formatDate(r[key]) : (r[key] as string | number | null)),
  }))
  protected readonly setup = signal<Setup>({ plans: [], insuredTypes: [], healthTypes: [], payors: [], groupCompanies: [] })
  protected readonly rows = signal<Row[]>([])
  protected readonly selected = signal<string | null>(null)
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: Row) => r._id
  protected v: Limit = { ...EMPTY }
  private index: number | null = null
  private lastFilters: unknown = null

  async ngOnInit(): Promise<void> {
    await this.refresh()
  }

  protected locked(): boolean {
    return this.mode() === 'selected'
  }

  protected entering(): boolean {
    return this.mode() === 'add' || this.mode() === 'edit'
  }

  /** CmdRefresh_Click → VisitComboListing (Supp. status defaults to A). */
  protected async refresh(): Promise<void> {
    this.reset()
    this.mode.set('search')
    try {
      this.setup.set(await firstValueFrom(this.http.get<Setup>('/api/maintenance/plan/annual-limits/setup')))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** TxtVLimitEffDate_LostFocus. */
  protected checkDate(): void {
    if (this.v.effDate && !parseDmy(this.v.effDate)) this.notify.notify('Invalid Date Format! ', 'error')
  }

  private body(): unknown {
    return { limit: { ...this.v, index: this.index } }
  }

  /** SearchVisit → VisitListing. */
  protected async search(body: unknown = this.body()): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.selected.set(null)
    try {
      const rows = (await firstValueFrom(this.http.post<Record<string, unknown>[] | null>(`${BASE}/search`, body))) ?? []
      this.rows.set(rows.map((r, i) => ({ ...r, _id: String(i) })))
      this.lastFilters = body
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
      this.mode.set('search')
    }
  }

  /** lstVisitLimit_Click: the row into the boxes, read-only. */
  protected select(r: Row): void {
    this.selected.set(r._id)
    const s = (k: string) => (r[k] == null ? '' : String(r[k]).trim())
    this.index = Number(r['VisitIndex'])
    this.v = {
      plnCode: s('PLNCode'), hltCode: s('HLTCode'), insCode: s('INSCode'), payCode: s('PAYCode'), grpCompany: s('GRPCompany'),
      visitLimit: r['VisitLimit'] == null ? '' : formatMoney(r['VisitLimit']), suppStatus: s('SuppVisitLimitStatus') || 'A',
      suppLimit: r['SuppVisitLimit'] == null ? '' : formatMoney(r['SuppVisitLimit']), effDate: formatDate(r['VisitEffDate']), version: s('VisitLimitVersion'),
    }
    this.mode.set('selected')
  }

  protected add(): void {
    this.reset()
    this.selected.set(null)
    this.mode.set('add')
  }

  protected clear(): void {
    this.reset()
    this.rows.set([])
    this.selected.set(null)
    this.mode.set('search')
  }

  /** SaveVisit, then VisitListing with the last criteria. */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to save the record?'))) return
    const editing = this.mode() === 'edit'
    this.busy.set(true)
    try {
      const answer = editing
        ? await firstValueFrom(this.http.put<{ message: string }>(BASE, this.body()))
        : await firstValueFrom(this.http.post<{ message: string }>(BASE, this.body()))
      this.reset()
      if (this.lastFilters) await this.search(this.lastFilters)
      this.notify.notify(answer.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** DeleteVisit, then the whole list. */
  protected async remove(): Promise<void> {
    if (!(await this.notify.ask('Do you want to delete the record?'))) return
    try {
      const answer = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/delete`, this.body()))
      this.notify.notify(answer.message, 'success')
      this.reset()
      await this.search({ limit: { ...EMPTY, suppStatus: '' } })
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** PrintVisit → RptVisitList. */
  protected async print(): Promise<void> {
    try {
      openBlob(await firstValueFrom(this.http.get('/api/maintenance/list-reports/visit-limit', { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  private reset(): void {
    this.v = { ...EMPTY }
    this.index = null
  }
}
