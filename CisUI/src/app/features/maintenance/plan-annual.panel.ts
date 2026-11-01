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

const BASE = '/api/maintenance/plan/annual-limits'

type Key = 'plnCode' | 'hltCode' | 'insCode' | 'payCode' | 'grpCompany' | 'annLimit' | 'pDental' | 'pSpec' | 'pMep' | 'suppStatus' | 'suppLimit'
  | 'sDental' | 'sSpec' | 'sMep' | 'effDate' | 'version'
type Limit = Record<Key, string>
const EMPTY = {
  plnCode: '', hltCode: '', insCode: '', payCode: '', grpCompany: '', annLimit: '', pDental: '', pSpec: '', pMep: '', suppStatus: 'A', suppLimit: '',
  sDental: '', sSpec: '', sMep: '', effDate: '', version: '',
} as Limit
type Row = Record<string, unknown> & { _id: string }
type Mode = 'search' | 'selected' | 'add' | 'edit'
interface Setup {
  plans: LookupOption[]
  insuredTypes: LookupOption[]
  healthTypes: LookupOption[]
  payors: LookupOption[]
  groupCompanies: LookupOption[]
}
const MONEY: { key: Key; label: string; required?: boolean; supp?: boolean }[] = [
  { key: 'annLimit', label: 'Principal Annual Limit', required: true },
  { key: 'pDental', label: 'Principal Dental Annual Limit', required: true },
  { key: 'pSpec', label: 'Principal Spectacles Annual Limit', required: true },
  { key: 'pMep', label: 'Principal MEP Annual Limit', required: true },
  { key: 'suppLimit', label: 'Supp Premium Limit', supp: true },
  { key: 'sDental', label: 'Supplementary Dental Annual Limit', supp: true },
  { key: 'sSpec', label: 'Supplementary Spectacles Annual Limit', supp: true },
  { key: 'sMep', label: 'Supplementary MEP Annual Limit', supp: true },
]
/** lstAnnualList columns. */
const LIST: [string, string, boolean?][] = [
  ['HLTCode', 'Health Type'], ['AnnualIndex', 'Index'], ['PLNCode', 'Plan Code'], ['INSCode', 'Insured Type'], ['AnnLimit', 'Annual Limit', true],
  ['PDentalAnnLimit', 'Prin. Dental', true], ['PSpecAnnLimit', 'Prin. Spectacles', true], ['PMEPAnnLimit', 'Prin. MEP', true], ['SuppLimit', 'Supp. Limit', true],
  ['SDentalAnnLimit', 'Supp. Dental', true], ['SSpecAnnLimit', 'Supp. Spectacles', true], ['SMEPAnnLimit', 'Supp. MEP', true], ['AnnEffDate', 'Eff. Date'],
  ['AnnLimitVersion', 'Version'], ['PAYCode', 'Payor'], ['GRPCompany', 'Group Company'], ['SuppLimitStatus', 'Supp. Status'],
]

/** frmMaintenancePlan tab 4 "Annual Limit" (CISDB AnnualLimit / VIEW_ANN_PAY_PLN). */
@Component({
  selector: 'app-plan-annual-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Annual Limit">
        <legend>Annual Limit</legend>
        <app-field label="Plan Code" [required]="true" hint="Searchable">
          <input name="plnCode" list="annual-plans" [disabled]="locked()" [(ngModel)]="a.plnCode" />
          <datalist id="annual-plans">
            @for (o of setup().plans; track o.code) {
              <option [value]="o.code + ' - ' + o.name"></option>
            }
          </datalist>
        </app-field>
        @for (c of combos; track c.key) {
          <app-field [label]="c.label" [required]="!!c.required" hint="Searchable">
            <select [name]="c.key" [attr.name]="c.key" [disabled]="locked() || (!!c.searchOnly && entering())" [(ngModel)]="a[c.key]">
              <option value=""></option>
              @for (o of setup()[c.list]; track o.code) {
                <option [value]="o.code">{{ o.code === o.name ? o.code : o.code + ' - ' + o.name }}</option>
              }
            </select>
          </app-field>
        }
        <app-field label="Supplementary">
          <select name="suppStatus" [disabled]="locked()" [(ngModel)]="a.suppStatus">
            <option value="A">A - Share Limit For Family</option>
            <option value="B">B - Share Limit For Supplementary</option>
            <option value="C">C - Separate Limit For Supplementary</option>
          </select>
        </app-field>
        @for (m of money; track m.key) {
          <app-field [label]="m.label" [required]="!!m.required || (!!m.supp && a.suppStatus !== 'A')">
            <input class="num" [name]="m.key" [attr.name]="m.key" inputmode="decimal"
              [disabled]="locked() || (m.key === 'suppLimit' && a.suppStatus === 'A')" [(ngModel)]="a[m.key]" />
          </app-field>
        }
        <app-field label="Annual Limit Eff. Date" [required]="true">
          <app-date-input name="effDate" ariaLabel="Annual Limit Eff. Date" [disabled]="locked()" [(ngModel)]="a.effDate" (focusout)="checkDate()" />
        </app-field>
        <app-field label="Annual Limit Version" [required]="true">
          <input name="version" maxlength="9" inputmode="numeric" [disabled]="locked()" [(ngModel)]="a.version" />
        </app-field>
      </fieldset>

      <app-action-bar label="Annual Limit actions">
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

      <app-data-grid caption="Annual limits" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
        [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class PlanAnnualPanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly combos: { key: Key; label: string; list: keyof Setup; required?: boolean; searchOnly?: boolean }[] = [
    { key: 'hltCode', label: 'Health Type', list: 'healthTypes', required: true },
    { key: 'insCode', label: 'Insured Type', list: 'insuredTypes', required: true },
    { key: 'payCode', label: 'Payor Code', list: 'payors', searchOnly: true },
    { key: 'grpCompany', label: 'Group Company', list: 'groupCompanies', searchOnly: true },
  ]
  protected readonly money = MONEY
  protected readonly columns: GridColumn<Row>[] = LIST.map(([key, label, isMoney]) => ({
    key, label, align: isMoney ? 'right' : undefined,
    value: (r: Row) => (isMoney ? formatMoney(r[key]) : key === 'AnnEffDate' ? formatDate(r[key]) : (r[key] as string | number | null)),
  }))
  protected readonly setup = signal<Setup>({ plans: [], insuredTypes: [], healthTypes: [], payors: [], groupCompanies: [] })
  protected readonly rows = signal<Row[]>([])
  protected readonly selected = signal<string | null>(null)
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: Row) => r._id
  protected a: Limit = { ...EMPTY }
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

  /** CmdRefresh_Click → AnnualComboListing (Supplementary defaults to A). */
  protected async refresh(): Promise<void> {
    this.reset()
    this.mode.set('search')
    try {
      this.setup.set(await firstValueFrom(this.http.get<Setup>(`${BASE}/setup`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** txtALEffDate_LostFocus. */
  protected checkDate(): void {
    if (this.a.effDate && !parseDmy(this.a.effDate)) this.notify.notify('Invalid Date Format! ', 'error')
  }

  private body(): unknown {
    return { limit: { ...this.a, index: this.index } }
  }

  /** SearchANN → AnnualListing. */
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

  /** lstAnnualList_Click: the row into the boxes, read-only. */
  protected select(r: Row): void {
    this.selected.set(r._id)
    const s = (k: string) => (r[k] == null ? '' : String(r[k]).trim())
    const m = (k: string) => (r[k] == null ? '' : formatMoney(r[k]))
    this.index = Number(r['AnnualIndex'])
    this.a = {
      plnCode: s('PLNCode'), hltCode: s('HLTCode'), insCode: s('INSCode'), payCode: s('PAYCode'), grpCompany: s('GRPCompany'), annLimit: m('AnnLimit'),
      pDental: m('PDentalAnnLimit'), pSpec: m('PSpecAnnLimit'), pMep: m('PMEPAnnLimit'), suppStatus: s('SuppLimitStatus') || 'A', suppLimit: m('SuppLimit'),
      sDental: m('SDentalAnnLimit'), sSpec: m('SSpecAnnLimit'), sMep: m('SMEPAnnLimit'), effDate: formatDate(r['AnnEffDate']), version: s('AnnLimitVersion'),
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

  /** SaveANN, then AnnualListing with the last criteria. */
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

  /** DeleteANN, then the whole list (AnnualListing without criteria). */
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

  /** PrintANN → RptAnnualList. */
  protected async print(): Promise<void> {
    try {
      openBlob(await firstValueFrom(this.http.get('/api/maintenance/list-reports/annual-limit', { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  private reset(): void {
    this.a = { ...EMPTY }
    this.index = null
  }
}
