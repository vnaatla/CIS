import { HttpClient, HttpParams } from '@angular/common/http'
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
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

const BASE = '/api/maintenance/plan/premiums'

interface Premium {
  insCode: string
  hltCode: string
  payCode: string
  ageCode: string
  plnCode: string
  effDate: string
  version: string
  suppStatus: string
}
const EMPTY: Premium = { insCode: '', hltCode: '', payCode: '', ageCode: '', plnCode: '', effDate: '', version: '', suppStatus: 'A' }
type Row = Record<string, unknown> & { _id: string }
type Mode = 'search' | 'selected' | 'add'
interface Setup {
  insuredTypes: LookupOption[]
  healthTypes: LookupOption[]
  payors: LookupOption[]
  ages: LookupOption[]
  plans: LookupOption[]
}

/** frmMaintenancePlan tab 3 "Premium" (CISDB PRM): keys, effective date, version, supplementary status and the seven age bands. */
@Component({
  selector: 'app-plan-premium-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Premium">
        <legend>Premium</legend>
        <app-field label="Insured Type" [required]="true" hint="Searchable">
          <select name="insCode" [disabled]="locked()" [(ngModel)]="p.insCode">
            <option value=""></option>
            @for (o of setup().insuredTypes; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Health Type" [required]="true" hint="Searchable">
          <select name="hltCode" [disabled]="locked()" [ngModel]="p.hltCode" (ngModelChange)="pickHealth($event)">
            <option value=""></option>
            @for (o of setup().healthTypes; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Payor" hint="Searchable">
          <select name="payCode" [disabled]="locked()" [(ngModel)]="p.payCode">
            <option value=""></option>
            @for (o of payors(); track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Age" hint="Searchable">
          <select name="ageCode" [disabled]="locked()" [(ngModel)]="p.ageCode">
            <option value=""></option>
            @for (o of setup().ages; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Plan Code" [required]="true" hint="Searchable">
          <input name="plnCode" list="premium-plans" [disabled]="locked()" [(ngModel)]="p.plnCode" />
          <datalist id="premium-plans">
            @for (o of setup().plans; track o.code) {
              <option [value]="o.code + ' - ' + o.name"></option>
            }
          </datalist>
        </app-field>
        <app-field label="Premium Eff. Date" [required]="true">
          <app-date-input name="effDate" ariaLabel="Premium Eff. Date" [disabled]="locked()" [(ngModel)]="p.effDate" />
        </app-field>
        <app-field label="Version">
          <input name="version" maxlength="9" inputmode="numeric" [disabled]="locked()" [(ngModel)]="p.version" />
        </app-field>
        <app-field label="Supplementary">
          <select name="suppStatus" [disabled]="locked()" [(ngModel)]="p.suppStatus">
            <option value="A">A - Share Limit for Family</option>
            <option value="B">B - Share Limit for Supplementary</option>
            <option value="C">C - Separate Limit for Supplementary</option>
          </select>
        </app-field>
        <div class="entry-table">
          <table class="grid" aria-label="Age band premiums">
            <thead><tr><th scope="col">Age Band</th><th scope="col" class="num">Premium</th><th scope="col" class="num">Supplementary Premium</th></tr></thead>
            <tbody>
              @for (i of bands; track i) {
                <tr>
                  <td>0{{ i + 1 }}</td>
                  <td><input class="num" [name]="'amt' + i" [attr.name]="'amt' + i" inputmode="decimal" [disabled]="locked()" [(ngModel)]="amounts[i]" /></td>
                  <td><input class="num" [name]="'supp' + i" [attr.name]="'supp' + i" inputmode="decimal" [disabled]="locked() || p.suppStatus === 'A'" [(ngModel)]="supps[i]" /></td>
                </tr>
              }
            </tbody>
          </table>
        </div>
      </fieldset>

      <app-action-bar label="Premium actions">
        <button appButton variant="primary" [busy]="busy()" [disabled]="mode() === 'add'" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [disabled]="true" title="The desktop's Edit saved nothing on this tab">Edit</button>
        <button appButton variant="danger" [disabled]="mode() !== 'selected'" (click)="remove()">Delete</button>
        <button appButton [disabled]="mode() !== 'add'" [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <button appButton variant="ghost" (click)="refresh()">Refresh</button>
        <button appButton variant="ghost" (click)="print()">Print</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="rows().length" />
      </app-action-bar>

      <app-data-grid caption="Premiums" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
        [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class PlanPremiumPanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly bands = [0, 1, 2, 3, 4, 5, 6]
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'PRMCode', label: 'Premium Code', align: 'right' },
    { key: 'INSCode', label: 'Insured Type' },
    { key: 'PAYCode', label: 'Payor' },
    { key: 'HLTCode', label: 'Health Type' },
    { key: 'AGECode', label: 'Age Code' },
    { key: 'PLNCode', label: 'Plan Code' },
    { key: 'PRMAmount', label: 'Premium', align: 'right', value: (r) => formatMoney(r['PRMAmount']) },
    { key: 'SuppPRMAmt', label: 'Supp. Premium', align: 'right', value: (r) => formatMoney(r['SuppPRMAmt']) },
    { key: 'PRMEffDate', label: 'Eff. Date', value: (r) => formatDate(r['PRMEffDate']), sortBy: (r) => String(r['PRMEffDate'] ?? '') },
    { key: 'PRMVersion', label: 'Version', align: 'right' },
    { key: 'SuppPRMStatus', label: 'Supp. Status' },
  ]
  protected readonly setup = signal<Setup>({ insuredTypes: [], healthTypes: [], payors: [], ages: [], plans: [] })
  protected readonly payors = signal<LookupOption[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly selected = signal<string | null>(null)
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: Row) => r._id
  protected p: Premium = { ...EMPTY }
  protected amounts = ['', '', '', '', '', '', '']
  protected supps = ['', '', '', '', '', '', '']
  private picked: Row | null = null
  private lastFilters: unknown = null

  async ngOnInit(): Promise<void> {
    await this.refresh()
  }

  protected locked(): boolean {
    return this.mode() === 'selected'
  }

  /** CmdRefresh_Click → ClearForm + PRMComboListing (Supplementary defaults to A). */
  protected async refresh(): Promise<void> {
    this.reset()
    this.mode.set('search')
    try {
      const s = await firstValueFrom(this.http.get<Setup>(`${BASE}/setup`))
      this.setup.set(s)
      this.payors.set(s.payors)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cboPmHLTCode_Change: payors of the health type. */
  protected async pickHealth(code: string): Promise<void> {
    this.p.hltCode = code
    if (!code) return this.payors.set(this.setup().payors)
    try {
      const params = new HttpParams({ fromObject: { healthType: code } })
      this.payors.set((await firstValueFrom(this.http.get<LookupOption[] | null>('/api/maintenance/plan/codes/payors', { params }))) ?? [])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private body(): unknown {
    return { premium: { ...this.p, amounts: this.amounts, suppAmounts: this.supps } }
  }

  /** SearchPRM → PRMListing. */
  protected async search(body: unknown = this.body()): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.selected.set(null)
    this.picked = null
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

  /** lstPMPRMList_Click: the row's keys and the insured type + plan's band amounts, read-only. */
  protected async select(r: Row): Promise<void> {
    this.selected.set(r._id)
    this.picked = r
    const s = (k: string) => (r[k] == null ? '' : String(r[k]).trim())
    this.p = {
      insCode: s('INSCode'), hltCode: s('HLTCode'), payCode: s('PAYCode'), ageCode: s('AGECode'), plnCode: s('PLNCode'),
      effDate: formatDate(r['PRMEffDate']), version: s('PRMVersion'), suppStatus: s('SuppPRMStatus') || 'A',
    }
    this.amounts = ['', '', '', '', '', '', '']
    this.supps = ['', '', '', '', '', '', '']
    try {
      const params = new HttpParams({ fromObject: { insCode: this.p.insCode, plnCode: this.p.plnCode } })
      const bands = (await firstValueFrom(this.http.get<{ ageCode: string; amount: number | null; suppAmount: number | null }[] | null>(`${BASE}/bands`, { params }))) ?? []
      for (const b of bands) {
        const i = Number(b.ageCode) - 1
        if (i < 0 || i > 6 || b.ageCode.length !== 2) continue
        this.amounts[i] = formatMoney(b.amount ?? 0)
        if (b.suppAmount != null) this.supps[i] = formatMoney(b.suppAmount)
      }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
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

  /** SavePRM, then PRMListing with the last criteria. */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to save this record?'))) return
    this.busy.set(true)
    try {
      const answer = await firstValueFrom(this.http.post<{ message: string }>(BASE, this.body()))
      this.reset()
      if (this.lastFilters) await this.search(this.lastFilters)
      this.notify.notify(answer.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** DeletePRM (the selected PRM row), then PRMListing. */
  protected async remove(): Promise<void> {
    if (!this.picked) return this.notify.notify('Please search the record', 'error')
    if (!(await this.notify.ask('Do you want to delete this record?'))) return
    try {
      const answer = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/${this.picked['PRMCode']}/delete`, this.body()))
      this.notify.notify(answer.message, 'success')
      this.reset()
      await this.search({ premium: { ...EMPTY, suppStatus: '', amounts: [], suppAmounts: [] } })
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** PrintPRM → RptPremiumList. */
  protected async print(): Promise<void> {
    try {
      openBlob(await firstValueFrom(this.http.get('/api/maintenance/list-reports/premium', { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  private reset(): void {
    this.p = { ...EMPTY }
    this.amounts = ['', '', '', '', '', '', '']
    this.supps = ['', '', '', '', '', '', '']
    this.picked = null
  }
}
