import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { Paged, PageRequest, ServerPaging } from '../../shared/lib/server-paging'
import type { LookupOption } from '../../shared/code-select.component'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { openBlob } from '../../shared/lib/download'
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

const BASE = '/api/maintenance/plan/codes'

type Key =
  | 'healthType' | 'payCode' | 'hisPayCode' | 'groupCompany' | 'clnContract' | 'annLmt' | 'visLmt' | 'prm' | 'mco' | 'sAnnLmt' | 'sVisLmt'
  | 'dcStatus' | 'noOfPlans' | 'insured' | 'hrApproval' | 'opFamilyStatus' | 'prinSpLimit' | 'prinGpLimit' | 'depSpLimit' | 'depGpLimit'
  | 'specialRemark' | 'opDependent' | 'pDentalStatus' | 'pDentalLimit' | 'pSpecStatus' | 'pSpecLimit' | 'pMepStatus' | 'sDentalStatus'
  | 'sDentalLimit' | 'sSpecStatus' | 'sSpecLimit' | 'sMepStatus'
type Plan = Record<Key, string>
type Mode = 'search' | 'selected' | 'add'
type Row = Record<string, unknown> & { _id: string }

const YESNO: [string, string][] = [['Y', 'Yes'], ['N', 'No']]
const STATUS: [string, string][] = [['C', 'C - Combine'], ['S', 'S - Stand Alone'], ['N', 'N - Not Applicable']]
const LIMIT: [string, string][] = [['0', '0 - Extraction Limit'], ['1', '1 - Others Limit'], ['2', '2 - No'], ['3', '3 - Yes']]

interface Pick {
  key: Key
  label: string
  options: [string, string][]
  required?: boolean
  /** Shown and searchable; Save always writes N (SavePLN). */
  searchOnly?: boolean
}
const PRINCIPAL: Pick[] = [
  { key: 'annLmt', label: 'Ann Lmt', options: YESNO, required: true },
  { key: 'visLmt', label: 'Vst Lmt', options: YESNO, required: true },
  { key: 'prm', label: 'Prem', options: YESNO, required: true },
  { key: 'mco', label: 'M C O', options: YESNO, required: true },
  { key: 'pDentalStatus', label: 'Dental Status', options: STATUS, searchOnly: true },
  { key: 'pDentalLimit', label: 'Dental Limit', options: LIMIT, searchOnly: true },
  { key: 'pSpecStatus', label: 'Spect Status', options: STATUS, searchOnly: true },
  { key: 'pSpecLimit', label: 'Spect Lmt', options: [['0', '0 - First Limit'], ...LIMIT.slice(1)], searchOnly: true },
  { key: 'pMepStatus', label: 'MEP Status', options: STATUS, searchOnly: true },
]
const SUPPLEMENTARY: Pick[] = [
  { key: 'sAnnLmt', label: 'Ann Lmt', options: YESNO, required: true },
  { key: 'sVisLmt', label: 'Vst Lmt', options: YESNO, required: true },
  { key: 'sDentalStatus', label: 'Dental Status', options: STATUS, searchOnly: true },
  { key: 'sDentalLimit', label: 'Dental Limit', options: [['0', '0 - No Extraction Limit'], ['1', '1 - No Others Limit'], ['2', '2 - No'], ['3', '3 - Yes']], searchOnly: true },
  { key: 'sSpecStatus', label: 'Spec. Status', options: STATUS, searchOnly: true },
  { key: 'sSpecLimit', label: 'Spec. Limit', options: [['0', '0 - No Extraction Limit'], ['1', '1 - No Others Limit'], ['2', '2 - No'], ['3', '3 - Yes']], searchOnly: true },
  { key: 'sMepStatus', label: 'MEP Status', options: STATUS, searchOnly: true },
]
/** Form_Load defaults of the three .frx combos. */
const EMPTY = {
  ...Object.fromEntries([...PRINCIPAL, ...SUPPLEMENTARY].map((p) => [p.key, ''])),
  healthType: '', payCode: '', hisPayCode: '', groupCompany: '', clnContract: '', dcStatus: '', noOfPlans: '', insured: '', hrApproval: 'N-No',
  opFamilyStatus: 'B-Both', prinSpLimit: '', prinGpLimit: '', depSpLimit: '', depGpLimit: '', specialRemark: '', opDependent: '',
} as Plan
/** PLNListing columns. */
const LIST: [string, string][] = [
  ['HLTCode', 'HLT Code'], ['PLNCode', 'Plan Code'], ['PLNDesc', 'Plan Description'], ['PAYCode', 'CIS Pay Code'], ['HISPAYCode', 'HIS Pay Code'],
  ['CoGRPCode', 'Group Company Code'], ['CoGRPName', 'Group Company Name'], ['ClnContract', 'Cln Contract'], ['Insured', 'Insured'],
  ['HRApproval', 'HR Approval'], ['PrinSPLimit', 'Prin. SP Limit'], ['PrinGPLimit', 'Prin. GP Limit'], ['DepSPLimit', 'Dep. SP Limit'],
  ['DepGPLimit', 'Dep. GP Limit'], ['OPFamilyStatus', 'OP Family Status'], ['PLNSpecialRemark', 'Special Remark'], ['isopdependent', 'OP Dependent'],
]

/** frmMaintenancePlan tab 2 "Plan Code" (CISDB PLN). */
@Component({
  selector: 'app-plan-code-panel',
  imports: [FormsModule, FieldComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Plan Code">
        <legend>Plan Code</legend>
        <app-field label="HLT Type" [required]="true" hint="Searchable">
          <select name="healthType" [disabled]="locked()" [ngModel]="p.healthType" (ngModelChange)="pickHealth($event)">
            <option value=""></option>
            @for (o of setup().healthTypes; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="CIS Pay Code" [required]="true" hint="Searchable">
          <select name="payCode" [disabled]="locked()" [ngModel]="p.payCode" (ngModelChange)="pickPayor($event)">
            <option value=""></option>
            @for (o of payors(); track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="HIS Pay Code" [required]="true" hint="Searchable">
          <select name="hisPayCode" [disabled]="locked()" [(ngModel)]="p.hisPayCode">
            <option value=""></option>
            @for (o of setup().hisPayors; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="DC Status" [required]="true">
          <select name="dcStatus" [disabled]="locked()" [ngModel]="p.dcStatus" (ngModelChange)="pickDc($event)">
            <option value=""></option>
            <option value="N">N - No</option>
            <option value="Y">Y - Yes</option>
          </select>
        </app-field>
        <app-field label="Grp Comp." [required]="true" hint="Searchable">
          <input name="groupCompany" list="plan-groups" [disabled]="locked()" [(ngModel)]="p.groupCompany" />
          <datalist id="plan-groups">
            @for (o of setup().groupCompanies; track o.code + o.name) {
              <option [value]="o.code + ' - ' + o.name"></option>
            }
          </datalist>
        </app-field>
        <app-field label="Cln Cont." [required]="true" hint="Searchable">
          <select name="clnContract" [disabled]="locked()" [(ngModel)]="p.clnContract">
            <option value=""></option>
            @for (o of setup().clinicContracts; track o.code) {
              <option [value]="o.code">{{ o.code }}</option>
            }
          </select>
        </app-field>
        <app-field label="No of Plan Code">
          <select name="noOfPlans" [disabled]="locked() || !nonSihat()" [(ngModel)]="p.noOfPlans">
            <option value=""></option>
            @for (n of ['1', '2', '3', '4', '5', '6']; track n) {
              <option [value]="n">{{ n }}</option>
            }
          </select>
        </app-field>
        <div class="entry-table">
          <table class="grid" aria-label="Plan codes and descriptions">
            <thead><tr><th scope="col">Plan Code</th><th scope="col">Plan Description</th></tr></thead>
            <tbody>
              @for (i of rows6; track i) {
                @if (i === 0 || !nonSihat()) {
                  <tr>
                    <td><input [name]="'plan' + i" [attr.name]="'plan' + i" maxlength="4" [disabled]="locked() || (nonSihat() && mode() !== 'search')" [(ngModel)]="codes[i]" /></td>
                    <td><input [name]="'desc' + i" [attr.name]="'desc' + i" maxlength="200" [disabled]="locked()" [(ngModel)]="descs[i]" /></td>
                  </tr>
                } @else if (+p.noOfPlans > i) {
                  <tr>
                    <td class="muted">(generated)</td>
                    <td><input [name]="'desc' + i" [attr.name]="'desc' + i" maxlength="200" [disabled]="locked()" [(ngModel)]="descs[i]" /></td>
                  </tr>
                }
              }
            </tbody>
          </table>
        </div>
      </fieldset>

      @for (group of groups; track group.title) {
        <fieldset [attr.aria-label]="group.title">
          <legend>{{ group.title }}</legend>
          @for (f of group.picks; track f.key) {
            <app-field [label]="f.label" [required]="!!f.required" [hint]="f.searchOnly ? 'Search only; Save writes N' : 'Searchable'">
              <select [name]="f.key" [attr.name]="f.key" [disabled]="locked()" [(ngModel)]="p[f.key]">
                <option value=""></option>
                @for (o of f.options; track o[0]) {
                  <option [value]="o[0]">{{ o[1] }}</option>
                }
              </select>
            </app-field>
          }
        </fieldset>
      }

      <fieldset aria-label="Other">
        <legend>Other</legend>
        <app-field label="Insured">
          <select name="insured" [disabled]="locked()" [(ngModel)]="p.insured"><option value=""></option><option value="N-No">N-No</option><option value="Y-Yes">Y-Yes</option></select>
        </app-field>
        <app-field label="HR approval">
          <select name="hrApproval" [disabled]="locked()" [(ngModel)]="p.hrApproval"><option value="N-No">N-No</option><option value="Y-Yes">Y-Yes</option></select>
        </app-field>
        <app-field label="OP Family Status">
          <select name="opFamilyStatus" [disabled]="locked()" [(ngModel)]="p.opFamilyStatus">
            <option value="B-Both">B-Both</option><option value="D-Dependent Only">D-Dependent Only</option><option value="E-Employee Only">E-Employee Only</option>
          </select>
        </app-field>
        <app-field label="OP Dependent" [required]="true">
          <select name="opDependent" [disabled]="locked()" [(ngModel)]="p.opDependent"><option value=""></option><option value="N">N - No</option><option value="Y">Y - Yes</option></select>
        </app-field>
        @for (l of limits; track l.key) {
          <app-field [label]="l.label">
            <input [name]="l.key" [attr.name]="l.key" maxlength="4" inputmode="decimal" [disabled]="locked()" [(ngModel)]="p[l.key]" />
          </app-field>
        }
        <app-field label="Special Remark" [wide]="true">
          <textarea name="specialRemark" rows="2" maxlength="1500" [disabled]="locked()" [(ngModel)]="p.specialRemark"></textarea>
        </app-field>
      </fieldset>

      <app-action-bar label="Plan Code actions">
        <button appButton variant="primary" [busy]="busy()" [disabled]="mode() === 'add'" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [disabled]="true" title="The desktop has no Edit for plans">Edit</button>
        <button appButton variant="danger" [disabled]="mode() !== 'selected'" (click)="remove()">Delete</button>
        <button appButton [disabled]="mode() !== 'add'" [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <button appButton variant="ghost" (click)="refresh()">Refresh</button>
        <button appButton variant="ghost" (click)="print()">Print</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="paging.total()" />
      </app-action-bar>

      <app-data-grid caption="Plans" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
        [serverTotal]="paging.total()" [serverPage]="paging.page()" (pageRequest)="onPage($event)"
        [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class PlanCodePanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly groups = [
    { title: 'Principal', picks: PRINCIPAL },
    { title: 'Supplementary', picks: SUPPLEMENTARY },
  ]
  protected readonly limits: { key: Key; label: string }[] = [
    { key: 'prinSpLimit', label: 'Principal SP Limit' },
    { key: 'prinGpLimit', label: 'Principal GP Limit' },
    { key: 'depSpLimit', label: 'Dependent SP Limit' },
    { key: 'depGpLimit', label: 'Dependent GP Limit' },
  ]
  protected readonly rows6 = [0, 1, 2, 3, 4, 5]
  protected readonly columns: GridColumn<Row>[] = LIST.map(([key, label]) => ({
    key, label, align: key.endsWith('Limit') ? 'right' : undefined, value: (r: Row) => r[key] as string | number | null,
  }))
  protected readonly setup = signal<{ healthTypes: LookupOption[]; payors: LookupOption[]; hisPayors: LookupOption[]; clinicContracts: LookupOption[]; groupCompanies: LookupOption[] }>(
    { healthTypes: [], payors: [], hisPayors: [], clinicContracts: [], groupCompanies: [] })
  protected readonly payors = signal<LookupOption[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly selected = signal<string | null>(null)
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: Row) => r._id
  protected p: Plan = { ...EMPTY }
  protected codes = ['', '', '', '', '', '']
  protected descs = ['', '', '', '', '', '']
  protected readonly paging = new ServerPaging()
  private lastFilters: unknown = null

  async ngOnInit(): Promise<void> {
    await this.refresh()
  }

  protected locked(): boolean {
    return this.mode() === 'selected'
  }

  protected nonSihat(): boolean {
    return this.p.healthType.startsWith('N')
  }

  /** CmdRefresh_Click → ClearForm + PLNComboListing. */
  protected async refresh(): Promise<void> {
    this.reset()
    this.mode.set('search')
    try {
      const s = await firstValueFrom(this.http.get<ReturnType<typeof this.setup>>(`${BASE}/setup`))
      this.setup.set(s)
      this.payors.set(s.payors)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cboPLNHLTCode_Click: CIS payors of the health type; N hides plan codes 2–6 and generates them. */
  protected async pickHealth(code: string): Promise<void> {
    this.p.healthType = code
    if (!code) return this.payors.set(this.setup().payors)
    try {
      const params = new HttpParams({ fromObject: { healthType: code } })
      this.payors.set((await firstValueFrom(this.http.get<LookupOption[] | null>(`${BASE}/payors`, { params }))) ?? [])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cboPLNPayor_Click: "DC" sets DC Status Y, anything else clears it. */
  protected pickPayor(code: string): void {
    this.p.payCode = code
    this.p.dcStatus = code.startsWith('DC') ? 'Y' : ''
  }

  /** CboDCStatus_Click: Y sets the payor to "DC - DIRECT CLIENT", N clears it. */
  protected pickDc(value: string): void {
    this.p.dcStatus = value
    this.p.payCode = value === 'Y' ? 'DC' : ''
  }

  private body(): unknown {
    return { plan: { ...this.p, planCodes: this.codes, descriptions: this.descs } }
  }

  /** SearchPLN → PLNListing. */
  protected async search(body: unknown = this.body(), page = 0): Promise<void> {
    this.busy.set(true)
    if (page === 0) this.rows.set([])
    this.selected.set(null)
    try {
      const result = await firstValueFrom(
        this.http.post<Paged<Record<string, unknown>> | null>(`${BASE}/search`, body, { params: new HttpParams({ fromObject: this.paging.query(page) }) }),
      )
      this.rows.set((result?.rows ?? []).map((r, i) => ({ ...r, _id: String(i) })))
      this.paging.loaded(page, result?.total ?? 0)
      this.lastFilters = body
    } catch (e) {
      this.paging.reset()
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
      this.mode.set('search')
    }
  }

  /** lstPlanList_Click: the plan's values, read-only (Edit stays disabled, as in the desktop). */
  protected async select(r: Row): Promise<void> {
    this.selected.set(r._id)
    try {
      const v = await firstValueFrom(this.http.get<Record<string, string | null>>(`${BASE}/${encodeURIComponent(String(r['PLNCode']))}`))
      const byKey = new Map(Object.entries(v).map(([k, x]) => [k.toLowerCase(), x]))
      const s = (k: string) => String(byKey.get(k.toLowerCase()) ?? '').trim()
      const insured = (x: string) => (x === 'Y' ? 'Y-Yes' : x === 'N' ? 'N-No' : '')
      this.p = {
        ...EMPTY, healthType: s('hltCode'), payCode: s('payCode'), hisPayCode: s('hisPayCode'),
        groupCompany: s('grpCompany') ? `${s('grpCompany')} - ${String(r['CoGRPName'] ?? '').trim()}` : '', clnContract: s('clnContract'),
        annLmt: s('annLmt'), visLmt: s('visLmt'), prm: s('prm'), mco: s('mco'), sAnnLmt: s('sAnnLmt'), sVisLmt: s('sVisLmt'), dcStatus: s('dcStatus'),
        pDentalStatus: s('pDentalStatus'), pDentalLimit: s('pDentalLimit'), pSpecStatus: s('pSpecStatus'), pSpecLimit: s('pSpecLimit'),
        pMepStatus: s('pMepStatus'), sDentalStatus: s('sDentalStatus'), sDentalLimit: s('sDentalLimit'), sSpecStatus: s('sSpecStatus'),
        sSpecLimit: s('sSpecLimit'), sMepStatus: s('sMepStatus'), insured: insured(s('insured')), hrApproval: insured(s('hrApproval')) || 'N-No',
        opFamilyStatus: s('opFamilyStatus') || 'B-Both', prinSpLimit: s('prinSpLimit'), prinGpLimit: s('prinGpLimit'), depSpLimit: s('depSpLimit'),
        depGpLimit: s('depGpLimit'), specialRemark: s('specialRemark'), opDependent: s('opDependent'),
      }
      this.codes = [s('plnCode'), '', '', '', '', '']
      this.descs = [s('plnDesc'), '', '', '', '', '']
      this.mode.set('selected')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected add(): void {
    this.reset()
    this.selected.set(null)
    this.mode.set('add')
  }

  /** Grid page / sort / size change: the last search, that page. */
  protected async onPage(r: PageRequest): Promise<void> {
    await this.search(this.lastFilters ?? this.body(), this.paging.change(r))
  }

  protected clear(): void {
    this.paging.reset()
    this.reset()
    this.rows.set([])
    this.selected.set(null)
    this.mode.set('search')
  }

  /** SavePLN, then PLNListing on the group company of the new plans (Strgrpcode). */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to save the record?'))) return
    this.busy.set(true)
    try {
      const group = this.p.groupCompany
      const answer = await firstValueFrom(this.http.post<{ message: string; codes: string[] }>(BASE, this.body()))
      this.reset()
      await this.search({ plan: { ...EMPTY, hrApproval: '', opFamilyStatus: '', groupCompany: group.split('-')[0].trim(), planCodes: [], descriptions: [] } })
      this.notify.notify(`${answer.message} (${answer.codes.join(', ')})`, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** DeletePLN, then the list again with the last criteria. */
  protected async remove(): Promise<void> {
    if (!(await this.notify.ask('Do you want to delete this record?'))) return
    try {
      const answer = await firstValueFrom(this.http.delete<{ message: string }>(`${BASE}/${encodeURIComponent(this.codes[0])}`))
      this.notify.notify(answer.message, 'success')
      this.reset()
      await this.search(this.lastFilters ?? this.body())
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** PrintPLN → RptPlanList. */
  protected async print(): Promise<void> {
    try {
      openBlob(await firstValueFrom(this.http.get('/api/maintenance/list-reports/plan', { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  private reset(): void {
    this.p = { ...EMPTY }
    this.codes = ['', '', '', '', '', '']
    this.descs = ['', '', '', '', '', '']
  }
}
