import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/tracking/claim'
type Row = Record<string, string>
const TYPES = ['Admission', 'Ageband', 'Agent No', 'Bordx', 'Clinic', 'Clinic Group', 'Department', 'Diagnosis', 'Gender', 'General Listing', 'Group Company', 'I/C',
  'Insured Type', 'Invoice Listing', 'MC', 'Month', 'Monthly Reports', 'Payor', 'Plan Code', 'Policy', 'Race', 'State', 'Utilization', 'Verification of Claims']
const RANGES: { key: string; label: string; date?: boolean; blank?: boolean }[] = [
  { key: 'dov', label: 'DOV', date: true, blank: true }, { key: 'dor', label: 'DOR', date: true, blank: true }, { key: 'dow', label: 'Claim Open Date', date: true, blank: true },
  { key: 'invoiceDate', label: 'Invoice Date', date: true, blank: true }, { key: 'bordxDate', label: 'Bordx Date', date: true, blank: true },
  { key: 'eff', label: 'Payor Eff. Date', date: true, blank: true }, { key: 'exp', label: 'Payor Exp. Date', date: true, blank: true },
  { key: 'joined', label: 'Date Joined', date: true, blank: true }, { key: 'dob', label: 'DOB', date: true, blank: true },
  { key: 'caseNo', label: 'Case No' }, { key: 'mbmNo', label: 'Member No' }, { key: 'bordx', label: 'Bordx No', blank: true }, { key: 'invoice', label: 'Invoice No', blank: true },
  { key: 'submission', label: 'Submission No', blank: true }, { key: 'opBatch', label: 'OP Batch No' }, { key: 'clinicCode', label: 'Clinic Code' },
  { key: 'plan', label: 'Plan Code', blank: true }, { key: 'ageband', label: 'Ageband' }, { key: 'polYear', label: 'Policy Year', blank: true },
  { key: 'mcDays', label: 'MC Days', blank: true }, { key: 'treatment', label: 'Treatment Reason', blank: true }, { key: 'ic', label: 'Patient IC No', blank: true },
  { key: 'mbmIc', label: 'Member IC No', blank: true }, { key: 'actual', label: 'Actual Amount' }, { key: 'approved', label: 'Approved Amount' },
  { key: 'osUtilized', label: 'O/S Utilized' },
]
const COMBOS: { key: string; label: string; items: string[]; blank?: boolean }[] = [
  { key: 'hlt', label: 'Health Type', items: ['S - Sihat', 'N - Non-Sihat'], blank: true },
  { key: 'ins', label: 'Insured Type', items: ['I - Individual', 'F - Family', 'G - Group Individual', 'H - Group Family'], blank: true },
  { key: 'caseStatus', label: 'Case Status', items: ['0 - Open', '1 - Open & Close', '2 - Delete', '3 - All'], blank: true },
  { key: 'claimability', label: 'Claimability', items: ['A - Accepted', 'R - Rejected'], blank: true },
  { key: 'race', label: 'Race', items: ['C - Chinese', 'I - Indian', 'M - Malay', 'O - Other'], blank: true },
  { key: 'sex', label: 'Sex', items: ['M - Male', 'F - Female'], blank: true },
  { key: 'relation', label: 'Relation', items: ['D - Daughter', 'H - Husband', 'S - Son', 'W - Wife'], blank: true },
  { key: 'dataStatus', label: 'Data Status', items: ['A - Active', 'D - Delete'], blank: true },
  { key: 'expired', label: 'Expired', items: ['N - No', 'Y - Yes'], blank: true },
  { key: 'premMode', label: 'Premium Mode', items: ['A - Annually', 'Q - Quarterly', 'S - Semi Annual', 'M - Monthly'], blank: true },
  { key: 'renewal', label: 'Renewal', items: ['N - New', 'P - Partial', 'R - Renew', 'T - Timely'], blank: true },
  { key: 'takeover', label: 'Takeover', items: ['N - None', 'Y - Take Over', 'C - Conversion'], blank: true },
  { key: 'panel', label: 'Panel Status', items: ['C - Confirmed', 'Y - Yes'] }, { key: 'pending', label: 'Claim Pending', items: ['N - No', 'Y - Yes'] },
  { key: 'clinicType', label: 'Clinic Type', items: ['0 - Specialised', '1 - General Practisional', '2 - Others'] },
  { key: 'payTo', label: 'Pay To', items: ['C - Clinic', 'M - Member'] },
]
const TEXTS = [{ key: 'agent', label: 'Agent Code', blank: true }, { key: 'policy', label: 'Policy No', blank: true }, { key: 'patient', label: 'Patient Name' },
  { key: 'branch', label: 'Branch' }, { key: 'year', label: 'Year (DOV)' }]
const code = (v: string) => (v.includes(' - ') ? v.slice(0, v.indexOf(' - ')) : v).trim()

/** Claim / Bordx > Claim Tracking (frmClaimRpt). */
@Component({
  selector: 'app-claim-tracking-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Claim Tracking" subtitle="Replaces frmClaimRpt">
      <div class="screen-form">
        <fieldset aria-label="Report">
          <legend>Report</legend>
          <app-field label="Report Type">
            <select name="reportType" [ngModel]="type()" (ngModelChange)="setType($event)">
              <option [ngValue]="-1"></option>@for (t of types; track $index) { <option [ngValue]="$index">{{ t }}</option> }
            </select>
          </app-field>
          <app-field label="Report Selection" [wide]="true">
            <select name="selection" [(ngModel)]="selection">
              @for (r of selections(); track r.selection) { <option [value]="r.selection">{{ r.selection || '(all)' }}{{ r.migrated ? '' : ' — not migrated yet' }}</option> }
            </select>
          </app-field>
          <app-field label="Member Type">
            <select name="memType" [(ngModel)]="memType"><option>Both</option><option>P - Principal</option><option>S - Supplementary</option></select>
          </app-field>
          <app-field label="Supp. Type">
            <select name="suppType" [(ngModel)]="suppType"><option>Both</option><option>C - Combined Limit</option><option>S - Separated Limit</option></select>
          </app-field>
          <app-field label="Bordx Type">
            <select name="bordxType" [(ngModel)]="f['bordxType']"><option>INSURER</option><option>ASO</option></select>
          </app-field>
          <app-action-bar label="Tracking actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <fieldset aria-label="Claim">
          <legend>Claim and member</legend>
          @for (t of texts; track t.key) {
            <app-field [label]="t.label"><input [name]="t.key" [attr.name]="t.key" maxlength="60" [(ngModel)]="f[t.key]" /></app-field>
            @if (t.blank) { <label class="check-field"><input type="checkbox" [name]="t.key + 'Blank'" [attr.name]="t.key + 'Blank'" [(ngModel)]="f[t.key + 'Blank']" /> Empty</label> }
          }
          @for (c of combos; track c.key) {
            <app-field [label]="c.label">
              <select [name]="c.key" [attr.name]="c.key" [(ngModel)]="f[c.key]"><option value=""></option>@for (i of c.items; track i) { <option [value]="i">{{ i }}</option> }</select>
            </app-field>
            @if (c.blank) { <label class="check-field"><input type="checkbox" [name]="c.key + 'Blank'" [attr.name]="c.key + 'Blank'" [(ngModel)]="f[c.key + 'Blank']" /> Empty</label> }
          }
          <app-field label="Policy Status">
            <select name="polStatus" [(ngModel)]="f['polStatus']"><option value=""></option>@for (i of lists()['polstatus'] ?? []; track i) { <option [value]="i">{{ i }}</option> }</select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="polStatusBlank" [(ngModel)]="f['polStatusBlank']" /> Empty</label>
          <app-field label="State">
            <select name="state" [(ngModel)]="f['state']"><option value=""></option>@for (i of lists()['states'] ?? []; track i) { <option [value]="i">{{ i }}</option> }</select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="stateBlank" [(ngModel)]="f['stateBlank']" /> Empty</label>
          <label class="check-field"><input type="checkbox" name="vsAmt" [(ngModel)]="f['vsAmt']" /> Amount (VS) not zero</label>
          <label class="check-field"><input type="checkbox" name="ieInvoice" [(ngModel)]="f['ieInvoice']" /> IE invoices only</label>
        </fieldset>
        <fieldset aria-label="Payor">
          <legend>Payor, company and clinic</legend>
          <app-field label="Payor" [wide]="true">
            <input name="payor" list="ct-payors" (focus)="payors()" (change)="companies()" [(ngModel)]="f['payor']" />
            <datalist id="ct-payors">@for (c of lists()['payors'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="payorBlank" [(ngModel)]="f['payorBlank']" /> Empty</label>
          <app-field label="Group Company" [wide]="true">
            <input name="company" list="ct-companies" (change)="departments()" [(ngModel)]="f['company']" />
            <datalist id="ct-companies">@for (c of lists()['companies'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="companyBlank" [(ngModel)]="f['companyBlank']" /> Empty</label>
          <app-field label="Department">
            <select name="deptSelection" [(ngModel)]="f['deptSelection']"><option value=""></option>@for (c of lists()['departments'] ?? []; track $index) { <option [value]="c">{{ c }}</option> }</select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="departmentBlank" [(ngModel)]="f['departmentBlank']" /> Empty</label>
          <app-field label="Clinic" [wide]="true">
            <input name="clinic" list="ct-clinics" (focus)="load('clinics')" [(ngModel)]="f['clinic']" />
            <datalist id="ct-clinics">@for (c of lists()['clinics'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="clinicBlank" [(ngModel)]="f['clinicBlank']" /> Empty</label>
          <app-field label="Clinic Group" [wide]="true">
            <input name="clinicGroup" list="ct-groups" (focus)="load('clinicgroups')" [(ngModel)]="f['clinicGroup']" />
            <datalist id="ct-groups">@for (c of lists()['clinicgroups'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
        </fieldset>
        <fieldset aria-label="Ranges">
          <legend>Ranges</legend>
          @for (r of ranges; track r.key) {
            @if (r.date) {
              <app-field [label]="r.label + ' From'"><app-date-input [name]="r.key + 'From'" [ariaLabel]="r.label + ' From'" [(ngModel)]="f[r.key + 'From']" /></app-field>
              <app-field [label]="r.label + ' To'"><app-date-input [name]="r.key + 'To'" [ariaLabel]="r.label + ' To'" [(ngModel)]="f[r.key + 'To']" /></app-field>
            } @else {
              <app-field [label]="r.label + ' From'"><input [name]="r.key + 'From'" [attr.name]="r.key + 'From'" maxlength="50" [(ngModel)]="f[r.key + 'From']" /></app-field>
              <app-field [label]="r.label + ' To'"><input [name]="r.key + 'To'" [attr.name]="r.key + 'To'" maxlength="50" [(ngModel)]="f[r.key + 'To']" /></app-field>
            }
            @if (r.blank) { <label class="check-field"><input type="checkbox" [name]="r.key + 'Blank'" [attr.name]="r.key + 'Blank'" [(ngModel)]="f[r.key + 'Blank']" /> Empty</label> }
          }
        </fieldset>
        <app-data-grid caption="Claim tracking" [columns]="columns()" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">Records: <strong>{{ records() }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class ClaimTrackingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly types = TYPES
  protected readonly ranges = RANGES
  protected readonly combos = COMBOS
  protected readonly texts = TEXTS
  protected readonly reports = signal<{ type: number; selection: string; migrated: boolean }[]>([])
  protected readonly type = signal(-1)
  protected readonly selections = computed(() => this.reports().filter((r) => r.type === this.type()))
  protected selection = ''
  protected memType = 'Both'
  protected suppType = 'Both'
  protected readonly lists = signal<Record<string, string[]>>({})
  protected readonly headers = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly records = signal(0)
  protected readonly busy = signal(false)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected f: any = this.empty()
  protected readonly key = (x: Row) => `${this.rows().indexOf(x)}`
  protected readonly columns = computed<GridColumn<Row>[]>(() => this.headers().map((h, i) => ({ key: `c${i}`, label: h || ' ' })))

  private empty(): Record<string, unknown> {
    return { bordxType: 'INSURER', caseStatus: '1 - Open & Close' }
  }

  async ngOnInit(): Promise<void> {
    try {
      this.reports.set(await firstValueFrom(this.http.get<{ type: number; selection: string; migrated: boolean }[]>(`${BASE}/reports`)))
      await Promise.all([this.load('polstatus'), this.load('states')])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async load(kind: string, key?: string): Promise<void> {
    if (!key && this.lists()[kind]) return
    const params = key ? new HttpParams().set('key', key) : undefined
    try {
      const items = await firstValueFrom(this.http.get<string[]>(`${BASE}/lookup/${kind}`, { params }))
      this.lists.update((l) => ({ ...l, [kind]: items }))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected setType(t: number): void {
    this.type.set(Number(t))
    this.selection = this.selections()[0]?.selection ?? ''
    this.rows.set([])
  }

  /** chkPay: the payors of the health type. */
  protected payors(): void {
    void this.load('payors', code(this.f['hlt'] ?? '') || 'S')
  }

  /** ChkGrpCom2: the group companies of the payor. */
  protected companies(): void {
    if (this.f['payor']) void this.load('companies', code(this.f['payor']))
  }

  /** chkDept2: the departments of the group company. */
  protected departments(): void {
    if (this.f['company']) void this.load('departments', this.f['company'])
  }

  protected async search(): Promise<void> {
    if (this.type() < 0) return this.notify.notify('Please select a report type.', 'info')
    const filters: Record<string, string> = {}
    for (const [k, v] of Object.entries(this.f)) {
      if (v === '' || v == null || v === false) continue
      const combo = [...COMBOS.map((c) => c.key), 'polStatus', 'payor', 'clinic'].includes(k)
      filters[k] = typeof v === 'boolean' ? 'true' : combo ? code(String(v)) : String(v)
    }
    if (filters['caseStatus'] === '3') delete filters['caseStatus']
    const mem = this.memType === 'Both' ? 'Both' : code(this.memType)
    const supp = this.suppType === 'Both' ? 'Both' : code(this.suppType)
    if (!(mem === 'Both' && supp === 'Both')) filters['memType'] = `${mem}|${supp}`
    this.busy.set(true)
    this.rows.set([])
    try {
      const r = await firstValueFrom(this.http.post<{ headers: string[]; rows: string[][]; records: number }>(`${BASE}/search`,
        { reportType: this.type(), selection: this.selection, filters }))
      this.headers.set(r.headers)
      this.rows.set(r.rows.map((cells) => Object.fromEntries(cells.map((c, i) => [`c${i}`, c]))))
      this.records.set(r.records)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.headers(), ...this.rows().map((x) => this.headers().map((_, i) => x[`c${i}`] ?? ''))], 'ClaimTracking.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
    this.records.set(0)
  }
}
