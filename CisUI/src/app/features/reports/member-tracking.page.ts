import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { ActivatedRoute } from '@angular/router'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

type Row = Record<string, string>
const TYPES = ['Ageband', 'Agents', 'Bordx', 'Branch', 'Date', 'Effective Period', 'Gender', 'General Listing', 'Group Company', 'IC/Pol Year/Renewal...',
  'INS Type & MBM Type', 'MBM By Month', 'MCO Invoice', 'Payor', 'Plan Code', 'Policy', 'Race', 'States']
/** The variant forms: their API base, title and form. */
const VARIANTS: Record<string, { base: string; title: string; form: string; types: string[] }> = {
  main: { base: '/api/tracking/member', title: 'MBM Tracking', form: 'FrmNewMBMTracking', types: TYPES },
  jupiter: { base: '/api/tracking/member-jupiter', title: 'MBM Tracking (Jupiter)', form: 'FrmNewMBMTracking2', types: TYPES },
  simple: { base: '/api/tracking/member-simple', title: 'MBM Tracking (Simple)', form: 'frmNewMBMTrackingSimple', types: [...TYPES, 'Head Count'] },
  jupitersimple: { base: '/api/tracking/member-jupiter-simple', title: 'MBM Tracking (Jupiter - Simple)', form: 'frmNewMBMTrackingRepSimple', types: TYPES.slice(0, 17) },
}
/** Ranges (from / to boxes) with their "empty" box; dates use the date input. */
const RANGES: { key: string; label: string; date?: boolean; blank?: boolean }[] = [
  { key: 'eff', label: 'Eff. Date', date: true, blank: true }, { key: 'exp', label: 'Exp. Date', date: true, blank: true },
  { key: 'joined', label: 'Date Joined', date: true, blank: true }, { key: 'dor', label: 'DOR', date: true, blank: true }, { key: 'dob', label: 'DOB', date: true, blank: true },
  { key: 'bordx', label: 'Bordx Date', date: true, blank: true }, { key: 'suppBordx', label: 'Supp. Bordx Date', date: true },
  { key: 'suppReg', label: 'DOR (Supp.)', date: true, blank: true },
  { key: 'cancel', label: 'Cancel Date', date: true, blank: true }, { key: 'invoiceDate', label: 'MCO Inv. Date', date: true, blank: true },
  { key: 'card', label: 'Card Sent Date', date: true, blank: true }, { key: 'endt', label: 'Endt. Eff. Date', date: true, blank: true },
  { key: 'plan', label: 'Plan Code', blank: true }, { key: 'age', label: 'Ageband', blank: true }, { key: 'ic', label: 'IC No', blank: true },
  { key: 'employee', label: 'Employee No', blank: true }, { key: 'batch', label: 'Batch No', blank: true }, { key: 'postcode', label: 'Postcode', blank: true },
  { key: 'polYear', label: 'Policy Year', blank: true }, { key: 'invoice', label: 'MCO Inv. No', blank: true }, { key: 'premium', label: 'Premium', blank: true },
  { key: 'mco', label: 'MCO Fee', blank: true }, { key: 'dayInsured', label: 'Day Insured', blank: true }, { key: 'ageDays', label: 'Age (Days)', blank: true },
  { key: 'delayBordx', label: 'Delay Bordx', blank: true }, { key: 'delayInvBordx', label: 'Delay Inv. Bordx', blank: true },
  { key: 'delayReg', label: 'Delay Reg.', blank: true }, { key: 'delayInvD', label: 'Delay Inv. D', blank: true }, { key: 'delayInvE', label: 'Delay Inv. E', blank: true },
]
/** Fixed combos (desktop items; the value sent is the code before " - "). */
const COMBOS: { key: string; label: string; items: string[]; blank?: boolean }[] = [
  { key: 'hlt', label: 'Health Type', items: ['S - Sihat', 'N - Non-Sihat'], blank: true },
  { key: 'status', label: 'Status', items: ['0 - Active', '1 - Cancel', '2 - Delete', '3 - Active & Cancel', '4 - Suspended'], blank: true },
  { key: 'dataStatus', label: 'Data Status', items: ['A - Active', 'D - Delete'], blank: true }, { key: 'expired', label: 'Expired', items: ['N - No', 'Y - Yes'], blank: true },
  { key: 'renewal', label: 'Renewal', items: ['N - New', 'P - Partial', 'R - Renew', 'T - Timely'], blank: true },
  { key: 'premMode', label: 'Premium Mode', items: ['A - Annually', 'Q - Quarterly', 'S - Semi Annual', 'M - Monthly'], blank: true },
  { key: 'adultChild', label: 'Adult / Child', items: ['A - Adult', 'C - Child'], blank: true }, { key: 'sex', label: 'Sex', items: ['M - Male', 'F - Female'], blank: true },
  { key: 'race', label: 'Race', items: ['C - Chinese', 'I - Indian', 'M - Malay', 'O - Other'], blank: true },
  { key: 'rel', label: 'Relation', items: ['D - Daughter', 'H - Husband', 'S - Son', 'W - Wife'], blank: true },
  { key: 'takeover', label: 'Takeover', items: ['N - None', 'Y - Take Over', 'C - Conversion'], blank: true },
  { key: 'label', label: 'Label Status', items: ['Y - Yes'], blank: true },
]
const TEXTS: { key: string; label: string }[] = [{ key: 'mbmNo', label: 'Member No' }, { key: 'policy', label: 'Policy No' }, { key: 'agent', label: 'Agent Code' }]
const code = (v: string) => (v.includes(' - ') ? v.slice(0, v.indexOf(' - ')) : v).trim()

/** Membership > MBM Tracking (FrmNewMBMTracking) and its variants (route data variant: jupiter = FrmNewMBMTracking2). */
@Component({
  selector: 'app-member-tracking-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="v.title" [subtitle]="'Replaces ' + v.form">
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
          @if (type() === 18) {
            <app-field label="Head Count From"><app-date-input name="hcFrom" ariaLabel="Head Count From" [(ngModel)]="f.hcFrom" /></app-field>
            <app-field label="Head Count To"><app-date-input name="hcTo" ariaLabel="Head Count To" [(ngModel)]="f.hcTo" /></app-field>
          }
          <app-field label="Member Type">
            <select name="memberType" [(ngModel)]="memberType"><option value="B">Both</option><option value="P">P - Principal</option><option value="S">S - Supplementary</option></select>
          </app-field>
          <app-field label="Billing">
            <select name="billing" [(ngModel)]="f['billing']"><option value=""></option><option>ASO</option><option>INSURER</option></select>
          </app-field>
          <app-field label="Principal Type">
            <select name="principalType" [(ngModel)]="principalType"><option>Both</option><option>C - Combined Limit</option><option>S - Separated Limit</option></select>
          </app-field>
          <app-action-bar label="Tracking actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print to File</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <fieldset aria-label="Member">
          <legend>Member</legend>
          @for (t of texts; track t.key) {
            <app-field [label]="t.label"><input [name]="t.key" [attr.name]="t.key" maxlength="60" [(ngModel)]="f[t.key]" /></app-field>
            <label class="check-field"><input type="checkbox" [name]="t.key + 'Blank'" [attr.name]="t.key + 'Blank'" [(ngModel)]="f[t.key + 'Blank']" /> Empty</label>
          }
          @for (c of combos; track c.key) {
            <app-field [label]="c.label">
              <select [name]="c.key" [attr.name]="c.key" [(ngModel)]="f[c.key]"><option value=""></option>@for (i of c.items; track i) { <option [value]="i">{{ i }}</option> }</select>
            </app-field>
            @if (c.blank) { <label class="check-field"><input type="checkbox" [name]="c.key + 'Blank'" [attr.name]="c.key + 'Blank'" [(ngModel)]="f[c.key + 'Blank']" /> Empty</label> }
          }
          <app-field label="Insured Type">
            <select name="ins" [(ngModel)]="f['ins']"><option value=""></option>@for (i of lists()['ins'] ?? []; track i) { <option [value]="i">{{ i }}</option> }</select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="insBlank" [(ngModel)]="f['insBlank']" /> Empty</label>
          <app-field label="State">
            <select name="state" [(ngModel)]="f['state']"><option value=""></option>@for (i of lists()['states'] ?? []; track i) { <option [value]="i">{{ i }}</option> }</select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="stateBlank" [(ngModel)]="f['stateBlank']" /> Empty</label>
          <label class="check-field"><input type="checkbox" name="memberTypeBlank" [(ngModel)]="f['memberTypeBlank']" /> Member type empty</label>
        </fieldset>
        <fieldset aria-label="Payor">
          <legend>Payor and company</legend>
          <app-field label="Payor" [wide]="true">
            <input name="payor" list="mt-payors" (focus)="payors()" (change)="companies()" [(ngModel)]="f['payor']" />
            <datalist id="mt-payors">@for (c of lists()['payors'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="payorBlank" [(ngModel)]="f['payorBlank']" /> Empty</label>
          <app-field label="Group Company" [wide]="true">
            <input name="company" list="mt-companies" (change)="departments()" [(ngModel)]="f['company']" />
            <datalist id="mt-companies">@for (c of lists()['companies'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="companyBlank" [(ngModel)]="f['companyBlank']" /> Empty</label>
          <app-field label="Department">
            <select name="department" [(ngModel)]="f['department']"><option value=""></option>@for (c of lists()['departments'] ?? []; track $index) { <option [value]="c">{{ c }}</option> }</select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="departmentBlank" [(ngModel)]="f['departmentBlank']" /> Empty</label>
        </fieldset>
        <fieldset aria-label="Ranges">
          <legend>Ranges</legend>
          @for (r of ranges; track r.key) {
            @if (r.date) {
              <app-field [label]="r.label + ' From'"><app-date-input [name]="r.key + 'From'" [attr.name]="r.key + 'From'" [ariaLabel]="r.label + ' From'" [(ngModel)]="f[r.key + 'From']" /></app-field>
              <app-field [label]="r.label + ' To'"><app-date-input [name]="r.key + 'To'" [attr.name]="r.key + 'To'" [ariaLabel]="r.label + ' To'" [(ngModel)]="f[r.key + 'To']" /></app-field>
            } @else {
              <app-field [label]="r.label + ' From'"><input [name]="r.key + 'From'" [attr.name]="r.key + 'From'" maxlength="50" [(ngModel)]="f[r.key + 'From']" /></app-field>
              <app-field [label]="r.label + ' To'"><input [name]="r.key + 'To'" [attr.name]="r.key + 'To'" maxlength="50" [(ngModel)]="f[r.key + 'To']" /></app-field>
            }
            @if (r.blank) { <label class="check-field"><input type="checkbox" [name]="r.key + 'Blank'" [attr.name]="r.key + 'Blank'" [(ngModel)]="f[r.key + 'Blank']" /> Empty</label> }
          }
        </fieldset>
        <app-data-grid caption="Member tracking" [columns]="columns()" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">
          Principals: <strong>{{ totals().principals }}</strong> · Supplementaries: <strong>{{ totals().supps }}</strong> · Total:
          <strong>{{ totals().principals + totals().supps }}</strong> · MCO: <strong>{{ money(totals().mco) }}</strong> · Premium: <strong>{{ money(totals().premium) }}</strong>
          · Access: <strong>{{ money(totals().access) }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class MemberTrackingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  protected readonly v = VARIANTS[inject(ActivatedRoute).snapshot.data['variant'] ?? 'main']

  protected readonly types = this.v.types
  protected readonly ranges = RANGES
  protected readonly combos = COMBOS
  protected readonly texts = TEXTS
  protected readonly money = (v: number) => formatMoney(v)
  protected readonly reports = signal<{ type: number; selection: string; migrated: boolean }[]>([])
  protected readonly type = signal(-1)
  protected readonly selections = computed(() => this.reports().filter((r) => r.type === this.type()))
  protected selection = ''
  protected memberType = 'P'
  protected principalType = 'Both'
  protected readonly lists = signal<Record<string, string[]>>({})
  protected readonly headers = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly totals = signal({ principals: 0, supps: 0, mco: 0, premium: 0, access: 0 })
  protected readonly busy = signal(false)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected f: any = {}
  protected readonly key = (x: Row) => `${this.rows().indexOf(x)}`
  protected readonly columns = computed<GridColumn<Row>[]>(() => this.headers().map((h, i) => ({ key: `c${i}`, label: h || ' ' })))

  async ngOnInit(): Promise<void> {
    try {
      this.reports.set(await firstValueFrom(this.http.get<{ type: number; selection: string; migrated: boolean }[]>(`${this.v.base}/reports`)))
      await Promise.all([this.load('ins'), this.load('states')])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private async load(kind: string, key?: string): Promise<void> {
    const params = key ? new HttpParams().set('key', key) : undefined
    const items = await firstValueFrom(this.http.get<string[]>(`${this.v.base}/lookup/${kind}`, { params }))
    this.lists.update((l) => ({ ...l, [kind]: items }))
  }

  protected setType(t: number): void {
    this.type.set(Number(t))
    this.selection = this.selections()[0]?.selection ?? ''
    this.rows.set([])
  }

  /** LoadPayor: the payors of the health type. */
  protected payors(): void {
    void this.load('payors', code(this.f['hlt'] ?? '') || 'S').catch((e) => this.notify.notify(errorMessage(e), 'error'))
  }

  /** ChkGrpCom2: the group companies of the payor. */
  protected companies(): void {
    if (this.f['payor']) void this.load('companies', code(this.f['payor'])).catch((e) => this.notify.notify(errorMessage(e), 'error'))
  }

  /** ChkDepartment2: the departments of the group company. */
  protected departments(): void {
    if (this.f['company']) void this.load('departments', this.f['company']).catch((e) => this.notify.notify(errorMessage(e), 'error'))
  }

  protected async search(): Promise<void> {
    if (this.type() < 0) return this.notify.notify('Please select a report type.', 'info')
    // The choice list may have arrived after the type was picked: the shown (first) choice is the one searched.
    if (!this.selections().some((r) => r.selection === this.selection)) this.selection = this.selections()[0]?.selection ?? ''
    const filters: Record<string, string> = {}
    for (const [k, v] of Object.entries(this.f)) {
      if (v === '' || v == null || v === false) continue
      const combo = [...COMBOS.map((c) => c.key), 'ins', 'payor'].includes(k)
      filters[k] = typeof v === 'boolean' ? 'true' : combo ? code(String(v)) : String(v)
    }
    if (this.principalType !== 'Both') filters['detailPrinType'] = code(this.principalType)
    if (this.type() === 7 && this.selection === 'General Listing (Prin. Name)')
      filters['prinType'] = `${this.memberType === 'B' ? 'Both' : this.memberType}|${this.principalType === 'Both' ? 'Both' : code(this.principalType)}`
    this.busy.set(true)
    this.rows.set([])
    try {
      const r = await firstValueFrom(this.http.post<{ headers: string[]; rows: string[][]; principals: number; supplementaries: number; mco: number; premium: number;
        access: number }>(`${this.v.base}/search`, { reportType: this.type(), selection: this.selection, memberType: this.memberType, filters }))
      this.headers.set(r.headers)
      this.rows.set(r.rows.map((cells) => Object.fromEntries(cells.map((c, i) => [`c${i}`, c]))))
      this.totals.set({ principals: r.principals, supps: r.supplementaries, mco: r.mco, premium: r.premium, access: r.access })
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.headers(), ...this.rows().map((x) => this.headers().map((_, i) => x[`c${i}`] ?? ''))], 'MBMTracking.csv')
  }

  protected clear(): void {
    this.f = {}
    this.rows.set([])
    this.totals.set({ principals: 0, supps: 0, mco: 0, premium: 0, access: 0 })
  }
}
