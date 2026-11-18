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
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

type Row = Record<string, string>
/** The two forms: FrmClaimRpt2 and its Jupiter copy. */
const VARIANTS: Record<string, { base: string; title: string; form: string }> = {
  main: { base: '/api/tracking/claim2', title: 'Claim Tracking 2', form: 'FrmClaimRpt2' },
  jupiter: { base: '/api/tracking/claim2-jupiter', title: 'Claim Tracking 2 (Jupiter)', form: 'FrmNewClaimRpt2' },
}
const RANGES: { key: string; label: string; date?: boolean; blank?: boolean }[] = [
  { key: 'dov', label: 'DOV', date: true, blank: true }, { key: 'dor', label: 'DOR', date: true, blank: true }, { key: 'dow', label: 'Claim Open Date', date: true, blank: true },
  { key: 'bordxDate', label: 'Bordx Date', date: true, blank: true }, { key: 'eff', label: 'Payor Eff. Date', date: true, blank: true },
  { key: 'exp', label: 'Payor Exp. Date', date: true, blank: true }, { key: 'joined', label: 'Date Joined', date: true, blank: true }, { key: 'dob', label: 'DOB', date: true, blank: true },
  { key: 'caseNo', label: 'Case No' }, { key: 'mbmNo', label: 'Member No' }, { key: 'bordx', label: 'Bordx No', blank: true }, { key: 'invoice', label: 'Invoice No', blank: true },
  { key: 'submission', label: 'Submission No', blank: true }, { key: 'plan', label: 'Plan Code', blank: true }, { key: 'treatment', label: 'Treatment Reason', blank: true },
  { key: 'polYear', label: 'Policy Year', blank: true }, { key: 'mcDays', label: 'MC Days', blank: true }, { key: 'ic', label: 'Patient IC No', blank: true },
  { key: 'mbmIc', label: 'Member IC No', blank: true }, { key: 'actual', label: 'Actual Amount' }, { key: 'approved', label: 'Approved Amount' },
]
const COMBOS: { key: string; label: string; items: string[] }[] = [
  { key: 'hlt', label: 'Health Type', items: ['S - Sihat', 'N - Non-Sihat'] },
  { key: 'ins', label: 'Insured Type', items: ['I - Individual', 'F - Family', 'G - Group Individual', 'H - Group Family'] },
  { key: 'caseStatus', label: 'Case Status', items: ['O - Open', 'D - Delete', 'A - All'] },
  { key: 'claimability', label: 'Claimability', items: ['A - Accepted', 'R - Rejected'] },
  { key: 'race', label: 'Race', items: ['C - Chinese', 'I - Indian', 'M - Malay', 'O - Other'] },
  { key: 'sex', label: 'Sex', items: ['M - Male', 'F - Female'] },
  { key: 'relation', label: 'Relation', items: ['D - Daughter', 'H - Husband', 'S - Son', 'W - Wife'] },
  { key: 'dataStatus', label: 'Data Status', items: ['A - Active', 'D - Delete'] },
  { key: 'expired', label: 'Expired', items: ['N - No', 'Y - Yes'] },
  { key: 'premMode', label: 'Premium Mode', items: ['A - Annually', 'Q - Quarterly', 'S - Semi Annual', 'M - Monthly'] },
  { key: 'renewal', label: 'Renewal', items: ['N - New', 'P - Partial', 'R - ReNew', 'T - Timely'] },
  { key: 'takeover', label: 'Takeover', items: ['N - None', 'Y - Take Over', 'C - Conversion'] },
]
const TEXTS = [{ key: 'agent', label: 'Agent Code', blank: true }, { key: 'policy', label: 'Policy No', blank: true }, { key: 'patient', label: 'Patient Name' },
  { key: 'branch', label: 'Branch', blank: true }]
/** Sent as their code (the text before " - "); the clinic is sent as picked (the desktop compared HOSName with the whole "code - name" text). */
const CODED = [...COMBOS.map((c) => c.key), 'polStatus', 'payor']
const code = (v: string) => (v.includes(' - ') ? v.slice(0, v.indexOf(' - ')) : v).trim()

/** Claim / Bordx > Claim Tracking 2 (FrmClaimRpt2) and its Jupiter copy (route data variant: jupiter). */
@Component({
  selector: 'app-claim-tracking2-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="v.title" [subtitle]="'Replaces ' + v.form">
      <div class="screen-form">
        <fieldset aria-label="Report">
          <legend>Report</legend>
          <div role="radiogroup" aria-label="Amount">
            <label class="check-field"><input type="radio" name="basis" [value]="1" [ngModel]="basis()" (ngModelChange)="basis.set($event)" /> Actual</label>
            <label class="check-field"><input type="radio" name="basis" [value]="2" [ngModel]="basis()" (ngModelChange)="basis.set($event)" /> Approved</label>
            <label class="check-field"><input type="radio" name="basis" [value]="3" [ngModel]="basis()" (ngModelChange)="basis.set($event)" /> Cost Detail</label>
          </div>
          @if (basis() !== 3) {
            <app-field label="Report" [span]="2">
              <select name="report" [(ngModel)]="report">@for (r of choices(); track r.report) { <option [ngValue]="r.report">{{ r.name }}</option> }</select>
            </app-field>
          }
          <app-field label="Member Type">
            <select name="memType" [(ngModel)]="memType"><option>Both</option><option>P - Principal</option><option>S - Supplementary</option></select>
          </app-field>
          <app-field label="Supp. Type">
            <select name="suppType" [(ngModel)]="suppType"><option>Both</option><option>C - Combined Limit</option><option>S - Separated Limit</option></select>
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
            <label class="check-field"><input type="checkbox" [name]="c.key + 'Blank'" [attr.name]="c.key + 'Blank'" [(ngModel)]="f[c.key + 'Blank']" /> Empty</label>
          }
          <app-field label="Policy Status">
            <select name="polStatus" [(ngModel)]="f['polStatus']"><option value=""></option>@for (i of lists()['polstatus'] ?? []; track i) { <option [value]="i">{{ i }}</option> }</select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="polStatusBlank" [(ngModel)]="f['polStatusBlank']" /> Empty</label>
          <app-field label="State">
            <select name="state" [(ngModel)]="f['state']"><option value=""></option>@for (i of lists()['states'] ?? []; track i) { <option [value]="i">{{ i }}</option> }</select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="stateBlank" [(ngModel)]="f['stateBlank']" /> Empty</label>
        </fieldset>
        <fieldset aria-label="Payor">
          <legend>Payor, company and clinic</legend>
          <app-field label="Payor" [span]="2">
            <input name="payor" list="ct2-payors" (focus)="payors()" (change)="companies()" [(ngModel)]="f['payor']" />
            <datalist id="ct2-payors">@for (c of lists()['payors'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="payorBlank" [(ngModel)]="f['payorBlank']" /> Empty</label>
          <app-field label="Group Company" [span]="2">
            <input name="company" list="ct2-companies" [(ngModel)]="f['company']" />
            <datalist id="ct2-companies">@for (c of lists()['companies'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="companyBlank" [(ngModel)]="f['companyBlank']" /> Empty</label>
          <app-field label="Clinic" [span]="2">
            <input name="clinicName" list="ct2-clinics" (focus)="load('clinics')" [(ngModel)]="f['clinicName']" />
            <datalist id="ct2-clinics">@for (c of lists()['clinics'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="clinicNameBlank" [(ngModel)]="f['clinicNameBlank']" /> Empty</label>
          <app-field label="Clinic Group" [span]="2">
            <input name="clinicGroup" list="ct2-groups" (focus)="load('clinicgroups')" [(ngModel)]="f['clinicGroup']" />
            <datalist id="ct2-groups">@for (c of lists()['clinicgroups'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="clinicGroupBlank" [(ngModel)]="f['clinicGroupBlank']" /> Empty</label>
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
        <app-data-grid caption="Claim tracking 2" [columns]="columns()" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">Records: <strong>{{ records() }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class ClaimTracking2Page implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  protected readonly v = VARIANTS[inject(ActivatedRoute).snapshot.data['variant'] ?? 'main']

  protected readonly ranges = RANGES
  protected readonly combos = COMBOS
  protected readonly texts = TEXTS
  protected readonly reports = signal<{ basis: number; report: number; name: string }[]>([])
  protected readonly basis = signal(1)
  protected readonly choices = computed(() => this.reports().filter((r) => r.basis === this.basis()))
  protected report = 4
  protected memType = 'Both'
  protected suppType = 'Both'
  protected readonly lists = signal<Record<string, string[]>>({})
  protected readonly headers = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly records = signal(0)
  protected readonly busy = signal(false)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected f: any = {}
  protected readonly key = (x: Row) => `${this.rows().indexOf(x)}`
  protected readonly columns = computed<GridColumn<Row>[]>(() => this.headers().map((h, i) => ({ key: `c${i}`, label: h || ' ' })))

  async ngOnInit(): Promise<void> {
    try {
      this.reports.set(await firstValueFrom(this.http.get<{ basis: number; report: number; name: string }[]>(`${this.v.base}/reports`)))
      await Promise.all([this.load('polstatus'), this.load('states')])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async load(kind: string, key?: string): Promise<void> {
    if (!key && this.lists()[kind]) return
    const params = key ? new HttpParams().set('key', key) : undefined
    try {
      const items = await firstValueFrom(this.http.get<string[]>(`${this.v.base}/lookup/${kind}`, { params }))
      this.lists.update((l) => ({ ...l, [kind]: items }))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected payors(): void {
    void this.load('payors', code(this.f['hlt'] ?? '') || 'S')
  }

  protected companies(): void {
    if (this.f['payor']) void this.load('companies', code(this.f['payor']))
  }

  protected async search(): Promise<void> {
    const filters: Record<string, string> = {}
    for (const [k, v] of Object.entries(this.f)) {
      if (v === '' || v == null || v === false) continue
      filters[k] = typeof v === 'boolean' ? 'true' : CODED.includes(k) ? code(String(v)) : String(v)
    }
    const mem = this.memType === 'Both' ? 'Both' : code(this.memType)
    const supp = this.suppType === 'Both' ? 'Both' : code(this.suppType)
    if (!(mem === 'Both' && supp === 'Both')) filters['memType'] = `${mem}|${supp}`
    this.busy.set(true)
    this.rows.set([])
    try {
      const r = await firstValueFrom(this.http.post<{ headers: string[]; rows: string[][]; records: number }>(`${this.v.base}/search`,
        { basis: this.basis(), report: this.basis() === 3 ? 0 : this.report, filters }))
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
    saveCsv([this.headers(), ...this.rows().map((x) => this.headers().map((_, i) => x[`c${i}`] ?? ''))], 'ClaimTracking2.csv')
  }

  protected clear(): void {
    this.f = {}
    this.rows.set([])
    this.records.set(0)
  }
}
