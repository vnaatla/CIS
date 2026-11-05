import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { LookupsService } from '../../core/lookups.service'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/lg-audit/audit-tracking'
type Row = Record<string, unknown>
interface TrackingRow {
  first: boolean
  row: Row
}
type Filters = Record<string, string | boolean>

/** Text ranges of the search frame: [label, key prefix, has "blank" check box]. */
const RANGES: [string, string, boolean][] = [
  ['Case No', 'case', true], ['Member No', 'member', true], ['IC No', 'ic', true], ['Invoice No', 'invoice', true], ['Bordx No', 'bordx', true],
  ['Plan', 'plan', true], ['OP Batch No', 'opBatch', true], ['OMNI Ref No', 'omniRef', true], ['Invoice Amount', 'amount', true],
]
/** Date ranges: [label, key prefix, has "blank" check box]. */
const DATES: [string, string, boolean][] = [
  ['Visit Date', 'visit', true], ['Register Date', 'register', true], ['Invoice Date', 'invoiceDate', true], ['Invoice Reg. Date', 'entered', true],
  ['Bordx Date', 'bordxDate', false],
]
const MONTHS = Array.from({ length: 12 }, (_, i) => String(i + 1))
const YEARS = Array.from({ length: 3000 - 1999 + 1 }, (_, i) => String(1999 + i))
const money = (v: unknown) => (v == null || v === '' ? '' : formatMoney(Number(v)))
const date = (v: unknown) => (!v || String(v).startsWith('1900') ? '' : formatDate(String(v)))

/** LG & Audit > Audit Tracking (FrmAuditInvoiceTracking): audited cases by any of the desktop's filters, one row per medication. */
@Component({
  selector: 'app-audit-tracking-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Audit Tracking" subtitle="Replaces FrmAuditInvoiceTracking">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          @for (r of ranges; track r[1]) {
            <app-field [label]="r[0] + ' From'"><input [attr.name]="r[1] + 'From'" [(ngModel)]="f[r[1] + 'From']" /></app-field>
            <app-field [label]="r[0] + ' To'"><input [attr.name]="r[1] + 'To'" [(ngModel)]="f[r[1] + 'To']" /></app-field>
            <label class="check-field"><input type="checkbox" [attr.name]="r[1] + 'Blank'" [(ngModel)]="f[r[1] + 'Blank']" /> {{ r[0] }} blank</label>
          }
          @for (r of dates; track r[1]) {
            <app-field [label]="r[0] + ' From'"><app-date-input [attr.name]="r[1] + 'From'" [ariaLabel]="r[0] + ' From'" [(ngModel)]="f[r[1] + 'From']" /></app-field>
            <app-field [label]="r[0] + ' To'"><app-date-input [attr.name]="r[1] + 'To'" [ariaLabel]="r[0] + ' To'" [(ngModel)]="f[r[1] + 'To']" /></app-field>
            @if (r[2]) {
              <label class="check-field"><input type="checkbox" [attr.name]="r[1] + 'Blank'" [(ngModel)]="f[r[1] + 'Blank']" /> {{ r[0] }} blank</label>
            }
          }
          <app-field label="Pay To">
            <select name="payTo" [(ngModel)]="f['payTo']"><option value=""></option><option value="C">C - Clinic</option><option value="M">M - Member</option></select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="payToBlank" [(ngModel)]="f['payToBlank']" /> Pay To blank</label>
          <app-field label="Clinic Code"><input name="clinic" maxlength="20" [(ngModel)]="f['clinic']" /></app-field>
          <label class="check-field"><input type="checkbox" name="clinicBlank" [(ngModel)]="f['clinicBlank']" /> Clinic blank</label>
          <app-field label="Payor">
            <select name="payor" [(ngModel)]="f['payor']">
              <option value=""></option>
              @for (p of payors(); track p.code) { <option [value]="p.code">{{ p.code }} - {{ p.name }}</option> }
            </select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="payorBlank" [(ngModel)]="f['payorBlank']" /> Payor blank</label>
          <app-field label="Group Company">
            <select name="groupCompany" [(ngModel)]="f['groupCompany']" (ngModelChange)="groupChanged($event)">
              <option value=""></option>
              @for (g of groups(); track g.code) { <option [value]="g.name">{{ g.name }}</option> }
            </select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="groupBlank" [(ngModel)]="f['groupBlank']" /> Group Company blank</label>
          <app-field label="Subsidiary">
            <select name="subsidiary" [(ngModel)]="f['subsidiary']"><option value=""></option>@for (x of subsidiaries(); track x) { <option [value]="x">{{ x }}</option> }</select>
          </app-field>
          <app-field label="Department">
            <select name="department" [(ngModel)]="f['department']"><option value=""></option>@for (x of departments(); track x) { <option [value]="x">{{ x }}</option> }</select>
          </app-field>
          <app-field label="Patient Name"><input name="patientName" maxlength="200" [(ngModel)]="f['patientName']" /></app-field>
          <label class="check-field"><input type="checkbox" name="nameBlank" [(ngModel)]="f['nameBlank']" /> Name blank</label>
          <app-field label="Case Status">
            <select name="caseStatus" [(ngModel)]="f['caseStatus']">
              <option value=""></option><option value="O">O - Open</option><option value="C">C - Close</option><option value="D">D - Delete</option>
            </select>
          </app-field>
          <app-field label="Claimability">
            <select name="claimability" [(ngModel)]="f['claimability']"><option value=""></option><option value="A">A - Accepted</option><option value="R">R - Rejected</option></select>
          </app-field>
          <app-field label="A/C Period Month From"><select name="monthFrom" [(ngModel)]="f['monthFrom']"><option value=""></option>@for (m of months; track m) { <option [value]="m">{{ m }}</option> }</select></app-field>
          <app-field label="A/C Period Month To"><select name="monthTo" [(ngModel)]="f['monthTo']"><option value=""></option>@for (m of months; track m) { <option [value]="m">{{ m }}</option> }</select></app-field>
          <label class="check-field"><input type="checkbox" name="monthBlank" [(ngModel)]="f['monthBlank']" /> Month blank</label>
          <app-field label="A/C Period Year From"><select name="yearFrom" [(ngModel)]="f['yearFrom']"><option value=""></option>@for (y of years; track y) { <option [value]="y">{{ y }}</option> }</select></app-field>
          <app-field label="A/C Period Year To"><select name="yearTo" [(ngModel)]="f['yearTo']"><option value=""></option>@for (y of years; track y) { <option [value]="y">{{ y }}</option> }</select></app-field>
          <label class="check-field"><input type="checkbox" name="yearBlank" [(ngModel)]="f['yearBlank']" /> Year blank</label>
          <app-field label="Clinic Type">
            <select name="clinicType" [(ngModel)]="f['clinicType']">
              <option value=""></option><option value="0">0 - Specialised</option><option value="1">1 - General Practisional</option><option value="2">2 - Others</option>
            </select>
          </app-field>
          <app-field label="FFS Clinic">
            <select name="ffs" [(ngModel)]="f['ffs']"><option value=""></option><option value="Y">Y - Yes</option><option value="N">N - No</option></select>
          </app-field>
          <app-field label="OMNI Sent">
            <select name="omniSent" [(ngModel)]="f['omniSent']"><option value=""></option><option value="Y">Y - Yes</option><option value="N">N - No</option></select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="omni" [(ngModel)]="f['omni']" /> OMNI Status: Y - Yes</label>
          <label class="check-field"><input type="checkbox" name="tims" [(ngModel)]="f['tims']" /> TIMS: Y - Yes</label>
          <label class="check-field"><input type="checkbox" name="audited" [(ngModel)]="f['audited']" /> Audit: Y - Yes</label>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print to File</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Audit tracking" [columns]="columns" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">Records: <strong>{{ rows().length }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class AuditTrackingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  private readonly lookups = inject(LookupsService)

  protected readonly ranges = RANGES
  protected readonly dates = DATES
  protected readonly months = MONTHS
  protected readonly years = YEARS
  protected readonly payors = signal<LookupOption[]>([])
  protected readonly groups = signal<LookupOption[]>([])
  protected readonly subsidiaries = signal<string[]>([])
  protected readonly departments = signal<string[]>([])
  protected readonly rows = signal<(Row & { _i: number })[]>([])
  protected readonly busy = signal(false)
  protected readonly key = (r: Row) => String(r['_i'])
  protected f: Filters = {}
  private searched: Filters | null = null

  protected readonly columns: GridColumn<Row>[] = [
    { key: 'MBMIcBcPp', label: 'IC No' }, { key: 'MBMName', label: 'Name' }, { key: 'CASNumber', label: 'Case No' },
    { key: 'DateVisit', label: 'Visit Date', value: (r) => date(r['DateVisit']) }, { key: 'MedInjType', label: 'Medication' },
    { key: 'MedInjDays', label: 'Days' }, { key: 'MedInjQty', label: 'Qty' },
    { key: 'MedCharges', label: 'Charges', align: 'right', value: (r) => money(r['MedCharges']) },
    { key: 'SPFee', label: 'Consultation Fee', align: 'right', value: (r) => money(r['SPFee']) },
    { key: 'INVAmount', label: 'Invoice Amount', align: 'right', value: (r) => money(r['INVAmount']) },
    { key: 'CasRemark', label: 'Remark' }, { key: 'AuditQuery', label: 'Query' }, { key: 'AuditAnswer', label: 'Answer' },
    { key: 'AuditRemarks', label: 'Audit Remarks' },
    { key: 'AuditDiscountReceived', label: 'Discount Received', align: 'right', value: (r) => money(r['AuditDiscountReceived']) },
    { key: 'ReasonDesc1', label: 'Diagnosis 1' }, { key: 'ReasonDesc2', label: 'Diagnosis 2' }, { key: 'ReasonDesc3', label: 'Diagnosis 3' },
    { key: 'ReasonDesc4', label: 'Diagnosis 4' }, { key: 'Diagnosis5', label: 'Diagnosis 5' }, { key: 'AuditBy', label: 'Audit By' },
    { key: 'AuditDate', label: 'Audit Date', value: (r) => date(r['AuditDate']) }, { key: 'CLNCode', label: 'Clinic Code' },
    { key: 'CLNName', label: 'Clinic Name' }, { key: 'MBMGroupCompany', label: 'Group Company' },
  ]

  async ngOnInit(): Promise<void> {
    try {
      const [payors, groups] = await Promise.all([firstValueFrom(this.lookups.payors()), firstValueFrom(this.lookups.groupCompanies())])
      this.payors.set(payors)
      this.groups.set(groups)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CboGrpCom_LostFocus: the group company's subsidiaries and departments. */
  protected async groupChanged(group: string): Promise<void> {
    this.f['subsidiary'] = ''
    this.f['department'] = ''
    try {
      const r = await firstValueFrom(this.http.get<{ subsidiaries: string[]; departments: string[] }>(`${BASE}/lookups`, { params: { groupCompany: group ?? '' } }))
      this.subsidiaries.set(r.subsidiaries)
      this.departments.set(r.departments)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.searched = null
    try {
      const filter = { ...this.f }
      const list = await firstValueFrom(this.http.post<TrackingRow[]>(`${BASE}/search`, filter))
      this.rows.set(list.map((x, i) => ({ ...x.row, _i: i })))
      this.searched = filter
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** CmdPrintFile_Click. */
  protected async print(): Promise<void> {
    if (!this.searched && this.rows().length === 0) return this.notify.notify('Please search again.!', 'error')
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/print`, this.searched ?? this.f, { responseType: 'blob' })), 'InvoiceAuditTracking.xlsx')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  protected clear(): void {
    this.f = {}
    this.rows.set([])
    this.searched = null
    this.subsidiaries.set([])
    this.departments.set([])
  }
}
