import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { ActivatedRoute, Router } from '@angular/router'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const money = (k: string, label: string): GridColumn<Row> => ({ key: k, label, align: 'right', value: (x) => formatMoney(Number(x[k] ?? 0)) })
const date = (k: string, label: string): GridColumn<Row> => ({ key: k, label, value: (x) => (x[k] ? formatDate(String(x[k])) : '') })
const t = (k: string, label: string): GridColumn<Row> => ({ key: k, label })
const COLUMNS: GridColumn<Row>[] = [
  t('CASNumber', 'Case No'), t('BordxRefNo', 'Bordx No'), date('BordxDate', 'Bordx Date'), t('BTBGBordxRefNo', 'BTBG Bordx'), t('BillNo', 'Invoice No'),
  date('BillDate', 'Invoice Date'), money('INVAmount', 'Inv Amount'), money('INVPaidByMem', 'Paid by Mem'), money('NetAmt', 'Net Amount'),
  money('CLMCoPayCLNColl', 'Co-Pay'), money('CLMBordxAmt', 'Bordx Amount'), money('CLMBTBGAmt', 'BTBG Amount'), money('ClmNonGPFee', 'Non-GP Fee'),
  money('CLMMCOProcFee', 'MCO Fee'), date('RegisterDate', 'Register Date'), date('DateVisit', 'Visit Date'), t('DaysText', 'Days'), t('INVPayee', 'Payee'),
  t('ClinicText', 'Clinic'), t('CLNType', 'Clinic Type'), t('MBMNumber', 'Member No'), t('PatientName', 'Patient'), t('PLNCode', 'Plan'), t('INSCode', 'Insurer'),
  t('MBMGroupCompany', 'Group Company'), date('INVEnteredDate', 'Entered'), t('INVEnteredBy', 'Entered By'), t('OPBatchNo', 'OP Batch'), t('CASStatus', 'Status'),
  t('Claimability', 'Claimability'), t('ClaimBordxStatus', 'Bordx Status'),
  { key: 'acc', label: 'A/C Period', value: (x) => (s(x['AccPeriodMonth']) || s(x['AccPeriodYear']) ? `${s(x['AccPeriodMonth'])}/${s(x['AccPeriodYear'])}` : '') },
  t('MCDays', 'MC Days'), t('ReasonDesc1', 'Reason 1'), t('ReasonDesc2', 'Reason 2'), t('ReasonDesc3', 'Reason 3'), t('ReasonDesc4', 'Reason 4'),
  t('Reason5Text', 'Reason 5'), t('OPNoEnterUser', 'OP Entered By'), t('CLNFFSStatus', 'FFS'), t('Subsidiary', 'Subsidiary'), t('ClaimLastUpdateUser', 'Updated By'),
  date('ClaimLastUpdateDate', 'Updated'), t('PVCLNNo', 'PV No'), date('PVCLNDate', 'PV Date'), t('PVCLNChequeNo', 'Cheque No'), date('PVCLNChequeDate', 'Cheque Date'),
  t('OMNIRefNo', 'OMNI Ref'), ...[1, 2, 3, 4, 5, 6, 7].map((i) => t(`MedInjType${i}`, `Med/Inj ${i}`)), ...[1, 2, 3, 4, 5].map((i) => t(`MedInjDays${i}`, `Days ${i}`)),
  ...[1, 2, 3, 4, 5].map((i) => t(`MedInjQty${i}`, `Qty ${i}`)), money('FeeTotal', 'Fee'), money('MedTotal', 'Medicine'), money('InjTotal', 'Injection'),
  t('CLMOMNISent', 'OMNI Sent'), t('CASMgmtNotes', 'Mgmt Notes'), t('CasRemark', 'Remark'), t('DoctorName', 'Doctor'), t('IcNo', 'IC No'), t('coverid', 'Cover ID'),
  date('OPBatchDate', 'OP Batch Date'), t('recordingpic', 'Recording PIC'), t('PreliminaryDiag', 'Preliminary Diag'), t('clmsubsidiary', 'Claim Subsidiary'),
  t('clmcostcenter', 'Cost Centre'), t('CLMDeptName', 'Department'), t('InternalRemarks', 'Internal Remarks'),
]

/** Invoice > Tracking (FrmInvoiceTracking) and Tracking (Jupiter) (FrmInvoiceTracking2, route data jupiter: fewer filters, one case status, MSNO).
 * Double-click opens the case in Invoice Adjustment, as the desktop. */
@Component({
  selector: 'app-invoice-tracking-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="jupiter ? 'Invoice Tracking (Jupiter)' : 'Invoice Tracking'" [subtitle]="jupiter ? 'Replaces FrmInvoiceTracking2' : 'Replaces FrmInvoiceTracking'">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Case and member</legend>
@if (!jupiter) {
          <app-field label="Entity">
            <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Health Connect</option><option [ngValue]="1">Sihat Johor</option></select>
          </app-field>
          }
          <app-field label="Case No From"><input name="caseFrom" maxlength="20" [(ngModel)]="f.caseFrom" /></app-field>
          <app-field label="Case No To"><input name="caseTo" maxlength="20" [(ngModel)]="f.caseTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="caseBlank" [(ngModel)]="f.caseBlank" /> Empty</label>
          <app-field label="Member No From"><input name="memberFrom" maxlength="30" [(ngModel)]="f.memberFrom" /></app-field>
          <app-field label="Member No To"><input name="memberTo" maxlength="30" [(ngModel)]="f.memberTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="memberBlank" [(ngModel)]="f.memberBlank" /> Empty</label>
          <app-field label="Patient Name"><input name="patient" maxlength="200" [(ngModel)]="f.patient" /></app-field>
          <label class="check-field"><input type="checkbox" name="patientBlank" [(ngModel)]="f.patientBlank" /> Empty</label>
          <app-field label="IC No From"><input name="icFrom" maxlength="30" [(ngModel)]="f.icFrom" /></app-field>
          <app-field label="IC No To"><input name="icTo" maxlength="30" [(ngModel)]="f.icTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="icBlank" [(ngModel)]="f.icBlank" /> Empty</label>
          <app-field label="Plan From"><input name="planFrom" maxlength="20" [(ngModel)]="f.planFrom" /></app-field>
          <app-field label="Plan To"><input name="planTo" maxlength="20" [(ngModel)]="f.planTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="planBlank" [(ngModel)]="f.planBlank" /> Empty</label>
@if (!jupiter) {
          <app-field label="Flex OP">
            <select name="flex" [(ngModel)]="f.flex"><option [ngValue]="null"></option><option [ngValue]="0">OutPatient</option><option [ngValue]="1">flexCare</option></select>
          </app-field>
          }
        </fieldset>
        <fieldset aria-label="Payor and clinic">
          <legend>Payor and clinic</legend>
          <app-field label="Payor" [wide]="true">
            <input name="payor" list="it-payors" (focus)="load('payors')" (change)="companiesFor()" [(ngModel)]="f.payor" />
            <datalist id="it-payors">@for (c of lists()['payors'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="payorBlank" [(ngModel)]="f.payorBlank" /> Empty</label>
          <app-field label="Group Company" [wide]="true">
            <input name="groupCompany" list="it-companies" (change)="companyLists()" [(ngModel)]="f.groupCompany" />
            <datalist id="it-companies">@for (c of lists()['companies'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="groupCompanyBlank" [(ngModel)]="f.groupCompanyBlank" /> Empty</label>
@if (!jupiter) {
          <app-field label="Subsidiary">
            <select name="subsidiary" [(ngModel)]="f.subsidiary"><option value=""></option>@for (c of lists()['subsidiaries'] ?? []; track $index) { <option [value]="c">{{ c }}</option> }</select>
          </app-field>
          }
@if (!jupiter) {
          <app-field label="Department">
            <select name="department" [(ngModel)]="f.department"><option value=""></option>@for (c of lists()['departments'] ?? []; track $index) { <option [value]="c">{{ c }}</option> }</select>
          </app-field>
          }
          <app-field label="Clinic" [wide]="true">
            <input name="clinic" list="it-clinics" (focus)="load('clinics')" [(ngModel)]="f.clinic" />
            <datalist id="it-clinics">@for (c of lists()['clinics'] ?? []; track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <label class="check-field"><input type="checkbox" name="payeeBlank" [(ngModel)]="f.payeeBlank" /> Payee empty</label>
          <app-field label="Clinic Type">
            <select name="clinicType" [(ngModel)]="f.clinicType"><option value=""></option><option>0 - Specialised</option><option>1 - General Practisional</option><option>2 - Others</option></select>
          </app-field>
          <app-field label="FFS">
            <select name="ffs" [(ngModel)]="f.ffs"><option value=""></option><option>Y - Yes</option><option>N - No</option></select>
          </app-field>
          <app-field label="Pay To">
            <select name="payTo" [(ngModel)]="f.payTo"><option value=""></option><option>C - Clinic</option><option>M - Member</option></select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="payToBlank" [(ngModel)]="f.payToBlank" /> Empty</label>
@if (!jupiter) {
          <app-field label="Payment Status">
            <select name="paid" [(ngModel)]="f.paid"><option value=""></option><option>Yes</option><option>No</option></select>
          </app-field>
          }
        </fieldset>
        <fieldset aria-label="Invoice and bordereaux">
          <legend>Invoice and bordereaux</legend>
          <app-field label="Invoice No From"><input name="invoiceFrom" maxlength="50" [(ngModel)]="f.invoiceFrom" /></app-field>
          <app-field label="Invoice No To"><input name="invoiceTo" maxlength="50" [(ngModel)]="f.invoiceTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="invoiceBlank" [(ngModel)]="f.invoiceBlank" /> Empty</label>
          <app-field label="Invoice Amount From"><input name="amountFrom" [(ngModel)]="f.amountFrom" /></app-field>
          <app-field label="Invoice Amount To"><input name="amountTo" [(ngModel)]="f.amountTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="amountBlank" [(ngModel)]="f.amountBlank" /> Empty</label>
          <app-field label="Bordx No From"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" /></app-field>
          <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="bordxBlank" [(ngModel)]="f.bordxBlank" /> Empty</label>
          <app-field label="OP Batch From"><input name="opBatchFrom" maxlength="30" [(ngModel)]="f.opBatchFrom" /></app-field>
          <app-field label="OP Batch To"><input name="opBatchTo" maxlength="30" [(ngModel)]="f.opBatchTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="opBatchBlank" [(ngModel)]="f.opBatchBlank" /> Empty</label>
@if (!jupiter) {
          <app-field label="OMNI Ref From"><input name="omniRefFrom" maxlength="30" [(ngModel)]="f.omniRefFrom" /></app-field>
          <app-field label="OMNI Ref To"><input name="omniRefTo" maxlength="30" [(ngModel)]="f.omniRefTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="omniRefBlank" [(ngModel)]="f.omniRefBlank" /> Empty</label>
          <app-field label="OMNI Sent">
            <select name="omniSent" [(ngModel)]="f.omniSent"><option value=""></option><option>Y - Yes</option><option>N - No</option></select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="omni" [(ngModel)]="f.omni" /> OMNI company</label>
          <label class="check-field"><input type="checkbox" name="tims" [(ngModel)]="f.tims" /> TIMS company</label>
          }
          <app-field label="A/C Month From"><input name="accMonthFrom" maxlength="2" [(ngModel)]="f.accMonthFrom" /></app-field>
          <app-field label="A/C Month To"><input name="accMonthTo" maxlength="2" [(ngModel)]="f.accMonthTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="accMonthBlank" [(ngModel)]="f.accMonthBlank" /> Empty</label>
          <app-field label="A/C Year From"><input name="accYearFrom" maxlength="4" [(ngModel)]="f.accYearFrom" /></app-field>
          <app-field label="A/C Year To"><input name="accYearTo" maxlength="4" [(ngModel)]="f.accYearTo" /></app-field>
          <label class="check-field"><input type="checkbox" name="accYearBlank" [(ngModel)]="f.accYearBlank" /> Empty</label>
        </fieldset>
        <fieldset aria-label="Dates and status">
          <legend>Dates and status</legend>
          @for (d of jupiter ? jupiterDates : dates; track d.key) {
            <app-field [label]="d.label + ' From'"><app-date-input [name]="d.key + 'From'" [ariaLabel]="d.label + ' From'" [(ngModel)]="f[d.key + 'From']" /></app-field>
            <app-field [label]="d.label + ' To'"><app-date-input [name]="d.key + 'To'" [ariaLabel]="d.label + ' To'" [(ngModel)]="f[d.key + 'To']" /></app-field>
            @if (d.blank) { <label class="check-field"><input type="checkbox" [name]="d.key + 'Blank'" [(ngModel)]="f[d.key + 'Blank']" /> Empty</label> }
          }
          @if (jupiter) {
            <app-field label="Case Status">
              <select name="status" [(ngModel)]="f.status"><option value=""></option><option>O - Open</option><option>C - Close</option><option>D - Delete</option></select>
            </app-field>
          } @else {
          <span class="radio-row" aria-label="Case status">
            @for (c of statuses; track c.code) {
              <label class="check-field"><input type="checkbox" [checked]="f.statuses.includes(c.code)" (change)="toggle(c.code)" /> {{ c.label }}</label>
            }
            <label class="check-field"><input type="checkbox" name="pending" [(ngModel)]="f.pending" /> Pending</label>
          </span>
          }
          <app-field label="Claimability">
            <select name="claimability" [(ngModel)]="f.claimability"><option value=""></option><option>A - Accepted</option><option>R - Rejected</option></select>
          </app-field>
          <app-action-bar label="Tracking actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print to File</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Invoices" [columns]="columns" [rows]="rows()" [rowKey]="key" (rowActivated)="open($event)" />
        <p class="totals" aria-label="Totals">
          Records: <strong>{{ rows().length }}</strong> · Invoice Amount: <strong>{{ fmt(total()) }}</strong> · Approved: <strong>{{ fmt(approved()) }}</strong>
          · BTBG: <strong>{{ fmt(btbg()) }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class InvoiceTrackingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  private readonly router = inject(Router)
  protected readonly jupiter = inject(ActivatedRoute).snapshot.data['variant'] === 'jupiter'
  private readonly base = this.jupiter ? '/api/tracking/invoice-jupiter' : '/api/tracking/invoice'

  protected readonly fmt = (v: number) => formatMoney(v)
  protected readonly columns = this.jupiter ? [t('MSNO', 'MS No'), ...COLUMNS] : COLUMNS
  protected readonly jupiterDates = [
    { key: 'visit', label: 'Visit Date', blank: true }, { key: 'register', label: 'Register Date', blank: true },
    { key: 'invoiceDate', label: 'Invoice Date', blank: true }, { key: 'entered', label: 'Invoice Reg Date', blank: true },
  ]
  protected readonly dates = [
    { key: 'bordxDate', label: 'Bordx Date', blank: false }, { key: 'visit', label: 'Visit Date', blank: true }, { key: 'register', label: 'Register Date', blank: true },
    { key: 'invoiceDate', label: 'Invoice Date', blank: true }, { key: 'entered', label: 'Invoice Reg Date', blank: true }, { key: 'opBatchDate', label: 'OP Batch Date', blank: false },
  ]
  protected readonly statuses = [{ code: 'O', label: 'Open' }, { code: 'D', label: 'Delete' }, { code: 'C', label: 'Close' }]
  protected readonly lists = signal<Record<string, string[]>>({})
  protected readonly rows = signal<Row[]>([])
  protected readonly total = signal(0)
  protected readonly approved = signal(0)
  protected readonly btbg = signal(0)
  protected readonly busy = signal(false)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected f: any = this.empty()
  protected readonly key = (x: Row) => s(x['CASNumber']) + '|' + this.rows().indexOf(x)

  private empty(): Record<string, unknown> {
    return { entity: 0, flex: null, statuses: [] as string[], pending: false }
  }

  ngOnInit(): void {
    void this.load('payors')
  }

  protected async load(kind: string, key?: string): Promise<void> {
    if (!key && this.lists()[kind]) return
    try {
      const params = key ? new HttpParams().set('key', key) : undefined
      const items = await firstValueFrom(this.http.get<string[]>(`${this.base}/lookup/${kind}`, { params }))
      this.lists.update((l) => ({ ...l, [kind]: items }))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected companiesFor(): void {
    if (s(this.f['payor'])) void this.load('companies', s(this.f['payor']))
  }

  protected companyLists(): void {
    const c = s(this.f['groupCompany'])
    if (!c) return
    void this.load('subsidiaries', c)
    void this.load('departments', c)
  }

  protected toggle(code: string): void {
    const list: string[] = this.f['statuses']
    this.f['statuses'] = list.includes(code) ? list.filter((x) => x !== code) : [...list, code]
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    try {
      const r = await firstValueFrom(this.http.post<{ rows: Row[]; total: number; approved: number; btbg: number }>(`${this.base}/search`, this.f))
      this.rows.set(r.rows)
      this.total.set(r.total)
      this.approved.set(r.approved)
      this.btbg.set(r.btbg)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected open(x: Row): void {
    void this.router.navigate(['/invoice/adjustment'], { queryParams: { case: s(x['CASNumber']) } })
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'InvoiceTracking.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
    this.total.set(0)
    this.approved.set(0)
    this.btbg.set(0)
  }
}
