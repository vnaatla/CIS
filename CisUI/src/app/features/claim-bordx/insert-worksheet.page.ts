import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
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

const BASE = '/api/claim-bordx/insert-worksheet'
type Row = Record<string, unknown>
interface Option {
  code: string
  name: string
}
interface Result {
  rows: Row[]
  total: number
  totalAmount: number
  uob: boolean
}

const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))

/** Claim Bordx > Worksheet > Insert Worksheet (frmWSBordx). */
@Component({
  selector: 'app-insert-worksheet-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Insert Worksheet into Bordereaux" subtitle="Replaces frmWSBordx">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Bordx Type" [required]="true">
            <select name="bordxType" [(ngModel)]="f.bordxType" (ngModelChange)="f.asoOption = null">
              <option [ngValue]="null"></option>
              @for (t of bordxTypes; track $index) { <option [ngValue]="$index">{{ t }}</option> }
            </select>
          </app-field>
          @if (f.bordxType === 1) {
            <div class="radio-row" role="radiogroup" aria-label="ASO type">
              <label class="check-field"><input type="radio" name="aso" [value]="0" [(ngModel)]="f.asoOption" /> Pure ASO</label>
              <label class="check-field"><input type="radio" name="aso" [value]="1" [(ngModel)]="f.asoOption" /> BTBG</label>
              <label class="check-field"><input type="radio" name="aso" [value]="2" [(ngModel)]="f.asoOption" /> ASO Type</label>
            </div>
          }
          <app-field label="Entity Type">
            <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Health Connect</option><option [ngValue]="1">Sihat Johor</option></select>
          </app-field>
          <app-field label="Group Company" [wide]="true">
            <input name="groupCompany" list="ws-groups" [(ngModel)]="f.groupCompany" (change)="companyChanged()" />
            <datalist id="ws-groups">@for (g of groups(); track g.code) { <option [value]="g.name + '|' + g.code"></option> }</datalist>
          </app-field>
          <app-field label="Report Type">
            <select name="reportType" [(ngModel)]="reportType"><option [ngValue]="0">By Group Company</option><option [ngValue]="1">By Department</option></select>
          </app-field>
          <app-field label="Department">
            <select name="department" [(ngModel)]="f.department"><option value=""></option>@for (x of departments(); track x) { <option [value]="x">{{ x }}</option> }</select>
          </app-field>
          <app-field label="Subsidiary">
            <select name="subsidiary" [(ngModel)]="f.subsidiary"><option value=""></option>@for (x of subsidiaries(); track x) { <option [value]="x">{{ x }}</option> }</select>
          </app-field>
          <app-field label="Scan"><label class="check-field"><input type="checkbox" name="scan" [(ngModel)]="f.scan" (ngModelChange)="clearList()" /> Scan worksheet no.</label></app-field>
          <app-field [label]="f.scan ? 'Worksheet No' : 'Case No From'"><input name="caseFrom" maxlength="30" [(ngModel)]="f.caseFrom" (keydown.enter)="search()" /></app-field>
          @if (!f.scan) {
            <app-field label="Case No To"><input name="caseTo" maxlength="30" [(ngModel)]="f.caseTo" /></app-field>
          }
          <app-field label="Visit Date From"><app-date-input name="visitFrom" ariaLabel="Visit Date From" [(ngModel)]="f.visitFrom" /></app-field>
          <app-field label="Visit Date To"><app-date-input name="visitTo" ariaLabel="Visit Date To" [(ngModel)]="f.visitTo" /></app-field>
          <app-field label="No Visit Date"><label class="check-field"><input type="checkbox" name="noVisitDate" [(ngModel)]="f.noVisitDate" /> Date of visit empty</label></app-field>
          <app-field label="Reg Date From"><app-date-input name="regFrom" ariaLabel="Reg Date From" [(ngModel)]="f.regFrom" /></app-field>
          <app-field label="Reg Date To"><app-date-input name="regTo" ariaLabel="Reg Date To" [(ngModel)]="f.regTo" /></app-field>
          <app-field label="Clinic Type">
            <select name="clinicType" [(ngModel)]="f.clinicType">
              <option [ngValue]="0">0 - All</option><option [ngValue]="1">1 - Specialised</option><option [ngValue]="2">2 - General Practisional</option><option [ngValue]="3">3 - Others</option>
            </select>
          </app-field>
          <app-field label="Clinic">
            <input name="clinic" list="ws-clinics" [(ngModel)]="f.clinic" />
            <datalist id="ws-clinics">@for (c of clinics(); track c.code) { <option [value]="c.code + ' - ' + c.name"></option> }</datalist>
          </app-field>
          <app-field label="OP Batch From"><input name="opFrom" maxlength="30" [(ngModel)]="f.opBatchFrom" /></app-field>
          <app-field label="OP Batch To"><input name="opTo" maxlength="30" [(ngModel)]="f.opBatchTo" /></app-field>
          <app-field label="Cash Type">
            <select name="cashType" [(ngModel)]="f.cashType"><option value=""></option><option value="C">C-Cashless</option><option value="R">R-Reimbursement</option></select>
          </app-field>
          <app-field label="Treatment">
            <select name="treatment" [(ngModel)]="f.treatment"><option value=""></option>@for (t of treatments(); track t.code) { <option [value]="t.code + '-' + t.name">{{ t.code }}-{{ t.name }}</option> }</select>
          </app-field>
          <app-field label="OMNI Status"><label class="check-field"><input type="checkbox" name="omni" [(ngModel)]="f.omni" /> Y - Yes</label></app-field>
          @if (uob()) {
            <app-field label="Outpatient / flexCare">
              <select name="uob" [(ngModel)]="f.uob" (ngModelChange)="searchUob()"><option [ngValue]="null"></option><option [ngValue]="0">Outpatient</option><option [ngValue]="1">flexCare</option></select>
            </app-field>
          }
          <app-action-bar label="Insert worksheet actions">
            <button appButton variant="primary" [busy]="busy()" [disabled]="uob()" (click)="search()">Search</button>
            <button appButton [disabled]="!rows().length" (click)="checkAll()">Check All</button>
            <button appButton variant="primary" [busy]="saving()" (click)="bordx()">Bordx</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Worksheets" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()" (checkedChange)="check($event)" />
        <p class="totals" aria-label="Totals">
          Selected: <strong>{{ order().length }}</strong> · Total: <strong>{{ result()?.total ?? 0 }}</strong> · Total Amount:
          <strong>{{ money(result()?.totalAmount ?? 0) }}</strong> · Bordx No: <strong>{{ bordxNo() || '-' }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class InsertWorksheetPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = m
  protected readonly bordxTypes = ['INSURER-KIS', 'ASO-KDC', 'CREDIT NOTE-CN', 'SIHAT JOHOR-KSJ', 'VOID-KIV']
  protected readonly groups = signal<Option[]>([])
  protected readonly treatments = signal<Option[]>([])
  protected readonly clinics = signal<Option[]>([])
  protected readonly departments = signal<string[]>([])
  protected readonly subsidiaries = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly result = signal<Result | null>(null)
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  /** Ticked cases in tick order (the desktop's CaseNoArray). */
  protected readonly order = signal<string[]>([])
  protected readonly uob = signal(false)
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly bordxNo = signal('')
  protected reportType = 0
  protected f = this.empty()
  private searched = this.empty()

  protected readonly key = (r: Row) => s(r['CASNumber'])
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'INVPayTo', label: 'Pay To' }, { key: 'PAYCode', label: 'Payor' }, { key: 'CASNumber', label: 'Case No' },
    { key: 'CLMBordxAmt', label: 'Bordx Amt', align: 'right', value: (r) => m(r['CLMBordxAmt']) },
    { key: 'CLMNonGPFee', label: 'Non GP Fee', align: 'right', value: (r) => m(r['CLMNonGPFee']) },
    { key: 'DateVisit', label: 'Visit Date', value: (r) => d(r['DateVisit']) }, { key: 'MBMnumber', label: 'Member No' }, { key: 'Name', label: 'Name' },
    { key: 'ClinicCode', label: 'Clinic' }, { key: 'CLNName', label: 'Clinic Name' }, { key: 'MBMGroupCompany', label: 'Group Company' }, { key: 'PLNCode', label: 'Plan' },
    { key: 'ReasonDesc1', label: 'Diagnosis' }, { key: 'Position', label: 'Position' }, { key: 'InvoiceNo', label: 'Invoice No' }, { key: 'OPBatchNo', label: 'OP Batch' },
    { key: 'CLNFFSStatus', label: 'FFS' }, { key: 'CLMMCOProcFee', label: 'FFS Amt', align: 'right', value: (r) => m(Math.abs(Number(r['CLMMCOProcFee'] ?? 0))) },
    { key: 'GRPMCOStatus', label: 'Grp FFS' }, { key: 'CLMTotalActual', label: 'Bill Amt', align: 'right', value: (r) => m(r['CLMTotalActual']) },
    { key: 'Subsidiary', label: 'Subsidiary' }, { key: 'INVBillAmt', label: 'Inv Bill Amt', align: 'right', value: (r) => m(r['INVBillAmt']) },
    { key: 'CLMPayableByMedix2C', label: 'Payable', align: 'right', value: (r) => m(r['CLMPayableByMedix2C']) }, { key: 'CLNType', label: 'Clinic Type' },
    { key: 'CASClaimsApp', label: 'Claims App' }, { key: 'MBMPolicyNo', label: 'Policy No' }, { key: 'DoctorName', label: 'Doctor' },
    { key: 'CLMCoPayColl', label: 'Co-Pay', align: 'right', value: (r) => m(r['CLMCoPayColl']) }, { key: 'CoverID', label: 'Cov' },
    { key: 'CLMSubsidiary', label: 'Claim Subsidiary' }, { key: 'clmcostcenter', label: 'Cost Center' }, { key: 'CLMDeptName', label: 'Department' },
  ]

  private empty() {
    return {
      bordxType: null as number | null, asoOption: null as number | null, entity: 0, groupCompany: '', department: '', subsidiary: '', scan: false, caseFrom: '',
      caseTo: '', visitFrom: '', visitTo: '', noVisitDate: false, regFrom: '', regTo: '', clinicType: 0, clinic: '', opBatchFrom: '', opBatchTo: '', cashType: '',
      treatment: '', omni: false, uob: null as number | null,
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ groups: Option[]; treatments: Option[]; clinics: Option[] }>(`${BASE}/lookups`))
      this.groups.set(l.groups)
      this.treatments.set(l.treatments)
      this.clinics.set(l.clinics)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CboPayor_LostFocus / CboPayor_Click. */
  protected async companyChanged(): Promise<void> {
    this.uob.set(false)
    this.clearList()
    if (!this.f.groupCompany.trim()) return
    try {
      const c = await firstValueFrom(this.http.get<{ departments: string[]; subsidiaries: string[] }>(`${BASE}/company`, { params: { groupCompany: this.f.groupCompany } }))
      this.departments.set(c.departments)
      this.subsidiaries.set(c.subsidiaries)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected clearList(): void {
    this.rows.set([])
    this.result.set(null)
    this.checked.set(new Set())
    this.order.set([])
  }

  /** cmdSearch_Click: scan mode appends one worksheet to the list. */
  protected async search(uob: number | null = null): Promise<void> {
    if (this.f.scan && this.rows().some((r) => s(r['CASNumber']) === this.f.caseFrom.trim())) return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
    this.bordxNo.set('')
    try {
      const body = { ...this.f, uob }
      const r = await firstValueFrom(this.http.post<Result>(`${BASE}/search`, body))
      this.searched = { ...this.f, uob }
      if (this.f.scan) {
        const rows = [...this.rows(), ...r.rows]
        this.rows.set(rows)
        this.result.set({ ...r, rows, total: rows.length, totalAmount: rows.reduce((t, x) => t + Number(x['CLMBordxAmt'] ?? 0), 0) })
        this.f.caseFrom = ''
      } else {
        this.clearList()
        this.rows.set(r.rows)
        this.result.set(r)
      }
      if (uob === null) this.uob.set(r.uob)
    } catch (e) {
      if (!this.f.scan) this.clearList()
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** CboUOB_Click. */
  protected async searchUob(): Promise<void> {
    if (this.f.uob === null) return this.notify.notify('Please Select Search Criteria!', 'error')
    if (!this.f.cashType) return
    await this.search(this.f.uob)
  }

  /** lstWorksheetTrackList_ItemCheck: a row the desktop refuses stays unticked with its message. */
  protected check(keys: Set<string>): void {
    const before = this.checked()
    const order = this.order().filter((k) => keys.has(k))
    for (const r of this.rows()) {
      const k = this.key(r)
      if (!keys.has(k) || before.has(k)) continue
      if (r['TickError']) {
        this.notify.notify(String(r['TickError']), 'error')
        keys.delete(k)
      } else order.push(k)
    }
    this.checked.set(keys)
    this.order.set(order)
  }

  /** CmdCheck_Click. */
  protected checkAll(): void {
    this.check(new Set(this.rows().map(this.key)))
  }

  /** CmdBordx_Click. */
  protected async bordx(): Promise<void> {
    if (!this.rows().length) return this.notify.notify('Please select a record before proceeding!', 'error')
    if (!this.order().length) return this.notify.notify('Please select at least one record to be updated! ', 'error')
    if (this.f.bordxType === null) return this.notify.notify('Please select Bordx Type!', 'error')
    if (this.uob() && this.searched.uob === null) return this.notify.notify('Please select Flexcare or Outpatient Type!', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    this.saving.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ bordxNo: string; message: string }>(`${BASE}/bordx`, { search: this.searched, cases: this.order(), reportType: this.reportType }))
      this.clearList()
      this.bordxNo.set(r.bordxNo)
      this.notify.notify(r.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  /** cmdPrint_Click: the list as Excel. */
  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((r) => this.columns.map((c) => (c.value ? c.value(r) : s(r[c.key]))))], 'WorksheetBordx.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.clearList()
    this.uob.set(false)
    this.bordxNo.set('')
  }
}
