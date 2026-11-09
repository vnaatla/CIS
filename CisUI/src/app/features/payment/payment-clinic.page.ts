import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob, saveCsv } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/payment/clinic'
type Row = Record<string, unknown>
interface Bank {
  code: string
  name: string
  pvCode: string
}
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })

/** Payment > To Clinic (frmPaymentClinic). */
@Component({
  selector: 'app-payment-clinic-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Payment to Clinic" subtitle="Replaces frmPaymentClinic">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Group Company" [wide]="true">
            <input name="company" list="pay-companies" [(ngModel)]="f.company" />
            <datalist id="pay-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-field label="Clinic">
            <input name="clinic" list="pay-clinics" [(ngModel)]="f.clinic" />
            <datalist id="pay-clinics">@for (c of clinics(); track c.code) { <option [value]="c.code + ' - ' + c.name"></option> }</datalist>
          </app-field>
          <app-field label="Payment">
            <select name="payment" [(ngModel)]="f.payment"><option [ngValue]="0">0 - All</option><option [ngValue]="1">1 - Paid</option><option [ngValue]="2">2 - Unpaid</option></select>
          </app-field>
          <app-field label="Entity Type">
            <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Health Connect</option><option [ngValue]="1">Sihat Johor</option></select>
          </app-field>
          <app-field label="Scan"><label class="check-field"><input type="checkbox" name="scan" [(ngModel)]="f.scan" /> Scan worksheet no.</label></app-field>
          <app-field [label]="f.scan ? 'Worksheet No' : 'Worksheet No From'"><input name="caseFrom" maxlength="30" [(ngModel)]="f.caseFrom" (keydown.enter)="search()" /></app-field>
          @if (!f.scan) {
            <app-field label="Worksheet No To"><input name="caseTo" maxlength="30" [(ngModel)]="f.caseTo" /></app-field>
          }
          <app-field label="Append"><label class="check-field"><input type="checkbox" name="append" [(ngModel)]="append" /> Append records</label></app-field>
          <app-field label="Invoice No"><input name="invoiceNo" maxlength="50" [(ngModel)]="f.invoiceNo" /></app-field>
          <app-field label="Invoice Date From"><app-date-input name="invFrom" ariaLabel="Invoice Date From" [(ngModel)]="f.invFrom" /></app-field>
          <app-field label="Invoice Date To"><app-date-input name="invTo" ariaLabel="Invoice Date To" [(ngModel)]="f.invTo" /></app-field>
          <app-field label="OP Batch From"><input name="opFrom" maxlength="30" [(ngModel)]="f.opFrom" /></app-field>
          <app-field label="OP Batch To"><input name="opTo" maxlength="30" [(ngModel)]="f.opTo" /></app-field>
          <app-field label="Invoice Amt From"><input class="num" name="amtFrom" maxlength="20" [(ngModel)]="f.amtFrom" /></app-field>
          <app-field label="Invoice Amt To"><input class="num" name="amtTo" maxlength="20" [(ngModel)]="f.amtTo" /></app-field>
          <app-field label="PV No From"><input name="pvFrom" maxlength="30" [(ngModel)]="f.pvFrom" /></app-field>
          <app-field label="PV No To"><input name="pvTo" maxlength="30" [(ngModel)]="f.pvTo" /></app-field>
          <app-field label="PV Date From"><app-date-input name="pvDateFrom" ariaLabel="PV Date From" [(ngModel)]="f.pvDateFrom" /></app-field>
          <app-field label="PV Date To"><app-date-input name="pvDateTo" ariaLabel="PV Date To" [(ngModel)]="f.pvDateTo" /></app-field>
          <app-field label="A/C Month From"><select name="mthFrom" [(ngModel)]="f.mthFrom"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Month To"><select name="mthTo" [(ngModel)]="f.mthTo"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Year From"><input name="yrFrom" maxlength="4" [(ngModel)]="f.yrFrom" /></app-field>
          <app-field label="A/C Year To"><input name="yrTo" maxlength="4" [(ngModel)]="f.yrTo" /></app-field>
          <app-field label="SUN Inv No From"><input name="sunFrom" maxlength="30" [(ngModel)]="f.sunFrom" /></app-field>
          <app-field label="SUN Inv No To"><input name="sunTo" maxlength="30" [(ngModel)]="f.sunTo" /></app-field>
          <app-field label="Empty" [wide]="true">
            <span class="radio-row">
              <label class="check-field"><input type="checkbox" name="blankInvoiceNo" [(ngModel)]="f.blankInvoiceNo" /> Invoice No</label>
              <label class="check-field"><input type="checkbox" name="blankInvDate" [(ngModel)]="f.blankInvDate" /> Invoice Date</label>
              <label class="check-field"><input type="checkbox" name="blankBatch" [(ngModel)]="f.blankBatch" /> OP Batch</label>
              <label class="check-field"><input type="checkbox" name="blankAmt" [(ngModel)]="f.blankAmt" /> Invoice Amt</label>
              <label class="check-field"><input type="checkbox" name="blankPv" [(ngModel)]="f.blankPv" /> PV No</label>
              <label class="check-field"><input type="checkbox" name="blankPvDate" [(ngModel)]="f.blankPvDate" /> PV Date</label>
              <label class="check-field"><input type="checkbox" name="blankMth" [(ngModel)]="f.blankMth" /> A/C Month</label>
              <label class="check-field"><input type="checkbox" name="blankYr" [(ngModel)]="f.blankYr" /> A/C Year</label>
            </span>
          </app-field>
          <app-field label="Upload Invoice Nos"><input type="file" name="upload" accept=".csv,.txt" (change)="upload($event)" /></app-field>
          <app-action-bar label="Payment search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [disabled]="!rows().length" (click)="checkAll()">Check All</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Payment voucher">
          <legend>Payment Voucher</legend>
          <app-field label="PV No"><input name="pvNo" maxlength="30" [(ngModel)]="pv.pvNo" [disabled]="!editing()" /></app-field>
          <app-field label="PV Date" [required]="true"><app-date-input name="pvDate" ariaLabel="PV Date" [(ngModel)]="pv.pvDate" /></app-field>
          <app-field label="Cheque No"><input name="chequeNo" maxlength="30" [(ngModel)]="pv.chequeNo" /></app-field>
          <app-field label="Cheque Date"><app-date-input name="chequeDate" ariaLabel="Cheque Date" [(ngModel)]="pv.chequeDate" /></app-field>
          <app-field label="Cheque Amt" [required]="true"><input class="num" name="chequeAmt" maxlength="20" [(ngModel)]="pv.chequeAmt" /></app-field>
          <app-field label="Bank" [required]="true">
            <select name="bank" [(ngModel)]="pv.bank"><option value=""></option>@for (b of banks(); track b.code) { <option [value]="b.code + ' - ' + b.name">{{ b.code }} - {{ b.name }}</option> }</select>
          </app-field>
          <app-field label="A/C No"><input name="account" maxlength="50" [(ngModel)]="pv.account" /></app-field>
          <app-field label="A/C Period Month" [required]="true"><select name="month" [(ngModel)]="pv.month"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Period Year" [required]="true"><input name="year" maxlength="4" [(ngModel)]="pv.year" /></app-field>
          <app-field label="Remark" [wide]="true"><input name="remark" maxlength="500" [(ngModel)]="pv.remark" /></app-field>
          <app-action-bar label="Payment voucher actions">
            @if (editing()) {
              <button appButton variant="primary" (click)="update()">Update</button>
            } @else {
              <button appButton variant="primary" [busy]="saving()" (click)="save()">Save</button>
              <button appButton [disabled]="!pv.index" (click)="editing.set(true)">Edit</button>
            }
            <button appButton (click)="voucher(false)">Preview PV</button>
            <button appButton (click)="voucher(true)">Print PV</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Clinic payments" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="check($event)" [selectable]="true" [selectedKey]="selectedKey()" (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong> · Amount to pay: <strong>{{ money(amount()) }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class PaymentClinicPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'))
  protected readonly banks = signal<Bank[]>([])
  protected readonly companies = signal<string[]>([])
  protected readonly clinics = signal<{ code: string; name: string }[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly editing = signal(false)
  protected readonly selectedKey = signal<string | null>(null)
  protected readonly amount = computed(() => this.rows().filter((r) => this.checked().has(this.key(r))).reduce((t, r) => t + Number(r['Balance'] ?? 0), 0))
  protected append = false
  protected f = this.empty()
  protected pv = this.emptyPv()
  protected readonly key = (r: Row) => `${s(r['CASNumber'])}|${s(r['PVCLNIndex'])}`
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'Status', label: 'Status' }, { key: 'ClinicCode', label: 'Clinic' }, { key: 'CLNName', label: 'Clinic Name' }, { key: 'InvoiceNo', label: 'Invoice No' },
    { key: 'InvoiceDate', label: 'Invoice Date', value: (r) => d(r['InvoiceDate']) }, { key: 'BillNo', label: 'Bill No' }, money('Payable', 'Payable'),
    money('Paid', 'Paid'), money('Balance', 'Balance'), { key: 'PatientName', label: 'Patient' }, { key: 'CASNumber', label: 'Case No' },
    { key: 'BordxRefNo', label: 'Bordx No' }, { key: 'DateVisit', label: 'Visit Date', value: (r) => d(r['DateVisit']) },
    { key: 'PVCLNDate', label: 'PV Date', value: (r) => d(r['PVCLNDate']) }, { key: 'INVEnteredBy', label: 'Entered By' }, { key: 'PVCLNNo', label: 'PV No' },
    { key: 'PVCLNChequeNo', label: 'Cheque No' }, { key: 'PVCLNChequeDate', label: 'Cheque Date', value: (r) => d(r['PVCLNChequeDate']) },
    money('PVCLNChequeAmt', 'Cheque Amt'), { key: 'period', label: 'A/C Period', value: (r) => (s(r['CLNAccPeriodMth']) ? `${s(r['CLNAccPeriodMth'])}/${s(r['CLNAccPeriodYr'])}` : '') },
    { key: 'PVCLNBank', label: 'Bank' }, { key: 'PVCLNAccount', label: 'A/C No' }, { key: 'CLNRemark', label: 'Remark' }, { key: 'OPBatchNo', label: 'OP Batch' },
    { key: 'GRPCoCode', label: 'Grp Code' }, { key: 'MBMGroupCompany', label: 'Group Company' }, { key: 'SUNInvNo', label: 'SUN Inv No' },
    { key: 'SUNPostDate', label: 'SUN Post Date', value: (r) => d(r['SUNPostDate']) }, money('CLMMCOProcFee', 'MCO Fee'), money('CLMCoPayCLNColl', 'Co-Pay'),
    money('CLMBordxAmt', 'Bordx Amt'), money('CLMFinalApprove', 'Approved'), money('CLMTotalActual', 'Bill Amt'),
  ]

  private empty() {
    return {
      company: '', clinic: '', payment: 0, entity: 0, scan: true, caseFrom: '', caseTo: '', invoiceNo: '', blankInvoiceNo: false, invFrom: '', invTo: '',
      blankBatch: false, opFrom: '', opTo: '', blankInvDate: false, amtFrom: '', amtTo: '', blankAmt: false, pvFrom: '', pvTo: '', blankPv: false, pvDateFrom: '',
      pvDateTo: '', blankPvDate: false, mthFrom: '', mthTo: '', blankMth: false, yrFrom: '', yrTo: '', blankYr: false, sunFrom: '', sunTo: '',
    }
  }

  private emptyPv() {
    return { index: '', pvNo: '', pvDate: '', chequeNo: '', chequeDate: '', chequeAmt: '', bank: 'CB - CITIBANK', account: '', month: '', year: '', remark: '' }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ banks: Bank[]; companies: string[]; clinics: { code: string; name: string }[] }>(`${BASE}/lookups`))
      this.banks.set(l.banks)
      this.companies.set(l.companies)
      this.clinics.set(l.clinics)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdSearch_Click: scan / append add to the list. */
  protected async search(body: object = this.f): Promise<void> {
    const keep = this.f.scan || this.append
    if (this.f.scan && this.rows().some((r) => s(r['CASNumber']) === this.f.caseFrom.trim())) return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, body))
      this.setRows(keep ? [...this.rows(), ...r.filter((x) => !this.rows().some((y) => this.key(y) === this.key(x)))] : r, keep)
      if (this.f.scan) this.f.caseFrom = ''
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private setRows(rows: Row[], keepTicks = false): void {
    this.rows.set(rows)
    if (!keepTicks) this.checked.set(new Set())
  }

  /** cmdUploadDirect_Click: invoice numbers, one per line (first column). */
  protected async upload(event: Event): Promise<void> {
    const file = (event.target as HTMLInputElement).files?.[0]
    if (!file) return this.notify.notify('Please select file to upload!', 'error')
    const invoiceNos = (await file.text()).split(/\r?\n/).slice(1).map((l) => l.split(',')[0].trim()).filter((x) => x)
    await this.search({ invoiceNos })
  }

  /** lstPaymentCLN_ItemCheck: a fully paid case cannot be ticked. */
  protected check(keys: Set<string>): void {
    const before = this.checked()
    for (const r of this.rows()) {
      const k = this.key(r)
      if (keys.has(k) && !before.has(k) && s(r['PVPayCLNStatus']) === 'F') {
        keys.delete(k)
        this.notify.notify('Payment has been made!', 'error')
      }
    }
    this.editing.set(false)
    this.checked.set(keys)
  }

  protected checkAll(): void {
    this.check(new Set(this.rows().map(this.key)))
  }

  /** lstPaymentCLN_Click: a paid row fills the PV frame. */
  protected select(r: Row): void {
    this.editing.set(false)
    this.selectedKey.set(this.key(r))
    if (!s(r['PVCLNIndex'])) {
      this.pv = this.emptyPv()
      return
    }
    const bank = this.banks().find((b) => b.code === s(r['PVCLNBank']))
    this.pv = {
      index: s(r['PVCLNIndex']), pvNo: s(r['PVCLNNo']), pvDate: d(r['PVCLNDate']), chequeNo: s(r['PVCLNChequeNo']), chequeDate: d(r['PVCLNChequeDate']),
      chequeAmt: s(r['PVCLNChequeAmt']), bank: bank ? `${bank.code} - ${bank.name}` : s(r['PVCLNBank']), account: s(r['PVCLNAccount']),
      month: s(r['CLNAccPeriodMth']).padStart(2, '0'), year: s(r['CLNAccPeriodYr']), remark: s(r['CLNRemark']),
    }
  }

  private ticked() {
    return this.rows().filter((r) => this.checked().has(this.key(r))).map((r) => ({ casNumber: s(r['CASNumber']), index: s(r['PVCLNIndex']) || null }))
  }

  /** CmdSave_Click. */
  protected async save(): Promise<void> {
    if (!this.ticked().length) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record ? '))) return
    this.saving.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ pvNo: string; message: string }>(`${BASE}/add`, { ...this.pv, pvNo: null, ticked: this.ticked() }))
      this.notify.notify(r.message, 'success')
      await this.showPv(r.pvNo)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  /** CmdUpdate_Click. */
  protected async update(): Promise<void> {
    if (!this.rows().length || !this.pv.index) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    try {
      await firstValueFrom(this.http.post(`${BASE}/update`, { ...this.pv, ticked: [] }))
      await this.showPv(this.pv.pvNo)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** The PV's rows after a save (Payment2CLNListing by PV no.). */
  private async showPv(pvNo: string): Promise<void> {
    this.editing.set(false)
    this.pv = { ...this.emptyPv(), pvNo }
    this.setRows(await firstValueFrom(this.http.get<Row[]>(`${BASE}/pv/${encodeURIComponent(pvNo)}`)))
  }

  /** CmdPrintPV_Click / CmdPreviewPV_Click. */
  protected async voucher(print: boolean): Promise<void> {
    if (print && !this.pv.pvNo.trim()) return this.notify.notify('Please Enter or Generate PV No', 'error')
    if (!print && !this.ticked().length) return this.notify.notify('Please select the record', 'error')
    try {
      const body = print ? { pvNo: this.pv.pvNo.trim(), cases: null } : { pvNo: null, cases: this.ticked().map((t) => t.casNumber) }
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/voucher`, body, { responseType: 'blob' })), `PV_${print ? this.pv.pvNo.trim() : 'Preview'}.xlsx`)
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  /** cmdPrint_Click: the list as CSV. */
  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((r) => this.columns.map((c) => (c.value ? c.value(r) : s(r[c.key]))))], 'PaymentClinic.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.pv = this.emptyPv()
    this.editing.set(false)
    this.setRows([])
  }
}
