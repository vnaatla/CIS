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

const BASE = '/api/payment/clinic-invoice'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })

/** Payment > To Clinic - Invoice (frmPaymentCLNInvoice, frmInvoiceDetPopup, frmPrintPVRange, frmPrintChequeRange). */
@Component({
  selector: 'app-payment-clinic-invoice-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Payment to Clinic - Invoice" subtitle="Replaces frmPaymentCLNInvoice">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Clinic" [wide]="true">
            <input name="clinic" list="pci-clinics" [(ngModel)]="f.clinic" />
            <datalist id="pci-clinics">@for (c of clinics(); track c.code) { <option [value]="c.code + ' - ' + c.name"></option> }</datalist>
          </app-field>
          <app-field label="Invoice Year From"><input name="yearFrom" maxlength="4" [(ngModel)]="f.yearFrom" /></app-field>
          <app-field label="Invoice Year To"><input name="yearTo" maxlength="4" [(ngModel)]="f.yearTo" /></app-field>
          <app-field label="Invoice Month From"><select name="monthFrom" [(ngModel)]="f.monthFrom"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="Invoice Month To"><select name="monthTo" [(ngModel)]="f.monthTo"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="Payment">
            <select name="payment" [(ngModel)]="f.payment"><option [ngValue]="0">0 - All</option><option [ngValue]="1">1 - Paid</option><option [ngValue]="2">2 - Unpaid</option></select>
          </app-field>
          <app-field label="Entity Type">
            <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Health Connect</option><option [ngValue]="1">Sihat Johor</option></select>
          </app-field>
          <app-field label="PV No From"><input name="pvFrom" maxlength="30" [(ngModel)]="f.pvFrom" /></app-field>
          <app-field label="PV No To"><input name="pvTo" maxlength="30" [(ngModel)]="f.pvTo" /></app-field>
          <app-field label="PV Date From"><app-date-input name="pvDateFrom" ariaLabel="PV Date From" [(ngModel)]="f.pvDateFrom" /></app-field>
          <app-field label="PV Date To"><app-date-input name="pvDateTo" ariaLabel="PV Date To" [(ngModel)]="f.pvDateTo" /></app-field>
          <app-field label="A/C Month From"><select name="acMonthFrom" [(ngModel)]="f.acMonthFrom"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Month To"><select name="acMonthTo" [(ngModel)]="f.acMonthTo"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Year From"><input name="acYearFrom" maxlength="4" [(ngModel)]="f.acYearFrom" /></app-field>
          <app-field label="A/C Year To"><input name="acYearTo" maxlength="4" [(ngModel)]="f.acYearTo" /></app-field>
          <app-field label="Empty" [wide]="true">
            <span class="radio-row">
              <label class="check-field"><input type="checkbox" name="blankPv" [(ngModel)]="f.blankPv" /> PV No</label>
              <label class="check-field"><input type="checkbox" name="blankPvDate" [(ngModel)]="f.blankPvDate" /> PV Date</label>
              <label class="check-field"><input type="checkbox" name="blankAcMonth" [(ngModel)]="f.blankAcMonth" /> A/C Month</label>
              <label class="check-field"><input type="checkbox" name="blankAcYear" [(ngModel)]="f.blankAcYear" /> A/C Year</label>
              <label class="check-field"><input type="checkbox" name="append" [(ngModel)]="append" /> Append records</label>
            </span>
          </app-field>
          <app-action-bar label="Payment search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [disabled]="!rows().length" (click)="checkAll()">Check All</button>
            <button appButton [disabled]="!selected()" (click)="invoices()">Invoices</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Payment voucher">
          <legend>Payment Voucher</legend>
          <app-field label="PV No"><input name="pvNo" maxlength="30" [(ngModel)]="pv.pvNo" [disabled]="!editing()" /></app-field>
          <app-field label="PV Date"><app-date-input name="pvDate" ariaLabel="PV Date" [(ngModel)]="pv.pvDate" /></app-field>
          <app-field label="Pay To"><input name="payTo" maxlength="200" [(ngModel)]="pv.payTo" /></app-field>
          <app-field label="Cheque No" [required]="true"><input name="chequeNo" maxlength="30" [(ngModel)]="pv.chequeNo" /></app-field>
          <app-field label="Cheque Date"><app-date-input name="chequeDate" ariaLabel="Cheque Date" [(ngModel)]="pv.chequeDate" /></app-field>
          <app-field label="Cheque Amt"><input class="num" name="chequeAmt" maxlength="20" [(ngModel)]="pv.chequeAmt" (ngModelChange)="wording()" /></app-field>
          <app-field label="Amount in Words" [wide]="true"><input name="wording" maxlength="500" [(ngModel)]="pv.wording" /></app-field>
          <app-field label="Bank">
            <select name="bank" [(ngModel)]="pv.bank"><option value=""></option>@for (b of banks(); track b.code) { <option [value]="b.code + ' - ' + b.name">{{ b.code }} - {{ b.name }}</option> }</select>
          </app-field>
          <app-field label="A/C Period Month"><select name="month" [(ngModel)]="pv.month"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Period Year"><input name="year" maxlength="4" [(ngModel)]="pv.year" /></app-field>
          <app-field label="Remark" [wide]="true"><input name="remark" maxlength="500" [(ngModel)]="pv.remark" /></app-field>
          <app-action-bar label="Payment voucher actions">
            @if (editing()) {
              <button appButton variant="primary" (click)="update()">Update</button>
            } @else {
              <button appButton variant="primary" [busy]="saving()" (click)="save()">Save</button>
              <button appButton [disabled]="!pv.pvNo || checked().size > 0" (click)="editing.set(true)">Edit</button>
            }
            <button appButton (click)="voucher(false)">Preview PV</button>
            <button appButton (click)="voucher(true)">Print PV</button>
            <button appButton (click)="printCheque()">Print Cheque</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Clinic invoice payments" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="check($event)" [selectable]="true" [selectedKey]="selectedKey()" (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong> · Amount to pay: <strong>{{ money(amount()) }}</strong></p>

        @if (invoiceRows().length) {
          <app-data-grid caption="Invoices" [columns]="invoiceColumns" [rows]="invoiceRows()" [rowKey]="invoiceKey" />
        }

        <fieldset aria-label="PV and cheque ranges">
          <legend>PV / Cheque Range</legend>
          <app-field label="PV No From"><input name="rangePvFrom" maxlength="30" [(ngModel)]="range.pvFrom" /></app-field>
          <app-field label="PV No To"><input name="rangePvTo" maxlength="30" [(ngModel)]="range.pvTo" /></app-field>
          <app-field label="Cheque No From"><input name="rangeChqFrom" maxlength="30" [(ngModel)]="range.chqFrom" /></app-field>
          <app-field label="Cheque No To"><input name="rangeChqTo" maxlength="30" [(ngModel)]="range.chqTo" /></app-field>
          <app-action-bar label="Range actions">
            <button appButton (click)="pvRange()">Search PV</button>
            <button appButton (click)="printPvRange()">Print PV Range</button>
            <button appButton (click)="cheques()">Search Cheques</button>
            <button appButton (click)="printCheques(range.chqFrom, range.chqTo)">Print Cheques</button>
          </app-action-bar>
        </fieldset>
        @if (rangeRows().length) {
          <app-data-grid caption="Range" [columns]="rangeColumns()" [rows]="rangeRows()" [rowKey]="rangeKey" />
        }
      </div>
    </app-page-shell>
  `,
})
export class PaymentClinicInvoicePage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'))
  protected readonly banks = signal<{ code: string; name: string }[]>([])
  protected readonly clinics = signal<{ code: string; name: string }[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly invoiceRows = signal<Row[]>([])
  protected readonly rangeRows = signal<Row[]>([])
  protected readonly rangeIsCheque = signal(false)
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly selectedKey = signal<string | null>(null)
  protected readonly selected = signal<Row | null>(null)
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly editing = signal(false)
  protected readonly amount = computed(() => this.rows().filter((r) => this.checked().has(this.key(r))).reduce((t, r) => t + Number(r['Balance'] ?? 0), 0))
  protected append = false
  protected f = this.empty()
  protected pv = this.emptyPv()
  protected range = { pvFrom: '', pvTo: '', chqFrom: '', chqTo: '' }
  protected readonly key = (r: Row) => [r['ClinicCode'], r['CASInvMonth'], r['CASInvYear'], r['PVCLNNo'], r['PVPayCLNStatus']].map(s).join('|')
  protected readonly invoiceKey = (r: Row) => [r['ClinicCode'], r['InvoiceNo'], r['PVCLNNo']].map(s).join('|')
  protected readonly rangeKey = (r: Row) => [r['PVCLNNo'], r['PVCLNChequeNo'], r['InvoiceNo']].map(s).join('|')
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'Status', label: 'Status' }, { key: 'ClinicCode', label: 'Clinic' }, { key: 'CLNName', label: 'Clinic Name' }, { key: 'CASInvMonth', label: 'Inv Month' },
    { key: 'CASInvYear', label: 'Inv Year' }, money('CLMPayableByMedix2C', 'Payable'), money('CLNPaidAmt', 'Paid'), money('Balance', 'Balance'),
    { key: 'PVCLNDate', label: 'PV Date', value: (r) => d(r['PVCLNDate']) }, { key: 'PVCLNNo', label: 'PV No' }, { key: 'PVCLNChequeNo', label: 'Cheque No' },
    { key: 'PVCLNChequeDate', label: 'Cheque Date', value: (r) => d(r['PVCLNChequeDate']) }, money('PVCLNChequeAmt', 'Cheque Amt'),
    { key: 'PVCLNBank', label: 'Bank' }, { key: 'PVCLNAccount', label: 'A/C' }, { key: 'CLNRemark', label: 'Remark' }, { key: 'CLNBankAcc', label: 'Bank Acc' },
    { key: 'CLNPName', label: 'Payee' }, { key: 'wording', label: 'Wording', value: (r) => `${s(r['PVWordingOne'])}${s(r['PVWordingTwo']) ? ' ' + s(r['PVWordingTwo']) : ''}` },
    { key: 'CLNPAdd1', label: 'Payee Add 1' }, { key: 'CLNPAdd2', label: 'Payee Add 2' }, { key: 'CLNPAdd3', label: 'Payee Add 3' }, money('CLMTotalActual', 'Bill Amt'),
    money('CoPay', 'Co-Pay'), money('CLMFinalApprove', 'Approved'), { key: 'CLNType', label: 'Clinic Type' }, { key: 'ClinicGrpCode', label: 'Grp Code' },
    { key: 'CLNFFSStatus', label: 'FFS' }, { key: 'CLNRegNo', label: 'Reg No' }, { key: 'CLNMCOProcfeePerc', label: 'MCO %' }, { key: 'CLNBankName', label: 'Bank Name' },
  ]
  protected readonly invoiceColumns: GridColumn<Row>[] = [
    { key: 'Status', label: 'Status' }, { key: 'ClinicCode', label: 'Clinic' }, { key: 'CLNName', label: 'Clinic Name' }, { key: 'InvoiceNo', label: 'Invoice No' },
    { key: 'InvoiceDate', label: 'Invoice Date', value: (r) => d(r['InvoiceDate']) }, money('CLMPayableByMedix2C', 'Payable'), money('CLNPaidAmt', 'Paid'),
    money('Balance', 'Balance'), { key: 'PVCLNNo', label: 'PV No' }, { key: 'PVCLNChequeNo', label: 'Cheque No' }, money('CLMFinalApprove', 'Approved'),
  ]

  protected rangeColumns(): GridColumn<Row>[] {
    return this.rangeIsCheque()
      ? [{ key: 'PVCLNChequeNo', label: 'Cheque No' }, { key: 'PVCLNChequeDate', label: 'Cheque Date', value: (r) => d(r['PVCLNChequeDate']) },
          money('PVCLNChequeAmt', 'Amount'), { key: 'CLNPName', label: 'Payee' }]
      : [{ key: 'PVCLNNo', label: 'PV No' }, { key: 'PVCLNDate', label: 'PV Date', value: (r) => d(r['PVCLNDate']) }, { key: 'CLNPName', label: 'Payee' },
          { key: 'InvoiceNo', label: 'Invoice No' }, { key: 'InvoiceDate', label: 'Invoice Date', value: (r) => d(r['InvoiceDate']) },
          { key: 'PVCLNChequeNo', label: 'Cheque No' }, money('PVCLNChequeAmt', 'Cheque Amt'), { key: 'CLNName', label: 'Clinic' }, { key: 'ClinicCode', label: 'Code' },
          money('CLMFinalApprove', 'Approved'), money('CLMMCOProcFee', 'MCO Fee')]
  }

  private empty() {
    return {
      clinic: '', yearFrom: '', yearTo: '', monthFrom: '', monthTo: '', payment: 0, entity: 0, pvFrom: '', pvTo: '', blankPv: false, pvDateFrom: '', pvDateTo: '',
      blankPvDate: false, acMonthFrom: '', acMonthTo: '', blankAcMonth: false, acYearFrom: '', acYearTo: '', blankAcYear: false,
    }
  }

  private emptyPv() {
    return { pvNo: '', pvDate: '', payTo: '', chequeNo: '', chequeDate: '', chequeAmt: '', wording: '', bank: '', month: '', year: '', remark: '' }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ banks: { code: string; name: string }[]; clinics: { code: string; name: string }[] }>('/api/payment/clinic/lookups'))
      this.banks.set(l.banks)
      this.clinics.set(l.clinics)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.editing.set(false)
    this.pv = this.emptyPv()
    try {
      const r = await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, this.f))
      this.setRows(this.append ? [...this.rows(), ...r.filter((x) => !this.rows().some((y) => this.key(y) === this.key(x)))] : r, this.append)
    } catch (e) {
      if (!this.append) this.setRows([])
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private setRows(rows: Row[], keep = false): void {
    this.rows.set(rows)
    this.invoiceRows.set([])
    if (!keep) this.checked.set(new Set())
  }

  /** ProcessCheckedItem: payee, payee address, duplicates and paid rows are refused. */
  protected check(keys: Set<string>): void {
    const before = this.checked()
    for (const r of this.rows()) {
      const k = this.key(r)
      if (!keys.has(k) || before.has(k)) continue
      const dup = this.rows().some((x) => x !== r && keys.has(this.key(x)) && s(x['ClinicCode']) === s(r['ClinicCode']) && s(x['CASInvMonth']) === s(r['CASInvMonth']))
      const error = dup ? 'Please Check Duplicate Records! ' : s(r['TickError'])
      if (error) {
        keys.delete(k)
        this.notify.notify(error, 'error')
      }
    }
    this.checked.set(keys)
  }

  protected checkAll(): void {
    this.check(new Set(this.rows().map(this.key)))
  }

  /** lstPaymentCLN_Click: the row's PV fills the frame. */
  protected select(r: Row): void {
    this.editing.set(false)
    this.selectedKey.set(this.key(r))
    this.selected.set(r)
    const bank = this.banks().find((b) => b.code === s(r['PVCLNBank']))
    this.pv = {
      pvNo: s(r['PVCLNNo']), pvDate: d(r['PVCLNDate']), payTo: s(r['CLNPName']), chequeNo: s(r['PVCLNChequeNo']), chequeDate: d(r['PVCLNChequeDate']),
      chequeAmt: s(r['PVCLNChequeAmt']), wording: `${s(r['PVWordingOne'])}${s(r['PVWordingTwo'])}`, bank: bank ? `${bank.code} - ${bank.name}` : s(r['PVCLNBank']),
      month: '', year: '', remark: s(r['CLNRemark']),
    }
  }

  /** lstPaymentCLN_DblClick: frmInvoiceDetPopup. */
  protected async invoices(): Promise<void> {
    const r = this.selected()
    if (!r) return
    try {
      const params = { clinic: s(r['ClinicCode']), month: s(r['CASInvMonth']), year: s(r['CASInvYear']) }
      this.invoiceRows.set(await firstValueFrom(this.http.get<Row[]>(`${BASE}/invoices`, { params })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** txtChequeAmt_Change: the amount wording is recalculated by the server on save; shown here as typed. */
  protected wording(): void {
    if (!this.pv.chequeAmt.trim()) this.pv.wording = ''
  }

  private ticked() {
    return this.rows().filter((r) => this.checked().has(this.key(r))).map((r) => ({ clinicCode: s(r['ClinicCode']), month: s(r['CASInvMonth']), year: s(r['CASInvYear']) }))
  }

  /** CmdSave_Click. */
  protected async save(): Promise<void> {
    if (!this.ticked().length) return this.notify.notify('Please select the record', 'error')
    if (!this.pv.chequeNo.trim()) return this.notify.notify('Please Enter Cheque Number!', 'error')
    if (!(await this.notify.ask('Do you want to update the record ? '))) return
    this.saving.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ first: string; last: string; message: string }>(`${BASE}/save`, { ...this.pv, entity: this.f.entity, ticked: this.ticked() }))
      this.notify.notify(r.message, 'success')
      await this.showPv(r.first, r.last)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  /** CmdUpdate_Click. */
  protected async update(): Promise<void> {
    if (!this.rows().length) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    try {
      await firstValueFrom(this.http.post(`${BASE}/update`, { ...this.pv, ticked: [] }))
      await this.showPv(this.pv.pvNo, this.pv.pvNo)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private async showPv(first: string, last: string): Promise<void> {
    this.editing.set(false)
    this.pv = this.emptyPv()
    this.setRows(await firstValueFrom(this.http.get<Row[]>(`${BASE}/pv`, { params: { pvNo: first, pvLast: last } })))
  }

  protected async voucher(print: boolean): Promise<void> {
    if (print && !this.pv.pvNo.trim()) return this.notify.notify('Please Enter or Generate PV No', 'error')
    if (!print && !this.ticked().length) return this.notify.notify('Please select the record', 'error')
    await this.download(`${BASE}/voucher`, print ? { pvNo: this.pv.pvNo.trim(), ticked: [] } : { pvNo: null, ticked: this.ticked() }, 'PV.xlsx')
  }

  /** cmdPrintChq_Click. */
  protected async printCheque(): Promise<void> {
    if (!(await this.notify.ask('Do you want to print now?'))) return
    if (!this.pv.pvNo.trim()) return this.notify.notify('Please Enter or Generate PV No', 'error')
    if (!this.pv.payTo.trim()) return this.notify.notify('Please Enter Payee', 'error')
    if (!this.pv.chequeNo.trim()) return this.notify.notify('Please Enter Cheque No', 'error')
    await this.download(`${BASE}/cheques/print`, { from: this.pv.chequeNo.trim(), to: this.pv.chequeNo.trim() }, 'Cheque.xlsx')
  }

  protected async pvRange(): Promise<void> {
    await this.loadRange(`${BASE}/pv-range`, { from: this.range.pvFrom, to: this.range.pvTo }, false)
  }

  protected async printPvRange(): Promise<void> {
    if (!this.rangeRows().length || this.rangeIsCheque()) return this.notify.notify('Please select PV to print', 'error')
    if (!(await this.notify.ask('Do you want to print now?'))) return
    await this.download(`${BASE}/pv-range/print`, { from: this.range.pvFrom, to: this.range.pvTo }, 'PVRange.xlsx')
  }

  protected async cheques(): Promise<void> {
    await this.loadRange(`${BASE}/cheques`, { from: this.range.chqFrom, to: this.range.chqTo }, true)
  }

  protected async printCheques(from: string, to: string): Promise<void> {
    if (!this.rangeRows().length || !this.rangeIsCheque()) return this.notify.notify('Please select cheque to print', 'error')
    if (!(await this.notify.ask('Do you want to print now?'))) return
    await this.download(`${BASE}/cheques/print`, { from, to }, 'Cheques.xlsx')
  }

  private async loadRange(url: string, body: object, cheque: boolean): Promise<void> {
    this.rangeRows.set([])
    try {
      const r = await firstValueFrom(this.http.post<Row[]>(url, body))
      this.rangeIsCheque.set(cheque)
      this.rangeRows.set(r)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private async download(url: string, body: object, name: string): Promise<void> {
    try {
      saveBlob(await firstValueFrom(this.http.post(url, body, { responseType: 'blob' })), name)
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((r) => this.columns.map((c) => (c.value ? c.value(r) : s(r[c.key]))))], 'PaymentClinicInvoice.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.pv = this.emptyPv()
    this.editing.set(false)
    this.selected.set(null)
    this.rangeRows.set([])
    this.setRows([])
  }
}
