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

const BASE = '/api/payment/member-legacy'
const MEMBER = '/api/payment/member'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })

/** Payment > To Member (old) (frmPaymentMember). */
@Component({
  selector: 'app-payment-member-legacy-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Payment to Member (old)" subtitle="Replaces frmPaymentMember">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Scan"><label class="check-field"><input type="checkbox" name="scan" [(ngModel)]="f.scan" /> Scan worksheet no.</label></app-field>
          <app-field [label]="f.scan ? 'Worksheet No' : 'Worksheet No From'"><input name="caseFrom" maxlength="30" [(ngModel)]="f.caseFrom" (keydown.enter)="search()" /></app-field>
          @if (!f.scan) {
            <app-field label="Worksheet No To"><input name="caseTo" maxlength="30" [(ngModel)]="f.caseTo" /></app-field>
          }
          <app-field label="Member No"><input name="mbmNo" maxlength="30" [(ngModel)]="f.mbmNo" /></app-field>
          <app-field label="Member Name"><input name="mbmName" maxlength="200" [(ngModel)]="f.mbmName" /></app-field>
          <app-field label="Group Company" [wide]="true">
            <input name="company" list="pml-companies" [(ngModel)]="f.company" />
            <datalist id="pml-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-field label="Payment">
            <select name="payment" [(ngModel)]="f.payment">@for (p of payments; track $index) { <option [ngValue]="$index">{{ $index }} - {{ p }}</option> }</select>
          </app-field>
          <app-field label="Bordx Status">
            <select name="bordx" [(ngModel)]="f.bordx"><option [ngValue]="null"></option><option [ngValue]="0">0 - Not Bordx</option><option [ngValue]="1">1 - Bordx</option></select>
          </app-field>
          <app-field label="PV No From"><input name="pvFrom" maxlength="30" [(ngModel)]="f.pvFrom" /></app-field>
          <app-field label="PV No To"><input name="pvTo" maxlength="30" [(ngModel)]="f.pvTo" /></app-field>
          <app-field label="PV Date From"><app-date-input name="pvDateFrom" ariaLabel="PV Date From" [(ngModel)]="f.pvDateFrom" /></app-field>
          <app-field label="PV Date To"><app-date-input name="pvDateTo" ariaLabel="PV Date To" [(ngModel)]="f.pvDateTo" /></app-field>
          <app-field label="Cheque No From"><input name="chequeFrom" maxlength="50" [(ngModel)]="f.chequeFrom" /></app-field>
          <app-field label="Cheque No To"><input name="chequeTo" maxlength="50" [(ngModel)]="f.chequeTo" /></app-field>
          <app-field label="CR No From"><input name="crFrom" maxlength="30" [(ngModel)]="f.crFrom" /></app-field>
          <app-field label="CR No To"><input name="crTo" maxlength="30" [(ngModel)]="f.crTo" /></app-field>
          <app-field label="OP Batch From"><input name="opFrom" maxlength="30" [(ngModel)]="f.opFrom" /></app-field>
          <app-field label="OP Batch To"><input name="opTo" maxlength="30" [(ngModel)]="f.opTo" /></app-field>
          <app-field label="A/C Month From"><select name="monthFrom" [(ngModel)]="f.monthFrom"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Month To"><select name="monthTo" [(ngModel)]="f.monthTo"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Year From"><input name="yearFrom" maxlength="4" [(ngModel)]="f.yearFrom" /></app-field>
          <app-field label="A/C Year To"><input name="yearTo" maxlength="4" [(ngModel)]="f.yearTo" /></app-field>
          <app-field label="Empty" [wide]="true">
            <span class="radio-row">
              <label class="check-field"><input type="checkbox" name="blankPv" [(ngModel)]="f.blankPv" /> PV No</label>
              <label class="check-field"><input type="checkbox" name="blankPvDate" [(ngModel)]="f.blankPvDate" /> PV Date</label>
              <label class="check-field"><input type="checkbox" name="blankCheque" [(ngModel)]="f.blankCheque" /> Cheque No</label>
              <label class="check-field"><input type="checkbox" name="blankCr" [(ngModel)]="f.blankCr" /> CR No</label>
              <label class="check-field"><input type="checkbox" name="blankBatch" [(ngModel)]="f.blankBatch" /> OP Batch</label>
              <label class="check-field"><input type="checkbox" name="blankMonth" [(ngModel)]="f.blankMonth" /> A/C Month</label>
              <label class="check-field"><input type="checkbox" name="blankYear" [(ngModel)]="f.blankYear" /> A/C Year</label>
            </span>
          </app-field>
          <app-action-bar label="Payment search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [disabled]="!rows().length" (click)="checkAll()">Check All</button>
            <button appButton (click)="add()">Add</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Payment voucher">
          <legend>Payment Voucher</legend>
          <app-field label="PV No"><input name="pvNo" [(ngModel)]="pv.pvNo" disabled /></app-field>
          <app-field label="PV Date" [required]="true"><app-date-input name="pvDate" ariaLabel="PV Date" [(ngModel)]="pv.pvDate" /></app-field>
          <app-field label="Pay To" [required]="true"><input name="payTo" maxlength="60" [(ngModel)]="pv.payTo" /></app-field>
          <app-field label="Cheque No" [required]="true"><input name="chequeNo" maxlength="50" [(ngModel)]="pv.chequeNo" /></app-field>
          <app-field label="Cheque Date" [required]="true"><app-date-input name="chequeDate" ariaLabel="Cheque Date" [(ngModel)]="pv.chequeDate" /></app-field>
          <app-field label="Cheque Amt" [required]="true"><input class="num" name="chequeAmt" maxlength="20" [(ngModel)]="pv.chequeAmt" (blur)="wording()" /></app-field>
          <app-field label="Amount in Words" [wide]="true"><input name="wording" maxlength="500" [(ngModel)]="pv.wording" /></app-field>
          <app-field label="Bank" [required]="true">
            <select name="bank" [(ngModel)]="pv.bank"><option value=""></option>@for (b of banks(); track b.code) { <option [value]="b.code + ' - ' + b.name">{{ b.code }} - {{ b.name }}</option> }</select>
          </app-field>
          <app-field label="Account No"><input name="account" maxlength="30" [(ngModel)]="pv.account" /></app-field>
          <app-field label="A/C Period Month" [required]="true"><select name="month" [(ngModel)]="pv.month"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Period Year" [required]="true"><input name="year" maxlength="4" [(ngModel)]="pv.year" /></app-field>
          <app-field label="Remark" [wide]="true"><input name="remark" maxlength="200" [(ngModel)]="pv.remark" /></app-field>
          <app-action-bar label="Payment voucher actions">
            @if (editing()) {
              <button appButton variant="primary" (click)="update()">Update</button>
            } @else {
              <button appButton variant="primary" [busy]="saving()" (click)="save()">Save</button>
              <button appButton [disabled]="!pv.index || checked().size > 0" (click)="editing.set(true)">Edit</button>
            }
            <button appButton (click)="preview()">Preview PV</button>
            <button appButton (click)="printPv()">Print PV</button>
            <button appButton (click)="printCheque()">Print Cheque</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Member payments (old)" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="check($event)" [selectable]="true" [selectedKey]="selectedKey()" (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong> · Selected amount: <strong>{{ money(amount()) }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class PaymentMemberLegacyPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'))
  protected readonly payments = ['All', 'Paid', 'Unpaid', 'Partial Paid', 'Paid & Partial Paid', 'Unpaid & Partial Paid']
  protected readonly banks = signal<{ code: string; name: string }[]>([])
  protected readonly companies = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly selectedKey = signal<string | null>(null)
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly editing = signal(false)
  protected readonly amount = computed(() => this.rows().filter((x) => this.checked().has(this.key(x))).reduce((t, x) => t + Number(x['AmountToPay'] ?? 0), 0))
  protected f = this.empty()
  protected pv = this.emptyPv()
  protected readonly key = (x: Row) => `${s(x['CASNumber'])}|${s(x['PVMBMIndex'])}`
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'DebitCreditRefNo', label: 'CR No' }, { key: 'Status', label: 'Status' }, { key: 'PatientName', label: 'Patient' }, { key: 'CASNumber', label: 'Case No' },
    money('CLMPayableByMedix2M', 'Payable'), money('Paid', 'Paid'), money('Balance', 'Balance'), money('CLMPostageFee', 'Postage'), { key: 'PVPayto', label: 'Pay To' },
    { key: 'PVNo', label: 'PV No' }, { key: 'PVDate', label: 'PV Date', value: (x) => d(x['PVDate']) }, { key: 'PVChequeNo', label: 'Cheque No' },
    { key: 'PVChequeDate', label: 'Cheque Date', value: (x) => d(x['PVChequeDate']) }, money('PVCLNChequeAmt', 'Cheque Amt'),
    { key: 'period', label: 'A/C Period', value: (x) => (s(x['PAYAccPeriodMth']) ? `${s(x['PAYAccPeriodMth'])}/${s(x['PAYAccPeriodYr'])}` : '') },
    { key: 'PVPayeeBank', label: 'Bank' }, { key: 'PVPayeeACNo', label: 'Account No' }, { key: 'Remark', label: 'Remark' }, { key: 'OPBatchNo', label: 'OP Batch' },
    { key: 'bordx', label: 'Bordx', value: (x) => (x['ClaimBordxStatus'] == null ? '' : s(x['ClaimBordxStatus']) === '0' ? 'No' : 'Yes') },
    { key: 'wording', label: 'Wording', value: (x) => `${s(x['PVWordingOne'])}${s(x['PVWordingTwo'])}` },
  ]

  private empty() {
    return {
      scan: true, caseFrom: '', caseTo: '', mbmNo: '', mbmName: '', company: '', payment: 0, bordx: 1 as number | null, pvFrom: '', pvTo: '', pvDateFrom: '',
      pvDateTo: '', chequeFrom: '', chequeTo: '', crFrom: '', crTo: '', opFrom: '', opTo: '', monthFrom: '', monthTo: '', yearFrom: '', yearTo: '', blankPv: false,
      blankPvDate: false, blankCheque: false, blankCr: false, blankBatch: false, blankMonth: false, blankYear: false,
    }
  }

  private emptyPv() {
    return {
      index: '', casNumber: '', pvNo: '', pvDate: '', payTo: '', chequeNo: '', chequeDate: '', chequeAmt: '', wording: '', bank: '', account: '', month: '', year: '',
      remark: '',
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ banks: { code: string; name: string }[]; companies: string[] }>(`${MEMBER}/lookups`))
      this.banks.set(l.banks)
      this.companies.set(l.companies)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CmdSearch_Click: scan mode appends. */
  protected async search(): Promise<void> {
    if (this.f.scan && this.rows().some((x) => s(x['CASNumber']) === this.f.caseFrom.trim())) return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
    this.editing.set(false)
    try {
      const rows = await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, this.f))
      this.rows.set(this.f.scan ? [...this.rows(), ...rows] : rows)
      if (!this.f.scan) this.checked.set(new Set())
      if (this.f.scan) this.f.caseFrom = ''
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected check(keys: Set<string>): void {
    const before = this.checked()
    for (const x of this.rows()) {
      const k = this.key(x)
      if (keys.has(k) && !before.has(k) && x['TickError']) {
        keys.delete(k)
        this.notify.notify(String(x['TickError']), 'error')
      }
    }
    this.editing.set(false)
    this.checked.set(keys)
  }

  protected checkAll(): void {
    this.check(new Set(this.rows().map(this.key)))
  }

  /** CmdAdd_Click. */
  protected add(): void {
    if (!this.checked().size) return this.notify.notify('Please checked at least 1 Claim  ! ', 'error')
    this.editing.set(false)
    this.pv = this.emptyPv()
  }

  /** txtChequeAmt_LostFocus: the amount in words. */
  protected async wording(): Promise<void> {
    const amt = Number(this.pv.chequeAmt.replace(/,/g, ''))
    if (!this.pv.chequeAmt.trim() || Number.isNaN(amt)) return
    try {
      this.pv.wording = (await firstValueFrom(this.http.get<{ wording: string }>(`${BASE}/wording`, { params: { amount: amt } }))).wording
    } catch {
      // The wording can still be typed.
    }
  }

  /** lstPaymentMBM_Click. */
  protected select(x: Row): void {
    this.editing.set(false)
    this.selectedKey.set(this.key(x))
    if (!s(x['PVMBMIndex'])) {
      this.pv = this.emptyPv()
      return
    }
    const bank = this.banks().find((b) => b.code === s(x['PVPayeeBank']))
    this.pv = {
      index: s(x['PVMBMIndex']), casNumber: s(x['CASNumber']), pvNo: s(x['PVNo']), pvDate: d(x['PVDate']), payTo: s(x['PVPayto']), chequeNo: s(x['PVChequeNo']),
      chequeDate: d(x['PVChequeDate']), chequeAmt: s(x['PVCLNChequeAmt']), wording: `${s(x['PVWordingOne'])}${s(x['PVWordingTwo'])}`,
      bank: bank ? `${bank.code} - ${bank.name}` : s(x['PVPayeeBank']), account: s(x['PVPayeeACNo']), month: s(x['PAYAccPeriodMth']).padStart(2, '0'),
      year: s(x['PAYAccPeriodYr']), remark: s(x['Remark']),
    }
  }

  private ticked() {
    return this.rows().filter((x) => this.checked().has(this.key(x))).map((x) => ({ casNumber: s(x['CASNumber']), index: s(x['PVMBMIndex']) || null }))
  }

  /** CmdSave_Click. */
  protected async save(): Promise<void> {
    if (!this.rows().length) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    this.saving.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ pvNo: string; message: string }>(`${BASE}/save`, { ...this.pv, ticked: this.ticked() }))
      this.pv = this.emptyPv()
      this.checked.set(new Set())
      this.rows.set(await firstValueFrom(this.http.get<Row[]>(`${BASE}/pv/${encodeURIComponent(r.pvNo)}`)))
      this.notify.notify(r.message, 'success')
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
      const pvNo = this.pv.pvNo
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/update`, { ...this.pv, ticked: [] }))
      this.editing.set(false)
      this.pv = this.emptyPv()
      this.rows.set(await firstValueFrom(this.http.get<Row[]>(`${BASE}/pv/${encodeURIComponent(pvNo)}`)))
      this.notify.notify(r.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async preview(): Promise<void> {
    if (!this.checked().size) return this.notify.notify('Please select the record', 'error')
    await this.download(this.http.post(`${MEMBER}/voucher`, { cases: this.ticked().map((t) => t.casNumber), payTo: this.pv.payTo, pvNo: this.pv.pvNo }, { responseType: 'blob' }), 'PV_Preview.xlsx')
  }

  /** CmdPrintPV_Click. */
  protected async printPv(): Promise<void> {
    if (!this.pv.pvNo.trim()) return this.notify.notify('Please Enter or Generate PV No', 'error')
    await this.download(this.http.get(`${BASE}/voucher`, { params: { pvNo: this.pv.pvNo, payTo: this.pv.payTo }, responseType: 'blob' }), 'PV.xlsx')
  }

  /** cmdPrintChq_Click. */
  protected async printCheque(): Promise<void> {
    if (!(await this.notify.ask('Do you want to print now?'))) return
    if (!this.pv.pvNo.trim()) return this.notify.notify('Please Enter or Generate PV No', 'error')
    if (!this.pv.payTo.trim()) return this.notify.notify('Please Enter Payee', 'error')
    if (!this.pv.chequeNo.trim()) return this.notify.notify('Please Enter Cheque No', 'error')
    await this.download(this.http.post(`${MEMBER}/cheques/print`, { from: this.pv.chequeNo.trim(), to: this.pv.chequeNo.trim() }, { responseType: 'blob' }), 'Cheque.xlsx')
  }

  private async download(request: ReturnType<HttpClient['get']>, name: string): Promise<void> {
    try {
      saveBlob((await firstValueFrom(request)) as Blob, name)
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'PaymentMemberOld.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.pv = this.emptyPv()
    this.editing.set(false)
    this.rows.set([])
    this.checked.set(new Set())
  }
}
