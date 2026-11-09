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

const BASE = '/api/payment/member'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })

/** Payment > To Member (frmPaymentMemberAuto08 with frmPrintPVRange / frmPrintChequeRange in member mode). */
@Component({
  selector: 'app-payment-member-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Payment to Member" subtitle="Replaces frmPaymentMemberAuto08">
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
            <input name="company" list="pm-companies" [(ngModel)]="f.company" />
            <datalist id="pm-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-field label="Payment">
            <select name="payment" [(ngModel)]="f.payment">@for (p of payments; track $index) { <option [ngValue]="$index">{{ $index }} - {{ p }}</option> }</select>
          </app-field>
          <app-field label="Bordx Status">
            <select name="bordx" [(ngModel)]="f.bordx"><option [ngValue]="null"></option><option [ngValue]="0">0 - Not Bordx</option><option [ngValue]="1">1 - Bordx</option></select>
          </app-field>
          <app-field label="Bordx No From"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" /></app-field>
          <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
          <app-field label="BTBG Bordx Status">
            <select name="btbg" [(ngModel)]="f.btbg"><option [ngValue]="null"></option><option [ngValue]="0">0 - Not Bordx</option><option [ngValue]="1">1 - Bordx</option></select>
          </app-field>
          <app-field label="BTBG Bordx From"><input name="btbgFrom" maxlength="20" [(ngModel)]="f.btbgFrom" /></app-field>
          <app-field label="BTBG Bordx To"><input name="btbgTo" maxlength="20" [(ngModel)]="f.btbgTo" /></app-field>
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
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Payment voucher">
          <legend>Payment Voucher</legend>
          <app-field label="PV No"><input name="pvNo" [(ngModel)]="pv.pvNo" disabled /></app-field>
          <app-field label="PV Date"><app-date-input name="pvDate" ariaLabel="PV Date" [(ngModel)]="pv.pvDate" /></app-field>
          <app-field label="Pay To"><input name="payTo" maxlength="60" [(ngModel)]="pv.payTo" /></app-field>
          <app-field label="Cheque No" [required]="true"><input name="chequeNo" maxlength="50" [(ngModel)]="pv.chequeNo" /></app-field>
          <app-field label="Cheque Date"><app-date-input name="chequeDate" ariaLabel="Cheque Date" [(ngModel)]="pv.chequeDate" /></app-field>
          <app-field label="Cheque Amt"><input class="num" name="chequeAmt" maxlength="20" [(ngModel)]="pv.chequeAmt" /></app-field>
          <app-field label="Amount in Words" [wide]="true"><input name="wording" maxlength="500" [(ngModel)]="pv.wording" /></app-field>
          <app-field label="Bank" [required]="true">
            <select name="bank" [(ngModel)]="pv.bank"><option value=""></option>@for (b of banks(); track b.code) { <option [value]="b.code + ' - ' + b.name">{{ b.code }} - {{ b.name }}</option> }</select>
          </app-field>
          <app-field label="A/C Period Month"><select name="month" [(ngModel)]="pv.month"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Period Year"><input name="year" maxlength="4" [(ngModel)]="pv.year" /></app-field>
          <app-field label="Remark" [wide]="true"><input name="remark" maxlength="200" [(ngModel)]="pv.remark" /></app-field>
          <app-field label="Void Payment">
            <select name="void" [(ngModel)]="pv.void"><option [ngValue]="false">0 - Select</option><option [ngValue]="true">1 - VOID</option></select>
          </app-field>
          <app-action-bar label="Payment voucher actions">
            @if (editing()) {
              <button appButton variant="primary" (click)="update()">Update</button>
            } @else {
              <button appButton variant="primary" [busy]="saving()" (click)="save()">Save</button>
              <button appButton [disabled]="!pv.index || checked().size > 0" (click)="editing.set(true)">Edit</button>
            }
            <button appButton (click)="preview()">Preview PV</button>
            <button appButton (click)="printCheque()">Print Cheque</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Member payments" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="check($event)" [selectable]="true" [selectedKey]="selectedKey()" (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong> · Selected amount: <strong>{{ money(amount()) }}</strong></p>

        <fieldset aria-label="PV and cheque ranges">
          <legend>PV / Cheque Range</legend>
          <app-field label="PV No From"><input name="rangePvFrom" maxlength="30" [(ngModel)]="range.pvFrom" /></app-field>
          <app-field label="PV No To"><input name="rangePvTo" maxlength="30" [(ngModel)]="range.pvTo" /></app-field>
          <app-field label="Cheque No From"><input name="rangeChqFrom" maxlength="50" [(ngModel)]="range.chqFrom" /></app-field>
          <app-field label="Cheque No To"><input name="rangeChqTo" maxlength="50" [(ngModel)]="range.chqTo" /></app-field>
          <app-field label="Print Acc No"><label class="check-field"><input type="checkbox" name="account" [(ngModel)]="range.account" /> Account no. only</label></app-field>
          <app-action-bar label="Range actions">
            <button appButton (click)="pvRange()">Search PV</button>
            <button appButton (click)="printPvRange()">Print PV</button>
            <button appButton (click)="cheques()">Search Cheques</button>
            <button appButton (click)="printCheques()">Print Cheques</button>
          </app-action-bar>
        </fieldset>
        @if (rangeRows().length) {
          <app-data-grid caption="Range" [columns]="rangeColumns" [rows]="rangeRows()" [rowKey]="rangeKey" />
        }
      </div>
    </app-page-shell>
  `,
})
export class PaymentMemberPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'))
  protected readonly payments = ['All', 'Paid', 'Unpaid', 'Partial Paid', 'Paid & Partial Paid', 'Unpaid & Partial Paid', 'Paid & Payment Void']
  protected readonly banks = signal<{ code: string; name: string }[]>([])
  protected readonly companies = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly rangeRows = signal<Row[]>([])
  protected readonly rangeIsCheque = signal(false)
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly selectedKey = signal<string | null>(null)
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly editing = signal(false)
  protected readonly amount = computed(() => this.rows().filter((r) => this.checked().has(this.key(r))).reduce((t, r) => t + Number(r['AmountToPay'] ?? 0), 0))
  protected f = this.empty()
  protected pv = this.emptyPv()
  protected range = { pvFrom: '', pvTo: '', chqFrom: '', chqTo: '', account: false }
  protected readonly key = (r: Row) => `${s(r['CASNumber'])}|${s(r['PVMBMIndex'])}`
  protected readonly rangeKey = (r: Row) => [r['PVNo'], r['PVChequeNo'], r['CASNumber']].map(s).join('|')
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'DebitCreditRefNo', label: 'CR No' }, { key: 'Status', label: 'Status' }, { key: 'PatientName', label: 'Patient' }, { key: 'CASNumber', label: 'Case No' },
    money('CLMPayableByMedix2M', 'Payable'), money('Paid', 'Paid'), money('Balance', 'Balance'), money('CLMPostageFee', 'Postage'), { key: 'PVPayto', label: 'Pay To' },
    { key: 'PVNo', label: 'PV No' }, { key: 'PVDate', label: 'PV Date', value: (r) => d(r['PVDate']) }, { key: 'PVChequeNo', label: 'Cheque No' },
    { key: 'PVChequeDate', label: 'Cheque Date', value: (r) => d(r['PVChequeDate']) }, money('PVCLNChequeAmt', 'Cheque Amt'),
    { key: 'period', label: 'A/C Period', value: (r) => (s(r['PAYAccPeriodMth']) ? `${s(r['PAYAccPeriodMth'])}/${s(r['PAYAccPeriodYr'])}` : '') },
    { key: 'PVPayeeBank', label: 'Bank' }, { key: 'MBMBankACNo', label: 'Bank Acc' }, { key: 'Remark', label: 'Remark' }, { key: 'OPBatchNo', label: 'OP Batch' },
    { key: 'bordx', label: 'Bordx', value: (r) => (r['ClaimBordxStatus'] == null ? '' : s(r['ClaimBordxStatus']) === '0' ? 'No' : 'Yes') },
    { key: 'wording', label: 'Wording', value: (r) => `${s(r['PVWordingOne'])} ${s(r['PVWordingTwo'])}`.trim() }, { key: 'MBMName', label: 'Member' },
    money('CLMTotalActual', 'Bill Amt'), { key: 'MBMGroupCompany', label: 'Group Company' }, { key: 'Subsidiary', label: 'Subsidiary' },
    { key: 'MBMGroupCategory', label: 'Category' }, { key: 'MBMnumber', label: 'Member No' }, { key: 'CoverID', label: 'Cov' }, { key: 'BordxRefNo', label: 'Bordx No' },
    { key: 'channel', label: 'Channel', value: (r) => (s(r['BusinessChanel']) ? `${s(r['BusinessChanel'])} - ${s(r['Descr'])}` : '') },
    { key: 'MBMEmployeeNo', label: 'Employee No' }, { key: 'CASTreatReason1', label: 'Treatment 1' }, { key: 'CASTreatReason2', label: 'Treatment 2' },
    { key: 'DateVisit', label: 'Visit Date', value: (r) => d(r['DateVisit']) }, { key: 'BTBGBordxRefNo', label: 'BTBG Bordx' }, { key: 'BTBGBordxStatus', label: 'BTBG Status' },
  ]

  protected get rangeColumns(): GridColumn<Row>[] {
    return this.rangeIsCheque()
      ? [{ key: 'PVChequeNo', label: 'Cheque No' }, { key: 'PVChequeDate', label: 'Cheque Date', value: (r) => d(r['PVChequeDate']) },
          money('PVCLNChequeAmt', 'Amount'), { key: 'PVPayto', label: 'Payee' }, { key: 'PVPayeeACNo', label: 'Account No' }]
      : [{ key: 'PVNo', label: 'PV No' }, { key: 'PVDate', label: 'PV Date', value: (r) => d(r['PVDate']) }, { key: 'MBMName', label: 'Name' },
          { key: 'PVChequeNo', label: 'Cheque No' }, money('PVCLNChequeAmt', 'Cheque Amt'), money('CLMFinalApprove', 'Approved'), { key: 'CASNumber', label: 'Case No' }]
  }

  private empty() {
    return {
      scan: true, caseFrom: '', caseTo: '', mbmNo: '', mbmName: '', company: '', payment: 0, bordx: 1 as number | null, bordxFrom: '', bordxTo: '',
      btbg: null as number | null, btbgFrom: '', btbgTo: '', pvFrom: '', pvTo: '', pvDateFrom: '', pvDateTo: '', chequeFrom: '', chequeTo: '', crFrom: '', crTo: '',
      opFrom: '', opTo: '', monthFrom: '', monthTo: '', yearFrom: '', yearTo: '', blankPv: false, blankPvDate: false, blankCheque: false, blankCr: false,
      blankBatch: false, blankMonth: false, blankYear: false,
    }
  }

  private emptyPv() {
    return {
      index: '', casNumber: '', pvNo: '', pvDate: '', payTo: '', chequeNo: '', chequeDate: '', chequeAmt: '', wording: '', bank: '', month: '', year: '', remark: '',
      void: false,
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ banks: { code: string; name: string }[]; companies: string[] }>(`${BASE}/lookups`))
      this.banks.set(l.banks)
      this.companies.set(l.companies)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdSearch_Click: scan mode appends. */
  protected async search(): Promise<void> {
    if (this.f.scan && this.rows().some((r) => s(r['CASNumber']) === this.f.caseFrom.trim())) return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
    this.editing.set(false)
    try {
      const r = await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, this.f))
      this.setRows(this.f.scan ? [...this.rows(), ...r] : r, this.f.scan)
      if (this.f.scan) this.f.caseFrom = ''
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private setRows(rows: Row[], keep = false): void {
    this.rows.set(rows)
    if (!keep) this.checked.set(new Set())
  }

  /** ProcessCheckedItem: broker channel notice; a paid case cannot be ticked. */
  protected check(keys: Set<string>): void {
    const before = this.checked()
    for (const r of this.rows()) {
      const k = this.key(r)
      if (!keys.has(k) || before.has(k)) continue
      if (r['TickNotice']) this.notify.notify(String(r['TickNotice']), 'info')
      if (r['TickError']) {
        keys.delete(k)
        this.notify.notify(String(r['TickError']), 'error')
      }
    }
    this.editing.set(false)
    this.checked.set(keys)
  }

  protected checkAll(): void {
    this.check(new Set(this.rows().map(this.key)))
  }

  /** lstPaymentMBM_Click. */
  protected select(r: Row): void {
    this.editing.set(false)
    this.selectedKey.set(this.key(r))
    if (!s(r['PVMBMIndex'])) {
      this.pv = this.emptyPv()
      return
    }
    const bank = this.banks().find((b) => b.code === s(r['PVPayeeBank']))
    this.pv = {
      index: s(r['PVMBMIndex']), casNumber: s(r['CASNumber']), pvNo: s(r['PVNo']), pvDate: d(r['PVDate']), payTo: s(r['PVPayto']), chequeNo: s(r['PVChequeNo']),
      chequeDate: d(r['PVChequeDate']), chequeAmt: s(r['PVCLNChequeAmt']), wording: `${s(r['PVWordingOne'])} ${s(r['PVWordingTwo'])}`.trim(),
      bank: bank ? `${bank.code} - ${bank.name}` : s(r['PVPayeeBank']), month: s(r['PAYAccPeriodMth']).padStart(2, '0'), year: s(r['PAYAccPeriodYr']),
      remark: s(r['Remark']), void: false,
    }
  }

  private ticked() {
    return this.rows().filter((r) => this.checked().has(this.key(r))).map((r) => ({ casNumber: s(r['CASNumber']), index: s(r['PVMBMIndex']) || null }))
  }

  /** CmdSave_Click. */
  protected async save(): Promise<void> {
    if (!this.ticked().length) return this.notify.notify('Please select record', 'error')
    if (!this.pv.chequeNo.trim() && !this.pv.void) return this.notify.notify('Please Enter Cheque Number!', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    this.saving.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ pvNo: string; message: string }>(`${BASE}/save`, { ...this.pv, ticked: this.ticked() }))
      this.notify.notify(r.message, 'success')
      this.pv = this.emptyPv()
      this.setRows(await firstValueFrom(this.http.get<Row[]>(`${BASE}/relist`, { params: { pvNo: r.pvNo } })))
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
      const r = await firstValueFrom(this.http.post<{ pvNo: string | null; caseNo: string | null; message: string }>(`${BASE}/update`, { ...this.pv, ticked: [] }))
      const params: Record<string, string> = r.pvNo ? { pvNo: r.pvNo } : { caseNo: r.caseNo ?? '' }
      this.editing.set(false)
      this.pv = this.emptyPv()
      this.setRows(await firstValueFrom(this.http.get<Row[]>(`${BASE}/relist`, { params })))
      this.notify.notify(r.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async preview(): Promise<void> {
    if (!this.ticked().length) return this.notify.notify('Please select the record', 'error')
    await this.download(`${BASE}/voucher`, { cases: this.ticked().map((t) => t.casNumber), payTo: this.pv.payTo, pvNo: this.pv.pvNo }, 'PV_Preview.xlsx')
  }

  /** cmdPrintChq_Click. */
  protected async printCheque(): Promise<void> {
    if (!(await this.notify.ask('Do you want to print now?'))) return
    if (!this.pv.pvNo.trim()) return this.notify.notify('Please Enter or Generate PV No', 'error')
    if (!this.pv.chequeNo.trim()) return this.notify.notify('Please Enter Cheque No', 'error')
    await this.download(`${BASE}/cheques/print`, { from: this.pv.chequeNo.trim(), to: this.pv.chequeNo.trim() }, 'Cheque.xlsx')
  }

  protected async pvRange(): Promise<void> {
    await this.loadRange(`${BASE}/pv-range`, { from: this.range.pvFrom, to: this.range.pvTo }, false)
  }

  protected async printPvRange(): Promise<void> {
    if (!this.rangeRows().length || this.rangeIsCheque()) return this.notify.notify('Please select PV to print', 'error')
    if (!(await this.notify.ask('Do you want to print now?'))) return
    await this.download(`${BASE}/voucher`, { pvFrom: this.range.pvFrom, pvTo: this.range.pvTo }, 'PVRange.xlsx')
  }

  protected async cheques(): Promise<void> {
    await this.loadRange(`${BASE}/cheques?account=${this.range.account}`, { from: this.range.chqFrom, to: this.range.chqTo }, true)
  }

  protected async printCheques(): Promise<void> {
    if (!this.rangeRows().length || !this.rangeIsCheque()) return this.notify.notify('Please select cheque to print', 'error')
    if (!(await this.notify.ask('Do you want to print now?'))) return
    await this.download(`${BASE}/cheques/print?account=${this.range.account}`, { from: this.range.chqFrom, to: this.range.chqTo }, 'Cheques.xlsx')
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
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((r) => this.columns.map((c) => (c.value ? c.value(r) : s(r[c.key]))))], 'PaymentMember.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.pv = this.emptyPv()
    this.editing.set(false)
    this.rangeRows.set([])
    this.setRows([])
  }
}
