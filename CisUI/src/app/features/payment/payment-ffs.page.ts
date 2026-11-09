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

const BASE = '/api/payment/ffs'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })

/** Payment > FFS Payment (FrmPaymentFFS). */
@Component({
  selector: 'app-payment-ffs-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="FFS Payment" subtitle="Replaces FrmPaymentFFS">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Scan"><label class="check-field"><input type="checkbox" name="scan" [(ngModel)]="f.scan" /> One bordx no.</label></app-field>
          <app-field [label]="f.scan ? 'Bordx No' : 'Bordx No From'"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" (keydown.enter)="search()" /></app-field>
          @if (!f.scan) {
            <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
          }
          <app-field label="Group Company" [wide]="true">
            <input name="company" list="pf-companies" [(ngModel)]="f.company" />
            <datalist id="pf-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-field label="Payment">
            <select name="payment" [(ngModel)]="f.payment"><option [ngValue]="0">0 - All</option><option [ngValue]="1">1 - Paid</option><option [ngValue]="2">2 - Unpaid</option></select>
          </app-field>
          <app-field label="Bordx Date From"><app-date-input name="bordxDateFrom" ariaLabel="Bordx Date From" [(ngModel)]="f.bordxDateFrom" /></app-field>
          <app-field label="Bordx Date To"><app-date-input name="bordxDateTo" ariaLabel="Bordx Date To" [(ngModel)]="f.bordxDateTo" /></app-field>
          <app-field label="PV No From"><input name="pvFrom" maxlength="30" [(ngModel)]="f.pvFrom" /></app-field>
          <app-field label="PV No To"><input name="pvTo" maxlength="30" [(ngModel)]="f.pvTo" /></app-field>
          <app-field label="PV Date From"><app-date-input name="pvDateFrom" ariaLabel="PV Date From" [(ngModel)]="f.pvDateFrom" /></app-field>
          <app-field label="PV Date To"><app-date-input name="pvDateTo" ariaLabel="PV Date To" [(ngModel)]="f.pvDateTo" /></app-field>
          <app-field label="A/C Month From"><select name="monthFrom" [(ngModel)]="f.monthFrom"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Month To"><select name="monthTo" [(ngModel)]="f.monthTo"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Year From"><input name="yearFrom" maxlength="4" [(ngModel)]="f.yearFrom" /></app-field>
          <app-field label="A/C Year To"><input name="yearTo" maxlength="4" [(ngModel)]="f.yearTo" /></app-field>
          <app-field label="Empty" [wide]="true">
            <span class="radio-row">
              <label class="check-field"><input type="checkbox" name="blankBordxDate" [(ngModel)]="f.blankBordxDate" /> Bordx Date</label>
              <label class="check-field"><input type="checkbox" name="blankPv" [(ngModel)]="f.blankPv" /> PV No</label>
              <label class="check-field"><input type="checkbox" name="blankPvDate" [(ngModel)]="f.blankPvDate" /> PV Date</label>
              <label class="check-field"><input type="checkbox" name="blankMonth" [(ngModel)]="f.blankMonth" /> A/C Month</label>
              <label class="check-field"><input type="checkbox" name="blankYear" [(ngModel)]="f.blankYear" /> A/C Year</label>
            </span>
          </app-field>
          <app-action-bar label="FFS search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="checkAll()">Check All</button>
            <button appButton (click)="add()">Add</button>
            <button appButton (click)="print()">Print to File</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Payment voucher">
          <legend>Payment Voucher</legend>
          <app-field label="PV No"><input name="pvNo" [(ngModel)]="pv.pvNo" disabled /></app-field>
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
          <app-field label="Payee" [required]="true">
            <input name="payee" list="pf-payees" [(ngModel)]="pv.payee" />
            <datalist id="pf-payees">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-field label="Remark" [wide]="true"><input name="remark" maxlength="500" [(ngModel)]="pv.remark" /></app-field>
          <app-action-bar label="Payment voucher actions">
            @if (editing()) {
              <button appButton variant="primary" (click)="update()">Update</button>
            } @else {
              <button appButton variant="primary" [busy]="saving()" (click)="save()">Save</button>
              <button appButton [disabled]="!pv.index || checked().size > 0" (click)="editing.set(true)">Edit</button>
            }
            <button appButton (click)="voucher(false)">Preview PV</button>
            <button appButton (click)="voucher(true)">Print PV</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="FFS payments" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="check($event)" [selectable]="true" [selectedKey]="selectedKey()" (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">
          Records: <strong>{{ rows().length }}</strong> · Total FFS: <strong>{{ money(totalFfs()) }}</strong> · Amount to pay: <strong>{{ money(amount()) }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class PaymentFfsPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'))
  protected readonly banks = signal<{ code: string; name: string }[]>([])
  protected readonly companies = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly totalFfs = signal(0)
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly selectedKey = signal<string | null>(null)
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly editing = signal(false)
  protected readonly amount = computed(() => this.rows().filter((r) => this.checked().has(this.key(r))).reduce((t, r) => t + Number(r['Balance'] ?? 0), 0))
  protected f = this.empty()
  protected pv = this.emptyPv()
  protected readonly key = (r: Row) => `${s(r['BordxRefNo'])}|${s(r['PaymentIndex'])}`
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'Status', label: 'Status' }, { key: 'BordxRefNo', label: 'Bordx No' }, { key: 'BordxDate', label: 'Bordx Date', value: (r) => d(r['BordxDate']) },
    money('CLMBordxAmt', 'Bordx Amt'), money('CLMMCOProcFee', 'FFS'), money('FFSPaidAmt', 'Paid'), money('Balance', 'Balance'),
    { key: 'GroupCompany', label: 'Group Company' }, { key: 'PVNo', label: 'PV No' }, { key: 'PVDate', label: 'PV Date', value: (r) => d(r['PVDate']) },
    { key: 'PVChqNo', label: 'Cheque No' }, { key: 'PVChqDate', label: 'Cheque Date', value: (r) => d(r['PVChqDate']) }, money('PVChqAmt', 'Cheque Amt'),
    { key: 'period', label: 'A/C Period', value: (r) => (s(r['FFSAccMth']) ? `${s(r['FFSAccMth'])}/${s(r['FFSAccYr'])}` : '') }, { key: 'PVBank', label: 'Bank' },
    { key: 'PVAcc', label: 'A/C No' }, { key: 'Remark', label: 'Remark' }, { key: 'GRPCoName', label: 'Payee' }, { key: 'PVRecNo', label: 'Receipt No' },
    { key: 'PVRecDate', label: 'Receipt Date', value: (r) => d(r['PVRecDate']) },
  ]

  private empty() {
    return {
      scan: true, bordxFrom: '', bordxTo: '', company: '', payment: 0, bordxDateFrom: '', bordxDateTo: '', pvFrom: '', pvTo: '', pvDateFrom: '', pvDateTo: '',
      monthFrom: '', monthTo: '', yearFrom: '', yearTo: '', blankBordxDate: false, blankPv: false, blankPvDate: false, blankMonth: false, blankYear: false,
    }
  }

  private emptyPv() {
    return {
      index: '', bordxNo: '', pvNo: '', pvDate: '', chequeNo: '', chequeDate: '', chequeAmt: '', bank: 'MB - MAYBANK', account: '', month: '', year: '',
      payee: 'HEALTH CONNECT SDN BHD', remark: '',
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ banks: { code: string; name: string }[]; companies: string[] }>(`${BASE}/lookups`))
      this.banks.set(l.banks)
      this.companies.set(l.companies)
      const mb = l.banks.find((b) => b.code === 'MB')
      if (mb) this.pv.bank = `${mb.code} - ${mb.name}`
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CmdSearch_Click: scan mode appends. */
  protected async search(): Promise<void> {
    if (this.f.scan && this.rows().some((r) => s(r['BordxRefNo']) === this.f.bordxFrom.trim())) return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ rows: Row[]; totalFfs: number }>(`${BASE}/search`, this.f))
      this.setRows(this.f.scan ? [...this.rows(), ...r.rows] : r.rows, this.f.scan)
      if (this.f.scan) this.f.bordxFrom = ''
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private setRows(rows: Row[], keep = false): void {
    this.rows.set(rows)
    this.totalFfs.set(rows.reduce((t, r) => t + Number(r['CLMMCOProcFee'] ?? 0), 0))
    if (!keep) this.checked.set(new Set())
  }

  protected check(keys: Set<string>): void {
    const before = this.checked()
    for (const r of this.rows()) {
      const k = this.key(r)
      if (keys.has(k) && !before.has(k) && r['TickError']) {
        keys.delete(k)
        this.notify.notify(String(r['TickError']), 'error')
      }
    }
    this.editing.set(false)
    this.checked.set(keys)
  }

  /** cmdCheckAll_Click: toggles every row. */
  protected checkAll(): void {
    if (!this.rows().length) return this.notify.notify('Please select a record.', 'info')
    const all = this.rows().every((r) => this.checked().has(this.key(r)) || r['TickError'])
    this.check(all ? new Set() : new Set(this.rows().map(this.key)))
  }

  /** CmdAdd_Click: a new PV for the ticked rows with their total. */
  protected add(): void {
    if (!this.checked().size) return this.notify.notify('Please checked at least 1 Claim  ! ', 'error')
    this.editing.set(false)
    this.pv = { ...this.emptyPv(), bank: this.pv.bank, chequeAmt: this.amount().toFixed(2) }
  }

  protected select(r: Row): void {
    this.editing.set(false)
    this.selectedKey.set(this.key(r))
    if (!s(r['PVNo'])) return
    const bank = this.banks().find((b) => b.code === s(r['PVBank']))
    this.pv = {
      index: s(r['PaymentIndex']), bordxNo: s(r['BordxRefNo']), pvNo: s(r['PVNo']), pvDate: d(r['PVDate']), chequeNo: s(r['PVChqNo']), chequeDate: d(r['PVChqDate']),
      chequeAmt: s(r['PVChqAmt']), bank: bank ? `${bank.code} - ${bank.name}` : s(r['PVBank']), account: s(r['PVAcc']), month: s(r['FFSAccMth']).padStart(2, '0'),
      year: s(r['FFSAccYr']), payee: s(r['GRPCoName']), remark: s(r['Remark']),
    }
  }

  private ticked(): string[] {
    return this.rows().filter((r) => this.checked().has(this.key(r))).map((r) => s(r['BordxRefNo']))
  }

  protected async save(): Promise<void> {
    if (!this.ticked().length) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record ? '))) return
    this.saving.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ pvNo: string; message: string }>(`${BASE}/save`, { ...this.pv, ticked: this.ticked() }))
      this.notify.notify(r.message, 'success')
      await this.showPv(r.pvNo)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  protected async update(): Promise<void> {
    if (!this.rows().length) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    try {
      await firstValueFrom(this.http.post(`${BASE}/update`, { ...this.pv, ticked: [] }))
      await this.showPv(this.pv.pvNo)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private async showPv(pvNo: string): Promise<void> {
    this.editing.set(false)
    const r = await firstValueFrom(this.http.get<{ rows: Row[] }>(`${BASE}/pv/${encodeURIComponent(pvNo)}`))
    this.pv = { ...this.emptyPv(), pvNo }
    this.setRows(r.rows)
  }

  protected async voucher(print: boolean): Promise<void> {
    if (print && !this.pv.pvNo.trim()) return this.notify.notify('Please Enter or Generate PV No', 'error')
    if (!print && !this.ticked().length) return this.notify.notify('Please select the record', 'error')
    try {
      const body = print ? { pvNo: this.pv.pvNo.trim(), ticked: [] } : { pvNo: null, ticked: this.ticked() }
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/voucher`, body, { responseType: 'blob' })), 'PV_FFS.xlsx')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((r) => this.columns.map((c) => (c.value ? c.value(r) : s(r[c.key]))))], 'PaymentFFS.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.pv = this.emptyPv()
    this.editing.set(false)
    this.setRows([])
  }
}
