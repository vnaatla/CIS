import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core'
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

const BASE = '/api/receipt/member'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })

/** Receipt > Fr Member (FrmReceiptMember). */
@Component({
  selector: 'app-receipt-member-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Receipt from Member" subtitle="Replaces FrmReceiptMember">
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
          <app-field label="Payment">
            <select name="payment" [(ngModel)]="f.payment">@for (p of payments; track $index) { <option [ngValue]="$index">{{ $index }} - {{ p }}</option> }</select>
          </app-field>
          <app-field label="Receipt No"><input name="receiptFrom" maxlength="30" [(ngModel)]="f.receiptFrom" /></app-field>
          <app-field label="Receipt Date From"><app-date-input name="receiptDateFrom" ariaLabel="Receipt Date From" [(ngModel)]="f.receiptDateFrom" /></app-field>
          <app-field label="Receipt Date To"><app-date-input name="receiptDateTo" ariaLabel="Receipt Date To" [(ngModel)]="f.receiptDateTo" /></app-field>
          <app-field label="DN No From"><input name="drFrom" maxlength="30" [(ngModel)]="f.drFrom" /></app-field>
          <app-field label="DN No To"><input name="drTo" maxlength="30" [(ngModel)]="f.drTo" /></app-field>
          <app-field label="DN Amount From"><input class="num" name="amountFrom" maxlength="20" [(ngModel)]="f.amountFrom" /></app-field>
          <app-field label="DN Amount To"><input class="num" name="amountTo" maxlength="20" [(ngModel)]="f.amountTo" /></app-field>
          <app-field label="A/C Month From"><input name="monthFrom" maxlength="2" [(ngModel)]="f.monthFrom" /></app-field>
          <app-field label="A/C Month To"><input name="monthTo" maxlength="2" [(ngModel)]="f.monthTo" /></app-field>
          <app-field label="A/C Year From"><input name="yearFrom" maxlength="4" [(ngModel)]="f.yearFrom" /></app-field>
          <app-field label="A/C Year To"><input name="yearTo" maxlength="4" [(ngModel)]="f.yearTo" /></app-field>
          <app-field label="Empty" [wide]="true">
            <span class="radio-row">
              <label class="check-field"><input type="checkbox" name="blankReceipt" [(ngModel)]="f.blankReceipt" /> Receipt No</label>
              <label class="check-field"><input type="checkbox" name="blankDr" [(ngModel)]="f.blankDr" /> DN No</label>
              <label class="check-field"><input type="checkbox" name="blankAmount" [(ngModel)]="f.blankAmount" /> DN Amount</label>
              <label class="check-field"><input type="checkbox" name="blankMonth" [(ngModel)]="f.blankMonth" /> A/C Month</label>
              <label class="check-field"><input type="checkbox" name="blankYear" [(ngModel)]="f.blankYear" /> A/C Year</label>
            </span>
          </app-field>
          <app-action-bar label="Receipt search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="export()">Export</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Receipt">
          <legend>Receipt</legend>
          <app-field label="Receipt No" [required]="true"><input name="receiptNo" maxlength="30" [(ngModel)]="r.receiptNo" /></app-field>
          <app-field label="Receipt Date" [required]="true"><app-date-input name="receiptDate" ariaLabel="Receipt Date" [(ngModel)]="r.receiptDate" /></app-field>
          <app-field label="Amount" [required]="true"><input class="num" name="amount" maxlength="20" [(ngModel)]="r.amount" /></app-field>
          <app-field label="Cheque No" [required]="true"><input name="chequeNo" maxlength="30" [(ngModel)]="r.chequeNo" /></app-field>
          <app-field label="A/C Period Month"><input name="month" maxlength="2" [(ngModel)]="r.month" /></app-field>
          <app-field label="A/C Period Year"><input name="year" maxlength="4" [(ngModel)]="r.year" /></app-field>
          <app-field label="Remarks" [wide]="true"><input name="remarks" maxlength="500" [(ngModel)]="r.remarks" /></app-field>
          <app-action-bar label="Receipt actions">
            @if (editing()) {
              <button appButton variant="primary" (click)="update()">Update</button>
            } @else {
              <button appButton variant="primary" [busy]="saving()" (click)="save()">Save</button>
              <button appButton (click)="editing.set(true)">Edit</button>
            }
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Member receipts" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="check($event)" [selectable]="true" [selectedKey]="selectedKey()" (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong> · Amount to receive: <strong>{{ money(amount()) }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class ReceiptMemberPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly payments = ['All', 'Paid', 'Unpaid', 'Partial Paid', 'Paid & Partial Paid', 'Unpaid & Partial Paid']
  protected readonly rows = signal<Row[]>([])
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly selectedKey = signal<string | null>(null)
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly editing = signal(false)
  protected readonly amount = computed(() => this.rows().filter((x) => this.checked().has(this.key(x))).reduce((t, x) => t + Number(x['AmountToReceive'] ?? 0), 0))
  protected f = this.empty()
  protected r = this.emptyReceipt()
  protected readonly key = (x: Row) => `${s(x['CASNumber'])}|${s(x['ReceiptIndex'])}`
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'DebitCreditRefNo', label: 'DN No' }, { key: 'Status', label: 'Status' }, { key: 'CASNumber', label: 'Case No' }, money('Amount', 'Amount'),
    money('Received', 'Received'), money('Balance', 'Balance'), { key: 'ReceiptNo', label: 'Receipt No' },
    { key: 'ReceiptDate', label: 'Receipt Date', value: (x) => d(x['ReceiptDate']) },
    { key: 'period', label: 'A/C Period', value: (x) => (s(x['RECMBMAccPeriodMth']) || s(x['RECMBMAccPeriodYr']) ? `${s(x['RECMBMAccPeriodMth'])}/${s(x['RECMBMAccPeriodYr'])}` : '') },
    { key: 'MBMNumber', label: 'Member No' }, { key: 'MBMName', label: 'Member' },
  ]

  private empty() {
    return {
      scan: true, caseFrom: '', caseTo: '', mbmNo: '', mbmName: '', payment: 0, receiptFrom: '', blankReceipt: false, receiptDateFrom: '', receiptDateTo: '',
      drFrom: '', drTo: '', blankDr: false, amountFrom: '', amountTo: '', blankAmount: false, monthFrom: '', monthTo: '', blankMonth: false, yearFrom: '', yearTo: '',
      blankYear: false,
    }
  }

  private emptyReceipt() {
    return { index: '', receiptNo: '', receiptDate: '', amount: '', chequeNo: '', remarks: '', month: '', year: '' }
  }

  /** cmdSearch_Click: scan mode appends. */
  protected async search(): Promise<void> {
    if (this.f.scan && this.rows().some((x) => s(x['CASNumber']) === this.f.caseFrom.trim())) return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
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
    this.checked.set(keys)
  }

  /** lstReceiptMBM_Click: in edit mode the row's receipt fills the frame. */
  protected async select(x: Row): Promise<void> {
    this.selectedKey.set(this.key(x))
    if (!this.editing()) return
    try {
      const rc = await firstValueFrom(this.http.get<Row>(`${BASE}/receipt`, { params: { index: s(x['ReceiptIndex']) } }))
      this.r = {
        index: s(rc['receiptIndex'] ?? rc['ReceiptIndex']), receiptNo: s(rc['ReceiptNo']), receiptDate: d(rc['ReceiptDate']), amount: s(rc['ReceiptAmt']),
        chequeNo: s(rc['ReceiptChequeNo']), remarks: s(rc['ReceiptRemarks']), month: s(rc['RECMBMAccPeriodMth']), year: s(rc['RECMBMAccPeriodYr']),
      }
    } catch (e) {
      this.r = this.emptyReceipt()
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdSave_Click. */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    this.saving.set(true)
    try {
      const ticked = this.rows().filter((x) => this.checked().has(this.key(x))).map((x) => s(x['CASNumber']))
      const res = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/save`, { ...this.r, ticked }))
      this.notify.notify(res.message, 'success')
      this.checked.set(new Set())
      if (!this.f.scan) await this.search()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  /** cmdUpdate_Click. */
  protected async update(): Promise<void> {
    if (!this.rows().length) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    try {
      const res = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/update`, { ...this.r, ticked: [] }))
      this.notify.notify(res.message, 'success')
      this.editing.set(false)
      this.r = this.emptyReceipt()
      if (!this.f.scan) await this.search()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected export(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'ReceiptMember.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.r = this.emptyReceipt()
    this.editing.set(false)
    this.rows.set([])
    this.checked.set(new Set())
  }
}
