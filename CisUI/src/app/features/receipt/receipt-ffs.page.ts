import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
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

const BASE = '/api/receipt/ffs'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })

/** Receipt > Fr FFS (FrmReceiptFFS). */
@Component({
  selector: 'app-receipt-ffs-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Receipt from FFS" subtitle="Replaces FrmReceiptFFS">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Scan"><label class="check-field"><input type="checkbox" name="scan" [(ngModel)]="f.scan" /> One bordx no.</label></app-field>
          <app-field [label]="f.scan ? 'Bordx No' : 'Bordx No From'"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" (keydown.enter)="search()" /></app-field>
          @if (!f.scan) {
            <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
          }
          <app-field label="Group Company" [wide]="true">
            <input name="company" list="rf-companies" [(ngModel)]="f.company" />
            <datalist id="rf-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-field label="Payment">
            <select name="payment" [(ngModel)]="f.payment"><option [ngValue]="0">0 - All</option><option [ngValue]="1">1 - Paid</option><option [ngValue]="2">2 - Unpaid</option></select>
          </app-field>
          <app-field label="Bordx Date From"><app-date-input name="bordxDateFrom" ariaLabel="Bordx Date From" [(ngModel)]="f.bordxDateFrom" /></app-field>
          <app-field label="Bordx Date To"><app-date-input name="bordxDateTo" ariaLabel="Bordx Date To" [(ngModel)]="f.bordxDateTo" /></app-field>
          <app-field label="Receipt No From"><input name="receiptFrom" maxlength="30" [(ngModel)]="f.receiptFrom" /></app-field>
          <app-field label="Receipt No To"><input name="receiptTo" maxlength="30" [(ngModel)]="f.receiptTo" /></app-field>
          <app-field label="Receipt Date From"><app-date-input name="receiptDateFrom" ariaLabel="Receipt Date From" [(ngModel)]="f.receiptDateFrom" /></app-field>
          <app-field label="Receipt Date To"><app-date-input name="receiptDateTo" ariaLabel="Receipt Date To" [(ngModel)]="f.receiptDateTo" /></app-field>
          <app-field label="A/C Month From"><select name="monthFrom" [(ngModel)]="f.monthFrom"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Month To"><select name="monthTo" [(ngModel)]="f.monthTo"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Year From"><input name="yearFrom" maxlength="4" [(ngModel)]="f.yearFrom" /></app-field>
          <app-field label="A/C Year To"><input name="yearTo" maxlength="4" [(ngModel)]="f.yearTo" /></app-field>
          <app-action-bar label="FFS receipt search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="checkAll()">Check All</button>
            <button appButton (click)="add()">Add</button>
            <button appButton (click)="export()">Export</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Receipt">
          <legend>Receipt</legend>
          <app-field label="Receipt No" [required]="true"><input name="receiptNo" maxlength="30" [(ngModel)]="r.receiptNo" /></app-field>
          <app-field label="Receipt Date" [required]="true"><app-date-input name="receiptDate" ariaLabel="Receipt Date" [(ngModel)]="r.receiptDate" /></app-field>
          <app-field label="Receipt Amt" [required]="true"><input class="num" name="receiptAmt" maxlength="20" [(ngModel)]="r.receiptAmt" /></app-field>
          <app-field label="Cheque No" [required]="true"><input name="chequeNo" maxlength="30" [(ngModel)]="r.chequeNo" /></app-field>
          <app-field label="A/C Period Month" [required]="true"><select name="month" [(ngModel)]="r.month"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
          <app-field label="A/C Period Year" [required]="true"><input name="year" maxlength="4" [(ngModel)]="r.year" /></app-field>
          <app-field label="Remark" [wide]="true"><input name="remark" maxlength="500" [(ngModel)]="r.remark" /></app-field>
          <app-action-bar label="Receipt actions">
            @if (editing()) {
              <button appButton variant="primary" (click)="update()">Update</button>
            } @else {
              <button appButton variant="primary" [busy]="saving()" (click)="save()">Save</button>
              <button appButton [disabled]="!r.index" (click)="editing.set(true)">Edit</button>
            }
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="FFS receipts" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()" (checkedChange)="check($event)"
          [selectable]="true" [selectedKey]="selectedKey()" (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">
          Records: <strong>{{ rows().length }}</strong> · Total FFS: <strong>{{ money(total()) }}</strong> · Checked: <strong>{{ money(amount()) }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class ReceiptFfsPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'))
  protected readonly companies = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly selectedKey = signal<string | null>(null)
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly editing = signal(false)
  protected readonly total = computed(() => this.rows().reduce((t, x) => t + Number(x['CLMMCOProcFee'] ?? 0), 0))
  protected readonly amount = computed(() => this.rows().filter((x) => this.checked().has(this.key(x))).reduce((t, x) => t + Number(x['CLMMCOProcFee'] ?? 0), 0))
  protected f = this.empty()
  protected r = this.emptyReceipt()
  protected readonly key = (x: Row) => `${s(x['BordxRefNo'])}|${s(x['ReceiptIndex'])}`
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'Status', label: 'Status' }, { key: 'BordxRefNo', label: 'Bordx No' }, { key: 'BordxDate', label: 'Bordx Date', value: (x) => d(x['BordxDate']) },
    money('CLMBordxAmt', 'Bordx Amt'), money('CLMMCOProcFee', 'FFS'), money('ReceiptAmt', 'Received'), money('Balance', 'Balance'), { key: 'ReceiptNo', label: 'Receipt No' },
    { key: 'ReceiptDate', label: 'Receipt Date', value: (x) => d(x['ReceiptDate']) }, { key: 'ReceiptRemarks', label: 'Remark' },
    { key: 'period', label: 'A/C Period', value: (x) => (s(x['ReceiptAccMth']) ? `${s(x['ReceiptAccMth'])}/${s(x['ReceiptAccYr'])}` : '') },
    { key: 'ReceiptChqNo', label: 'Cheque No' }, money('TotalRecAmt', 'Total Received'), { key: 'GroupCompany', label: 'Group Company' },
  ]

  private empty() {
    return {
      scan: true, bordxFrom: '', bordxTo: '', company: '', payment: 0, bordxDateFrom: '', bordxDateTo: '', receiptFrom: '', receiptTo: '', receiptDateFrom: '',
      receiptDateTo: '', monthFrom: '', monthTo: '', yearFrom: '', yearTo: '',
    }
  }

  private emptyReceipt() {
    return { index: '', bordxNo: '', receiptNo: '', receiptDate: '', receiptAmt: '', chequeNo: '', month: '', year: '', remark: '' }
  }

  async ngOnInit(): Promise<void> {
    try {
      this.companies.set((await firstValueFrom(this.http.get<{ companies: string[] }>(`${BASE}/lookups`))).companies)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    if (this.f.scan && this.rows().some((x) => s(x['BordxRefNo']) === this.f.bordxFrom.trim())) return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
    try {
      const res = await firstValueFrom(this.http.post<{ rows: Row[] }>(`${BASE}/search`, this.f))
      this.rows.set(this.f.scan ? [...this.rows(), ...res.rows] : res.rows)
      if (!this.f.scan) this.checked.set(new Set())
      if (this.f.scan) this.f.bordxFrom = ''
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

  /** cmdCheckAll_Click: toggles every row. */
  protected checkAll(): void {
    if (!this.rows().length) return this.notify.notify('Please select a record.', 'info')
    const all = this.rows().every((x) => this.checked().has(this.key(x)) || x['TickError'])
    this.check(all ? new Set() : new Set(this.rows().map(this.key)))
  }

  /** CmdAdd_Click. */
  protected add(): void {
    if (!this.checked().size) return this.notify.notify('Please checked at least 1 Claim  ! ', 'error')
    this.editing.set(false)
    this.r = { ...this.emptyReceipt(), receiptAmt: this.amount().toFixed(2) }
  }

  /** LstFFSRcpt_Click. */
  protected select(x: Row): void {
    this.selectedKey.set(this.key(x))
    this.editing.set(false)
    if (!s(x['ReceiptIndex'])) {
      this.r = this.emptyReceipt()
      return
    }
    this.r = {
      index: s(x['ReceiptIndex']), bordxNo: s(x['BordxRefNo']), receiptNo: s(x['ReceiptNo']), receiptDate: d(x['ReceiptDate']), receiptAmt: s(x['CLMMCOProcFee']),
      chequeNo: s(x['ReceiptChqNo']), month: s(x['ReceiptAccMth']).padStart(2, '0'), year: s(x['ReceiptAccYr']), remark: s(x['ReceiptRemarks']),
    }
  }

  protected async save(): Promise<void> {
    if (!this.checked().size) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record ? '))) return
    this.saving.set(true)
    try {
      const ticked = this.rows().filter((x) => this.checked().has(this.key(x))).map((x) => s(x['BordxRefNo']))
      const res = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/save`, { ...this.r, ticked }))
      await this.relist(res.message)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  protected async update(): Promise<void> {
    if (!(await this.notify.ask('Do you want to update the record ? '))) return
    try {
      const res = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/update`, { ...this.r, ticked: [] }))
      await this.relist(res.message)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private async relist(message: string): Promise<void> {
    const receiptNo = this.r.receiptNo.trim()
    this.editing.set(false)
    this.r = this.emptyReceipt()
    this.checked.set(new Set())
    this.rows.set((await firstValueFrom(this.http.get<{ rows: Row[] }>(`${BASE}/receipt/${encodeURIComponent(receiptNo)}`))).rows)
    this.notify.notify(message, 'success')
  }

  protected export(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'ReceiptFFS.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.r = this.emptyReceipt()
    this.editing.set(false)
    this.rows.set([])
    this.checked.set(new Set())
  }
}
