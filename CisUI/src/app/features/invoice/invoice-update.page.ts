import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { parseDmy } from '../../shared/lib/masks'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/invoice/update'
type Row = Record<string, unknown>

const money = (v: unknown) => (typeof v === 'number' ? formatMoney(v) : '0.00')

/** Invoice > Other > Invoice Update (frmUpdateCLNInvoice): find cases and stamp the checked ones with an invoice number and date. */
@Component({
  selector: 'app-invoice-update-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Invoice Update" subtitle="Replaces frmUpdateCLNInvoice">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="WS No. From"><input name="casFrom" data-autofocus [(ngModel)]="f.casFrom" /></app-field>
          <app-field label="WS No. To"><input name="casTo" [(ngModel)]="f.casTo" /></app-field>
          <app-field label="Clinic">
            <input name="clinic" list="iu-clinics" [(ngModel)]="f.clinic" />
            <datalist id="iu-clinics">
              @for (c of clinics(); track c.code) {
                <option [value]="c.code + ' - ' + c.name"></option>
              }
            </datalist>
          </app-field>
          <app-field label="Payor">
            <input name="groupCompany" list="iu-groups" [(ngModel)]="f.groupCompany" />
            <datalist id="iu-groups">
              @for (g of groups(); track $index) {
                <option [value]="g"></option>
              }
            </datalist>
          </app-field>
          <app-field label="Invoice No."><input name="invoiceNo" [(ngModel)]="f.invoiceNo" /></app-field>
          <label class="check-field"><input type="checkbox" name="noInvoice" [(ngModel)]="f.noInvoice" /> Without Invoice No.</label>
          <app-field label="Visit Date From"><app-date-input name="visitFrom" ariaLabel="Visit Date From" [(ngModel)]="f.visitFrom" (focusout)="checkDate(f.visitFrom)" /></app-field>
          <app-field label="Visit Date To"><app-date-input name="visitTo" ariaLabel="Visit Date To" [(ngModel)]="f.visitTo" (focusout)="checkDate(f.visitTo)" /></app-field>
          <app-field label="Invoice Date From"><app-date-input name="invoiceFrom" ariaLabel="Invoice Date From" [(ngModel)]="f.invoiceFrom" (focusout)="checkDate(f.invoiceFrom)" /></app-field>
          <app-field label="Invoice Date To"><app-date-input name="invoiceTo" ariaLabel="Invoice Date To" [(ngModel)]="f.invoiceTo" (focusout)="checkDate(f.invoiceTo)" /></app-field>
          <app-field label="Invoice Amt From"><input class="num" name="amountFrom" [(ngModel)]="f.amountFrom" /></app-field>
          <app-field label="Invoice Amt To"><input class="num" name="amountTo" [(ngModel)]="f.amountTo" /></app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Cases" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="checked.set($event)" />
        <p class="totals">Records: <strong>{{ rows().length }}</strong> · Checked: <strong>{{ checked().size }}</strong></p>

        <fieldset aria-label="Update">
          <legend>Update</legend>
          <app-field label="Invoice No." [required]="true"><input name="newInvoiceNo" maxlength="50" [(ngModel)]="invoiceNo" /></app-field>
          <app-field label="Invoice Date" [required]="true">
            <app-date-input name="newInvoiceDate" ariaLabel="Invoice Date" [(ngModel)]="invoiceDate" (focusout)="checkDate(invoiceDate)" />
          </app-field>
          <app-action-bar label="Update actions">
            <button appButton (click)="checkAll()">Check</button>
            <button appButton variant="primary" [busy]="busy()" (click)="update()">Update</button>
          </app-action-bar>
        </fieldset>
      </div>
    </app-page-shell>
  `,
})
export class InvoiceUpdatePage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly clinics = signal<LookupOption[]>([])
  protected readonly groups = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly busy = signal(false)
  protected readonly key = (r: Row) => String(r['CASNumber'])
  protected f = this.emptyFilters()
  protected invoiceNo = ''
  protected invoiceDate = ''

  protected readonly columns: GridColumn<Row>[] = [
    { key: 'CASNumber', label: 'Case No' }, { key: 'MBMNumber', label: 'Member No' }, { key: 'CoverID', label: 'Cov' },
    { key: 'PatientName', label: 'Patient Name' }, { key: 'DateVisit', label: 'Visit Date', value: (r) => formatDate(r['DateVisit']) },
    { key: 'CLNName', label: 'Clinic' }, { key: 'BillNo', label: 'Invoice No' }, { key: 'BillDate', label: 'Invoice Date', value: (r) => formatDate(r['BillDate']) },
    { key: 'INVEnteredDate', label: 'Entered Date', value: (r) => formatDate(r['INVEnteredDate']) }, { key: 'INVEnteredBy', label: 'Entered By' },
    { key: 'INVAmount', label: 'Invoice Amt', align: 'right', value: (r) => money(r['INVAmount']) },
    { key: 'CLMMCOProcFee', label: 'FFS', align: 'right', value: (r) => money(r['CLMMCOProcFee']) },
    { key: 'CLMFinalApprove', label: 'Final Approve', align: 'right', value: (r) => money(r['CLMFinalApprove']) },
    { key: 'MBMGroupCompany', label: 'Group Company' },
  ]

  async ngOnInit(): Promise<void> {
    try {
      const s = await firstValueFrom(this.http.get<{ clinics: LookupOption[]; groupCompanies: string[] }>(`${BASE}/setup`))
      this.clinics.set(s.clinics)
      this.groups.set(s.groupCompanies)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected checkDate(v: string): void {
    if (v && !parseDmy(v)) this.notify.notify('Invalid Date Format! ', 'error')
  }

  /** CmdSearch_Click. */
  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.checked.set(new Set())
    try {
      this.rows.set(await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, this.f)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** CmdPrint_Click: the list as Excel. */
  protected async print(): Promise<void> {
    if (this.rows().length === 0) return this.notify.notify('Report cannot be printed without any record.', 'info')
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/export`, this.f, { responseType: 'blob' })), 'InvoiceUpdate.xlsx')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  /** CmdCheck_Click: toggles every row. */
  protected checkAll(): void {
    if (this.rows().length === 0) return this.notify.notify('Please select a record.', 'info')
    const now = this.checked()
    this.checked.set(new Set(this.rows().map(this.key).filter((k) => !now.has(k))))
  }

  /** CmdUpdate_Click. */
  protected async update(): Promise<void> {
    if (this.checked().size === 0) return this.notify.notify('Please select the record', 'error')
    if (!parseDmy(this.invoiceDate)) return this.notify.notify('Invalid Invoice Date!', 'error')
    if (!this.invoiceNo) return this.notify.notify('Please Enter Cheque Amt!', 'error')
    if (!(await this.notify.ask('Do you want to update the record ? '))) return
    this.busy.set(true)
    try {
      const body = { casNumbers: [...this.checked()], invoiceNo: this.invoiceNo, invoiceDate: this.invoiceDate }
      const r = await firstValueFrom(this.http.post<{ message: string; rows: Row[] }>(BASE, body))
      this.rows.set(r.rows)
      this.checked.set(new Set())
      this.notify.notify(r.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected clear(): void {
    this.f = this.emptyFilters()
    this.rows.set([])
    this.checked.set(new Set())
    this.invoiceNo = ''
    this.invoiceDate = ''
  }

  private emptyFilters() {
    return {
      casFrom: '', casTo: '', clinic: '', groupCompany: '', invoiceNo: '', visitFrom: '', visitTo: '', noInvoice: false, invoiceFrom: '', invoiceTo: '',
      amountFrom: '', amountTo: '',
    }
  }
}
