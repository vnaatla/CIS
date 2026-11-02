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

const BASE = '/api/invoice/op-batch'
type Row = Record<string, unknown>
interface OpBatchList {
  rows: Row[]
  totalInvoice: number
  totalFfs: number
  multipleMonths: boolean
}

const money = (v: unknown) => (typeof v === 'number' ? formatMoney(v) : '0.00')
const text = (r: Row, k: string) => String(r[k] ?? '').trim()

/** lstInvoiceListing_ItemCheck: why a case cannot be checked. */
function checkMessage(r: Row): string | null {
  if (text(r, 'OPBatchNo') && text(r, 'BillNo')) return 'Selected Case Already OP Batch And Invoice No generated..! '
  if (text(r, 'Pending') === 'Y') return 'Selected Case is Pending..! '
  if (text(r, 'Claimability') !== 'A') return 'Please check Case Claimability..! '
  return null
}

/** Invoice > Other > OP Batch/Invoice (frmUpdateInvoiceBatch): give checked open cases an OP batch number and / or an invoice number. */
@Component({
  selector: 'app-op-batch-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="OP Batch / Invoice" subtitle="Replaces frmUpdateInvoiceBatch">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="WS No. From"><input name="casFrom" data-autofocus [(ngModel)]="f.casFrom" /></app-field>
          <app-field label="WS No. To"><input name="casTo" [(ngModel)]="f.casTo" /></app-field>
          <app-field label="Clinic">
            <input name="clinic" list="ob-clinics" [(ngModel)]="f.clinic" />
            <datalist id="ob-clinics">
              @for (c of clinics(); track c.code) {
                <option [value]="c.code + ' - ' + c.name"></option>
              }
            </datalist>
          </app-field>
          <app-field label="Payor">
            <input name="groupCompany" list="ob-groups" [(ngModel)]="f.groupCompany" />
            <datalist id="ob-groups">
              @for (g of groups(); track $index) {
                <option [value]="g"></option>
              }
            </datalist>
          </app-field>
          <app-field label="Visit Date From"><app-date-input name="visitFrom" ariaLabel="Visit Date From" [(ngModel)]="f.visitFrom" (focusout)="checkDate(f.visitFrom)" /></app-field>
          <app-field label="Visit Date To"><app-date-input name="visitTo" ariaLabel="Visit Date To" [(ngModel)]="f.visitTo" (focusout)="checkDate(f.visitTo)" /></app-field>
          <app-field label="Invoice No."><input name="invoiceFilter" [(ngModel)]="f.invoiceNo" /></app-field>
          <app-field label="Invoice Date From"><app-date-input name="invoiceFrom" ariaLabel="Invoice Date From" [(ngModel)]="f.invoiceFrom" (focusout)="checkDate(f.invoiceFrom)" /></app-field>
          <app-field label="Invoice Date To"><app-date-input name="invoiceTo" ariaLabel="Invoice Date To" [(ngModel)]="f.invoiceTo" (focusout)="checkDate(f.invoiceTo)" /></app-field>
          <app-field label="Invoice Amt From"><input class="num" name="amountFrom" [(ngModel)]="f.amountFrom" /></app-field>
          <app-field label="Invoice Amt To"><input class="num" name="amountTo" [(ngModel)]="f.amountTo" /></app-field>
          <app-field label="OP Batch From"><input name="batchFrom" [(ngModel)]="f.batchFrom" /></app-field>
          <app-field label="OP Batch To"><input name="batchTo" [(ngModel)]="f.batchTo" /></app-field>
          <app-field label="Bordx Ref From"><input name="bordxFrom" [(ngModel)]="f.bordxFrom" /></app-field>
          <app-field label="Bordx Ref To"><input name="bordxTo" [(ngModel)]="f.bordxTo" /></app-field>
          <app-field label="BTBG Ref From"><input name="btbgFrom" [(ngModel)]="f.btbgFrom" /></app-field>
          <app-field label="BTBG Ref To"><input name="btbgTo" [(ngModel)]="f.btbgTo" /></app-field>
          <app-field label="Register Date From"><app-date-input name="regFrom" ariaLabel="Register Date From" [(ngModel)]="f.regFrom" (focusout)="checkDate(f.regFrom)" /></app-field>
          <app-field label="Register Date To"><app-date-input name="regTo" ariaLabel="Register Date To" [(ngModel)]="f.regTo" (focusout)="checkDate(f.regTo)" /></app-field>
          <app-field label="Clinic Type">
            <select name="clinicType" [(ngModel)]="f.clinicType">
              <option value=""></option><option value="0">0 - Specialised</option><option value="1">1 - General Practisional</option><option value="2">2 - Others</option>
            </select>
          </app-field>
          <app-field label="FFS Clinic">
            <select name="ffs" [(ngModel)]="f.ffs"><option value=""></option><option value="Y">Y - Yes</option><option value="N">N - No</option></select>
          </app-field>
          <app-field label="Invoiced (OP Batch)">
            <select name="hasBatch" [(ngModel)]="f.hasBatch"><option value=""></option><option value="Y">Y - Yes</option><option value="N">N - No</option></select>
          </app-field>
          <app-field label="Pay To">
            <select name="payTo" [(ngModel)]="f.payTo"><option value=""></option><option value="C-Cash Less">C-Cash Less</option><option value="M-Reimbursement">M-Reimbursement</option></select>
          </app-field>
          <app-field label="Referral Letter">
            <select name="referral" [(ngModel)]="f.referral"><option value=""></option><option value="YES">YES</option><option value="NO">NO</option></select>
          </app-field>
          <app-field label="Entity">
            <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Clinic cases</option><option [ngValue]="1">SJ cases</option></select>
          </app-field>
          <label class="check-field"><input type="checkbox" name="noBatch" [(ngModel)]="f.noBatch" /> Without OP Batch</label>
          <label class="check-field"><input type="checkbox" name="noInvoice" [(ngModel)]="f.noInvoice" /> Without Invoice No.</label>
          <label class="check-field"><input type="checkbox" name="noBordx" [(ngModel)]="f.noBordx" /> Without Bordx</label>
          <label class="check-field"><input type="checkbox" name="noBtbgBordx" [(ngModel)]="f.noBtbgBordx" /> Without BTBG Bordx</label>
          <label class="check-field"><input type="checkbox" name="omni" [(ngModel)]="f.omni" /> OMNI Status: Y - Yes</label>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Cases" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="onChecked($event)" />
        <p class="totals" aria-label="Totals">
          Records: <strong>{{ rows().length }}</strong> · Invoice Amt: <strong>{{ money(totals().invoice) }}</strong> ·
          FFS: <strong>{{ money(totals().ffs) }}</strong> · Checked: <strong>{{ checked().size }}</strong>
        </p>

        <fieldset aria-label="Update">
          <legend>Update</legend>
          <app-field label="Invoice No."><input name="newInvoiceNo" maxlength="40" [(ngModel)]="invoiceNo" /></app-field>
          <app-field label="Invoice No. Used"><input name="lastInvoice" [value]="lastInvoice()" disabled /></app-field>
          <app-field label="OP Batch No."><input name="lastBatch" [value]="lastBatch()" disabled /></app-field>
          <app-action-bar label="Update actions">
            <button appButton (click)="checkAll()">Check</button>
            <button appButton variant="primary" [busy]="busy()" (click)="update()">Update</button>
          </app-action-bar>
        </fieldset>
      </div>
    </app-page-shell>
  `,
})
export class OpBatchPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = money
  protected readonly clinics = signal<LookupOption[]>([])
  protected readonly groups = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly totals = signal({ invoice: 0, ffs: 0 })
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly lastBatch = signal('')
  protected readonly lastInvoice = signal('')
  protected readonly busy = signal(false)
  protected readonly key = (r: Row) => String(r['CASNumber'])
  protected f = this.emptyFilters()
  protected invoiceNo = ''
  private searched = this.emptyFilters()
  private multipleMonths = false

  protected readonly columns: GridColumn<Row>[] = [
    { key: 'CASNumber', label: 'Case No' }, { key: 'MBMNumber', label: 'Member No' }, { key: 'CoverID', label: 'Cov' },
    { key: 'PatientName', label: 'Patient Name' }, { key: 'DateVisit', label: 'Visit Date', value: (r) => formatDate(r['DateVisit']) },
    { key: 'MBMGroupCompany', label: 'Group Company' }, { key: 'INVAmount', label: 'Invoice Amt', align: 'right', value: (r) => money(r['INVAmount']) },
    { key: 'CLMMCOProcFee', label: 'FFS', align: 'right', value: (r) => money(r['CLMMCOProcFee']) },
    { key: 'CLMFinalApprove', label: 'Final Approve', align: 'right', value: (r) => money(r['CLMFinalApprove']) },
    { key: 'BillNo', label: 'Invoice No' }, { key: 'MBMDepartment', label: 'Department' }, { key: 'CLNName', label: 'Clinic' },
    { key: 'BillDate', label: 'Invoice Date', value: (r) => formatDate(r['BillDate']) }, { key: 'INVEnteredDate', label: 'Entered Date' },
    { key: 'INVEnteredBy', label: 'Entered By' }, { key: 'OPBatchNo', label: 'OP Batch No' }, { key: 'PNO', label: 'PV No' },
    { key: 'PDATE', label: 'PV Date', value: (r) => formatDate(r['PDATE']) }, { key: 'PCNO', label: 'Cheque No' }, { key: 'ReasonDesc1', label: 'Reason 1' },
    { key: 'ReasonDesc2', label: 'Reason 2' }, { key: 'ReasonDesc3', label: 'Reason 3' }, { key: 'ReasonDesc4', label: 'Reason 4' },
    { key: 'ReasonDesc5', label: 'Reason 5' }, { key: 'CASDiagnosis5', label: 'Diagnosis' }, { key: 'MCDays', label: 'MC Days' },
    { key: 'CLMCoPayCLNColl', label: 'Co-Pay Coll.', align: 'right' }, { key: 'CLMFinalApprove1', label: 'Balance', align: 'right', value: (r) => money(r['CLMFinalApprove1']) },
    { key: 'ICNo', label: 'IC No' }, { key: 'OPNO', label: 'OP' }, { key: 'PType', label: 'Pay Type' },
    { key: 'OPBatchDate', label: 'OP Batch Date', value: (r) => formatDate(r['OPBatchDate']) }, { key: 'LastUpdateDate', label: 'Last Update', value: (r) => formatDate(r['LastUpdateDate']) },
    { key: 'Pending', label: 'Pending' }, { key: 'Claimability', label: 'Claimability' },
  ]

  async ngOnInit(): Promise<void> {
    try {
      const s = await firstValueFrom(this.http.get<{ clinics: LookupOption[]; groupCompanies: string[] }>('/api/invoice/update/setup'))
      this.clinics.set(s.clinics)
      this.groups.set([...s.groupCompanies].sort())
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected checkDate(v: string): void {
    if (v && !parseDmy(v)) this.notify.notify('Invalid Date Format! ', 'error')
  }

  /** CmdSearch_Click (the invoice number is cleared, as in the desktop). */
  protected async search(): Promise<void> {
    this.busy.set(true)
    this.reset()
    try {
      const list = await firstValueFrom(this.http.post<OpBatchList>(`${BASE}/search`, this.f))
      this.rows.set(list.rows)
      this.totals.set({ invoice: list.totalInvoice, ffs: list.totalFfs })
      this.multipleMonths = list.multipleMonths
      this.searched = { ...this.f }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected async print(): Promise<void> {
    if (this.rows().length === 0) return this.notify.notify('Report cannot be printed without any record.', 'info')
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/export`, this.searched, { responseType: 'blob' })), 'OPBatchInvoice.xlsx')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  /** lstInvoiceListing_ItemCheck: a case that cannot be checked is unchecked with the desktop's message. */
  protected onChecked(next: Set<string>): void {
    for (const k of next) {
      if (this.checked().has(k)) continue
      const row = this.rows().find((r) => this.key(r) === k)
      const message = row ? checkMessage(row) : null
      if (message) {
        next.delete(k)
        this.notify.notify(message, 'error')
      }
    }
    this.checked.set(new Set(next))
  }

  /** CmdCheck_Click: checks every case that may be checked. */
  protected checkAll(): void {
    if (this.rows().length === 0) return this.notify.notify('Please select a record.', 'info')
    const ok = this.rows().filter((r) => !checkMessage(r))
    this.checked.set(new Set(ok.map(this.key)))
    if (ok.length < this.rows().length)
      this.notify.notify('Un Selected Few Cases Already OP Batch or Case Pending Or Case Claimability Or Invoice No generated ..! ', 'error')
  }

  /** CmdUpdate_Click. */
  protected async update(): Promise<void> {
    if (this.invoiceNo.trim()) {
      if (this.rows().length === 0) return this.notify.notify('Please Search Data..!', 'info')
      if (!parseDmy(this.searched.visitFrom) || !parseDmy(this.searched.visitTo)) return this.notify.notify('Please select Visit Date.', 'info')
      if (!this.searched.clinic) return this.notify.notify('Above claims must belong to same clinic code.', 'info')
      if (this.multipleMonths) return this.notify.notify('All the claims visit date must belong to same Month & Year.', 'info')
    }
    if (this.checked().size === 0) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record ? '))) return
    this.busy.set(true)
    try {
      const body = { filters: this.searched, casNumbers: [...this.checked()], invoiceNo: this.invoiceNo }
      const r = await firstValueFrom(this.http.post<{ message: string; batchNo: string | null; invoiceNo: string | null }>(BASE, body))
      this.notify.notify(r.message, 'success')
      this.reset()
      this.lastBatch.set(r.batchNo ?? '')
      this.lastInvoice.set(r.invoiceNo ?? '')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected clear(): void {
    this.f = this.emptyFilters()
    this.reset()
  }

  private reset(): void {
    this.rows.set([])
    this.totals.set({ invoice: 0, ffs: 0 })
    this.checked.set(new Set())
    this.invoiceNo = ''
    this.lastBatch.set('')
    this.lastInvoice.set('')
  }

  private emptyFilters() {
    return {
      noBatch: false, noBordx: false, bordxFrom: '', bordxTo: '', noBtbgBordx: false, btbgFrom: '', btbgTo: '', casFrom: '', casTo: '', clinic: '',
      groupCompany: '', invoiceNo: '', visitFrom: '', visitTo: '', noInvoice: false, invoiceFrom: '', invoiceTo: '', amountFrom: '', amountTo: '',
      batchFrom: '', batchTo: '', omni: false, regFrom: '', regTo: '', clinicType: '', ffs: '', hasBatch: '', payTo: '', referral: '', entity: 0,
    }
  }
}
