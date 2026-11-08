import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatDate } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/mco-fees/invoice-gst-zeroised'
const LOOKUPS = '/api/mco-fees/invoice-gst'
interface Option {
  code: string
  name: string
}
type InvRow = Record<string, unknown>
interface Result {
  rows: InvRow[]
  principals: number
  supplementaries: number
  invoiceNo: string
  multipleInvoices: boolean
}

/** MCO Fees > MCO Invoice GST Zeroised (frmMCOFeeGSTZeroised). */
@Component({
  selector: 'app-mco-gst-zeroised-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="MCO Invoice GST Zeroised" subtitle="Replaces frmMCOFeeGSTZeroised">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Criteria" [required]="true">
            <select name="criteria" [(ngModel)]="f.criteria" (ngModelChange)="criteriaChanged()">
              <option [ngValue]="null"></option><option [ngValue]="0">Bordereaux Batch</option><option [ngValue]="1">Bordereaux Range</option>
              <option [ngValue]="2">Invoice No</option>
            </select>
          </app-field>
          @switch (f.criteria) {
            @case (0) {
              <app-field label="Bordx Date"><app-date-input name="bordxDate" ariaLabel="Bordx Date" [(ngModel)]="f.bordxDate" /></app-field>
              <app-field label="Batch Num"><input name="batchNo" maxlength="10" [(ngModel)]="f.batchNo" /></app-field>
            }
            @case (1) {
              <app-field label="Bordx From"><app-date-input name="from" ariaLabel="Bordx From" [(ngModel)]="f.from" /></app-field>
              <app-field label="To"><app-date-input name="to" ariaLabel="To" [(ngModel)]="f.to" /></app-field>
            }
            @case (2) {
              <app-field label="Invoice Num"><input name="invoiceNo" maxlength="30" [(ngModel)]="f.invoiceNo" /></app-field>
            }
          }
          <app-field label="Listing Version">
            <select name="version" [(ngModel)]="f.version">
              <option [ngValue]="0">Invoice - by Payor</option><option [ngValue]="1">Invoice - by Group Company</option>
            </select>
          </app-field>
          @if (f.version === 0) {
            <app-field label="Payor">
              <select name="payor" [(ngModel)]="f.payor" [disabled]="f.criteria === 2"><option value=""></option>@for (p of payors(); track p.code) { <option [value]="p.code + ' - ' + p.name">{{ p.code }} - {{ p.name }}</option> }</select>
            </app-field>
          } @else {
            <app-field label="Group Company" [wide]="true">
              <input name="groupCompany" list="inv-groups" [(ngModel)]="f.groupCompany" [disabled]="f.criteria === 2" (change)="loadDepartments()" />
              <datalist id="inv-groups">@for (g of groups(); track g.code) { <option [value]="g.code + ' - ' + g.name"></option> }</datalist>
            </app-field>
            <app-field label="Subsidiary">
              <select name="subsidiary" [(ngModel)]="f.subsidiary"><option value=""></option>@for (x of departments(); track x) { <option [value]="x">{{ x }}</option> }</select>
            </app-field>
          }
          <app-field label="Policy No"><input name="policyNo" maxlength="50" [(ngModel)]="policyNo" /></app-field>
          <app-field label="Print Option">
            <select name="printOption" [(ngModel)]="f.printOption" [disabled]="f.criteria === 2"><option [ngValue]="0">Print New</option><option [ngValue]="1">Reprint</option></select>
          </app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [disabled]="!result()" (click)="checkAll()">{{ allChecked() ? 'Uncheck All' : 'Check All' }}</button>
            <button appButton (click)="printExcel()">Print Excel</button>
            <button appButton [disabled]="!result() || printed()" [busy]="printing()" (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="MCO GST zeroised members" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()" (checkedChange)="check($event)" />
        <p class="totals" aria-label="Totals">
          Principal: <strong>{{ result()?.principals ?? 0 }}</strong> · Supplementary: <strong>{{ result()?.supplementaries ?? 0 }}</strong> ·
          Invoice No: <strong>{{ result()?.invoiceNo || '-' }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class McoGstZeroisedPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly payors = signal<Option[]>([])
  protected readonly groups = signal<Option[]>([])
  protected readonly departments = signal<string[]>([])
  protected readonly rows = signal<InvRow[]>([])
  protected readonly result = signal<Result | null>(null)
  protected readonly busy = signal(false)
  protected readonly printing = signal(false)
  protected readonly printed = signal(false)
  protected f = this.empty()
  private searched = this.empty()
  protected readonly key = (r: InvRow) => `${r['MBMCNumber']}-${r['MBMCCoverID']}`
  protected readonly columns: GridColumn<InvRow>[] = [
    { key: 'PAYCode', label: 'Payor' }, { key: 'MBMCNumber', label: 'Member No' }, { key: 'MBMCCoverID', label: 'Cov' }, { key: 'RelCode', label: 'Rel' },
    { key: 'MBMBordxDate', label: 'Bordx Date', value: (r) => (r['MBMBordxDate'] ? formatDate(String(r['MBMBordxDate'])) : '') },
    { key: 'MBMPolicyNo', label: 'Policy No' }, { key: 'MBMBatchNo', label: 'Batch' }, { key: 'MBMName', label: 'Name' },
    { key: 'MBMGroupCompany', label: 'Group Company' }, { key: 'invoiced', label: 'Invoiced', value: (r) => (r['MBMCInvoiceNo'] ? 'YES' : 'NO') },
    { key: 'MBMCInvoiceNo', label: 'Invoice No' }, { key: 'Subsidiary', label: 'Subsidiary' },
  ]

  private empty() {
    return {
      criteria: null as number | null, bordxDate: '', batchNo: '', from: '', to: '', invoiceNo: '', version: 0, payor: '', groupCompany: '', subsidiary: '',
      printOption: 0,
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ payors: Option[]; groups: Option[] }>(`${LOOKUPS}/lookups`))
      this.payors.set(l.payors)
      this.groups.set(l.groups)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** INVCriterialSelDisp: batch / range print new by payor; invoice no reprint by group company. */
  protected criteriaChanged(): void {
    const c = this.f.criteria
    this.f = { ...this.empty(), criteria: c, printOption: c === 2 ? 1 : 0, version: c === 2 ? 1 : 0 }
    this.rows.set([])
    this.result.set(null)
  }

  /** BindSubsidary. */
  protected async loadDepartments(): Promise<void> {
    try {
      this.departments.set(await firstValueFrom(this.http.get<string[]>(`${LOOKUPS}/subsidiaries`, { params: { groupCompany: this.f.groupCompany } })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.result.set(null)
    this.printed.set(false)
    try {
      const r = await firstValueFrom(this.http.post<Result>(`${BASE}/search`, { search: this.f, policyNo: this.policyNo }))
      this.searched = { ...this.f }
      this.result.set(r)
      this.rows.set(r.rows)
      this.checked.set(new Set())
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected policyNo = ''

  protected allChecked(): boolean {
    return this.rows().length > 0 && this.checked().size === this.rows().length
  }

  /** lvwInvoice_ItemCheck: an invoiced member cannot be checked. */
  protected check(keys: Set<string>): void {
    const invoiced = this.rows().filter((r) => keys.has(this.key(r)) && r['MBMCInvoiceNo'])
    if (invoiced.length) {
      this.notify.notify('The selected member already invoice generated,Please check..!', 'error')
      invoiced.forEach((r) => keys.delete(this.key(r)))
    }
    this.checked.set(keys)
  }

  /** CmdCheck_Click. */
  protected checkAll(): void {
    if (!this.rows().length) return this.notify.notify('Please select a record.', 'info')
    this.check(this.allChecked() ? new Set() : new Set(this.rows().map(this.key)))
  }

  /** cmdINVPrint_Click. */
  protected async print(): Promise<void> {
    this.printing.set(true)
    try {
      const body = { search: this.searched, policyNo: this.policyNo, checked: [...this.checked()] }
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/print`, body, { responseType: 'blob' })), 'MCOInvoiceGSTZeroised.xlsx')
      this.printed.set(true)
      this.f.printOption = 1
      await this.search()
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    } finally {
      this.printing.set(false)
    }
  }

  /** CmdPrintExcel_Click. */
  protected async printExcel(): Promise<void> {
    if (this.f.criteria !== 2 || !this.f.invoiceNo.trim()) return this.notify.notify('Please search with invoice no.', 'info')
    const bySubsidiary = await this.notify.ask('Are you sure the above invoice no print by Subsidary ')
    try {
      const body = { search: this.f, policyNo: this.policyNo, bySubsidiary }
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/print-invoice`, body, { responseType: 'blob' })), `${this.f.invoiceNo.trim()}.xlsx`)
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
    this.result.set(null)
  }
}
