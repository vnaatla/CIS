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
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/mco-fees/fee-listing'
type Row = Record<string, unknown>
interface Option {
  code: string
  name: string
}
interface Result {
  rows: Row[]
  principals: number
  supplementaries: number
  canPrint: boolean
  multipleInvoices: boolean
}

const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')

/** MCO Fees > Fee Listing (frmMCOFeeListNew). */
@Component({
  selector: 'app-mco-fee-listing-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="MCO Fee Listing" subtitle="Replaces frmMCOFeeListNew">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="MCO Criteria" [required]="true">
            <select name="criteria" [(ngModel)]="f.criteria" (ngModelChange)="criteriaChanged()">
              <option [ngValue]="null"></option><option [ngValue]="0">Bordereaux Batch</option><option [ngValue]="1">Bordereaux Range</option>
              <option [ngValue]="2">Invoice</option>
            </select>
          </app-field>
          @switch (f.criteria) {
            @case (0) {
              <app-field label="Bordx Date"><app-date-input name="bordxDate" ariaLabel="Bordx Date" [(ngModel)]="f.bordxDate" /></app-field>
              <app-field label="Batch No."><input name="batchNo" maxlength="10" [(ngModel)]="f.batchNo" /></app-field>
            }
            @case (1) {
              <app-field label="From"><app-date-input name="from" ariaLabel="From" [(ngModel)]="f.from" /></app-field>
              <app-field label="To"><app-date-input name="to" ariaLabel="To" [(ngModel)]="f.to" /></app-field>
            }
            @case (2) {
              <app-field label="Invoice No."><input name="invoiceNo" maxlength="30" [(ngModel)]="f.invoiceNo" /></app-field>
            }
          }
          <app-field label="Listing Version" [required]="true">
            <select name="version" [(ngModel)]="f.version">
              <option [ngValue]="0">MCO Listing - by Payor</option><option [ngValue]="1">MCO Listing - by Group Company</option>
              <option [ngValue]="2">MCO Listing - by Department</option>
            </select>
          </app-field>
          @if (f.version === 0) {
            <app-field label="Payor">
              <select name="payor" [(ngModel)]="f.payor"><option value=""></option>@for (p of payors(); track p.code) { <option [value]="p.code + ' - ' + p.name">{{ p.code }} - {{ p.name }}</option> }</select>
            </app-field>
          } @else {
            <app-field label="Group Company" [wide]="true">
              <input name="groupCompany" list="mco-groups" [(ngModel)]="f.groupCompany" />
              <datalist id="mco-groups">@for (g of groups(); track g.code) { <option [value]="g.code + ' - ' + g.name"></option> }</datalist>
            </app-field>
          }
          <app-field label="Print Status">
            <select name="invoiced" [(ngModel)]="f.invoiced" (ngModelChange)="rows.set([])"><option value="N">N - No</option><option value="Y">Y - Yes</option></select>
          </app-field>
          <app-field label="Report Type">
            <select name="reportType" [(ngModel)]="reportType"><option [ngValue]="null"></option><option [ngValue]="0">MCO Invoice</option><option [ngValue]="1">Fee Listing</option></select>
          </app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [disabled]="!result()" (click)="print(false)">Print Preview</button>
            <button appButton [disabled]="!result()?.canPrint" [busy]="printing()" (click)="print(true)">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="MCO fees" [columns]="columns" [rows]="rows()" [rowKey]="key" [selectable]="true" [selectedKey]="selected()" (rowSelected)="pick($any($event))" />
        <p class="totals" aria-label="Totals">Principal: <strong>{{ result()?.principals ?? 0 }}</strong> · Supplementary: <strong>{{ result()?.supplementaries ?? 0 }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class McoFeeListingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly payors = signal<Option[]>([])
  protected readonly groups = signal<Option[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly result = signal<Result | null>(null)
  protected readonly selected = signal<string | null>(null)
  protected readonly busy = signal(false)
  protected readonly printing = signal(false)
  protected reportType: number | null = 1
  protected f = this.empty()
  private searched = this.empty()
  protected readonly key = (r: Row) => `${s(r['MBMNumber'])}-${s(r['MBMCCoverID'])}`
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'PAYCode', label: 'Payor' }, { key: 'MBMNumber', label: 'Member No' }, { key: 'MBMCCoverID', label: 'Cov' }, { key: 'MBMCRELCode', label: 'Rel' },
    { key: 'MBMBordxDate', label: 'Bordx Date', value: (r) => d(r['MBMBordxDate']) }, { key: 'MBMPolEffDate', label: 'Eff Date', value: (r) => d(r['MBMPolEffDate']) },
    { key: 'MBMPolExpDate', label: 'Exp Date', value: (r) => d(r['MBMPolExpDate']) },
    { key: 'MBMCMCOFee', label: 'MCO Fee', align: 'right', value: (r) => formatMoney(Number(r['MBMCMCOFee'] ?? 0)) }, { key: 'MBMBatchNo', label: 'Batch' },
    { key: 'MBMName', label: 'Name' }, { key: 'MBMGroupCompany', label: 'Group Company' }, { key: 'GRPCoCode', label: 'Grp Code' }, { key: 'MBMDepartment', label: 'Department' },
    { key: 'invoiced', label: 'Invoiced', value: (r) => (s(r['MBMCInvoiceNo']) ? 'YES' : 'NO') }, { key: 'MBMCInvoiceNo', label: 'Invoice No' },
    { key: 'MBMCInvoiceDate', label: 'Invoice Date', value: (r) => d(r['MBMCInvoiceDate']) },
  ]

  private empty() {
    return { criteria: null as number | null, bordxDate: '', batchNo: '', from: '', to: '', invoiceNo: '', version: 1, payor: '', groupCompany: '', invoiced: 'N' }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ payors: Option[]; groups: Option[] }>(`${BASE}/lookups`))
      this.payors.set(l.payors)
      this.groups.set(l.groups)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected criteriaChanged(): void {
    this.f = { ...this.empty(), criteria: this.f.criteria, version: this.f.version, invoiced: this.f.invoiced }
    this.rows.set([])
    this.result.set(null)
  }

  protected pick(r: Row): void {
    this.selected.set(this.key(r))
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.result.set(null)
    try {
      const r = await firstValueFrom(this.http.post<Result>(`${BASE}/search`, this.f))
      this.searched = { ...this.f }
      this.result.set(r)
      this.rows.set(r.rows)
      this.selected.set(r.rows.length ? this.key(r.rows[0]) : null)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected async print(print: boolean): Promise<void> {
    const row = this.rows().find((r) => this.key(r) === this.selected()) ?? this.rows()[0]
    const body = { search: this.searched, reportType: this.reportType, selectedPayor: s(row?.['PAYCode']), selectedBatch: s(row?.['MBMBatchNo']), print }
    this.printing.set(true)
    try {
      const blob = await firstValueFrom(this.http.post(`${BASE}/print`, body, { responseType: 'blob' }))
      saveBlob(blob, this.reportType === 0 ? 'MCOInvoice.pdf' : 'MCOFeeListing.pdf')
      if (print && this.reportType === 0) await this.search()
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    } finally {
      this.printing.set(false)
    }
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
    this.result.set(null)
  }
}
