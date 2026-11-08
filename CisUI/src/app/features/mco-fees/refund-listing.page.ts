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

const BASE = '/api/mco-fees/refund-listing'
type Row = Record<string, unknown>
interface Result {
  rows: Row[]
  principals: number
  supplementaries: number
  principalRefund: number
  utilisedFee: number
  supplementaryRefund: number
  totalRefund: number
}

const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null ? '' : formatMoney(Number(v)))

/** MCO Fees > Refund Listing (frmMCORefundListing). */
@Component({
  selector: 'app-mco-refund-listing-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="MCO Refund Listing" subtitle="Replaces frmMCORefundListing">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Listing Version" [required]="true">
            <select name="version" [(ngModel)]="f.version" (ngModelChange)="versionChanged()">
              <option [ngValue]="null"></option><option [ngValue]="0">MCO Listing - by Payor</option>
              <option [ngValue]="1">MCO Listing - by Group Company</option><option [ngValue]="2">MCO Listing - by Department</option>
            </select>
          </app-field>
          @if (f.version === 0) {
            <app-field label="Payor">
              <select name="payor" [(ngModel)]="f.payor"><option value=""></option>@for (p of payors(); track p.code) { <option [value]="p.code + ' - ' + p.name">{{ p.code }} - {{ p.name }}</option> }</select>
            </app-field>
          } @else if (f.version !== null) {
            <app-field label="Group Company" [wide]="true">
              <input name="groupCompany" list="refund-groups" [(ngModel)]="f.groupCompany" (change)="loadDepartments()" />
              <datalist id="refund-groups">@for (g of companies(); track g) { <option [value]="g"></option> }</datalist>
            </app-field>
            @if (f.version === 2) {
              <app-field label="Department">
                <select name="department" [(ngModel)]="f.department"><option value=""></option>@for (x of departments(); track x) { <option [value]="x">{{ x }}</option> }</select>
              </app-field>
            }
          }
          <app-field label="Date Selection">
            <select name="dateSelection" [(ngModel)]="f.dateSelection" (ngModelChange)="rows.set([])">
              <option [ngValue]="0">Endt. Eff Date</option><option [ngValue]="1">Endt. Bordx Date</option>
            </select>
          </app-field>
          <app-field label="From"><app-date-input name="from" ariaLabel="From" [(ngModel)]="f.from" /></app-field>
          <app-field label="To"><app-date-input name="to" ariaLabel="To" [(ngModel)]="f.to" /></app-field>
          <app-field label="Print Status">
            <select name="status" [(ngModel)]="f.status"><option value="N">N - No</option><option value="Y">Y - Yes</option></select>
          </app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="file('export', 'MCORefundListing.xlsx')">Print</button>
            <button appButton (click)="file('credit-note', 'MCOCreditNote.xlsx')">Print CR Note</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="MCO refunds" [columns]="columns" [rows]="rows()" [rowKey]="key" />
        @if (result(); as r) {
          <p class="totals" aria-label="Totals">
            Principal: <strong>{{ r.principals }}</strong> · Supplementary: <strong>{{ r.supplementaries }}</strong> ·
            Refund (principal): <strong>{{ money(r.principalRefund) }}</strong> · MCO Fee Utilised: <strong>{{ money(r.utilisedFee) }}</strong> ·
            Refund (supplementary): <strong>{{ money(r.supplementaryRefund) }}</strong> · Total Refund: <strong>{{ money(r.totalRefund) }}</strong>
          </p>
        }
      </div>
    </app-page-shell>
  `,
})
export class McoRefundListingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = m
  protected readonly payors = signal<{ code: string; name: string }[]>([])
  protected readonly companies = signal<string[]>([])
  protected readonly departments = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly result = signal<Result | null>(null)
  protected readonly busy = signal(false)
  protected f = this.empty()
  protected readonly key = (r: Row) => `${r['member']}-${r['cover']}-${r['principal']}`
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'member', label: 'Member No' }, { key: 'cover', label: 'Cov' }, { key: 'name', label: 'Name' }, { key: 'ic', label: 'IC No' },
    { key: 'bordxDate', label: 'Bordx Date', value: (r) => d(r['bordxDate']) }, { key: 'effDate', label: 'Eff Date', value: (r) => d(r['effDate']) },
    { key: 'expDate', label: 'Exp Date', value: (r) => d(r['expDate']) }, { key: 'daysInsured', label: 'Days Insured' },
    { key: 'endtEffDate', label: 'Endt Eff Date', value: (r) => d(r['endtEffDate']) }, { key: 'endtBordxDate', label: 'Endt Bordx Date', value: (r) => d(r['endtBordxDate']) },
    { key: 'daysUtilised', label: 'Days Utilised' }, { key: 'mcoFee', label: 'MCO Fee', align: 'right', value: (r) => m(r['mcoFee']) },
    { key: 'mcoUtilised', label: 'MCO Utilised', align: 'right', value: (r) => m(r['mcoUtilised']) },
    { key: 'mcoRefund', label: 'MCO Refund', align: 'right', value: (r) => m(r['mcoRefund']) },
    { key: 'mcoUtilisedFee', label: 'MCO Fee Utilised', align: 'right', value: (r) => m(r['mcoUtilisedFee']) },
    { key: 'invoiceNo', label: 'Invoice No' }, { key: 'invoiceDate', label: 'Invoice Date', value: (r) => d(r['invoiceDate']) }, { key: 'cases', label: 'Cases' },
    { key: 'creditNote', label: 'CR Note' }, { key: 'creditNoteDate', label: 'CR Note Date', value: (r) => d(r['creditNoteDate']) },
  ]

  private empty() {
    return { version: null as number | null, dateSelection: 0, from: '', to: '', payor: '', groupCompany: '', department: '', status: 'N' }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ payors: { code: string; name: string }[]; companies: string[] }>(`${BASE}/lookups`))
      this.payors.set(l.payors)
      this.companies.set(l.companies)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected versionChanged(): void {
    this.f.dateSelection = 0
    this.f.payor = ''
    this.f.groupCompany = ''
    this.f.department = ''
  }

  /** cboMCOGrpCom_Click: the company's departments. */
  protected async loadDepartments(): Promise<void> {
    if (this.f.version !== 2) return
    try {
      const l = await firstValueFrom(this.http.get<{ departments: string[] }>(`${BASE}/lookups`, { params: { company: this.f.groupCompany } }))
      this.departments.set(l.departments)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.result.set(null)
    try {
      const r = await firstValueFrom(this.http.post<Result>(`${BASE}/search`, this.f))
      this.result.set(r)
      this.rows.set(r.rows)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected async file(kind: 'export' | 'credit-note', name: string): Promise<void> {
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/${kind}`, this.f, { responseType: 'blob' })), name)
      if (kind === 'credit-note' && this.f.status === 'N') await this.search()
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
