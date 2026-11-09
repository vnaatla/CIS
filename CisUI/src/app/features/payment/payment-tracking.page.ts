import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { ActivatedRoute } from '@angular/router'
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

const BASE = '/api/payment/tracking'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const isDate = (v: unknown) => typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v)
const money = /Amt|Paid|Payable|Fee|Approve|Actual/i

/** Payment > Tracking (FrmPayment) and Tracking (Jupiter) (frmPayment2). */
@Component({
  selector: 'app-payment-tracking-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="title" [subtitle]="'Replaces ' + form">
      <div class="screen-form">
        @if (!fullView()) {
          <fieldset aria-label="Search">
            <legend>Search</legend>
            <app-field label="Report Type" [required]="true">
              <select name="reportType" [(ngModel)]="f.reportType"><option [ngValue]="null"></option><option [ngValue]="0">Pay to Clinic</option><option [ngValue]="1">Pay to Member</option></select>
            </app-field>
            <app-field label="Report Selection">
              <select name="selection" [(ngModel)]="f.selection">@for (x of selections; track x) { <option [value]="x">{{ x }}</option> }</select>
            </app-field>
            <app-field label="Payment">
              <select name="payment" [(ngModel)]="f.payment"><option [ngValue]="0">0 - All</option><option [ngValue]="1">1 - Paid</option><option [ngValue]="2">2 - Partial Paid</option></select>
            </app-field>
            <app-field label="Claim No From"><input name="caseFrom" maxlength="30" [(ngModel)]="f.caseFrom" /></app-field>
            <app-field label="Claim No To"><input name="caseTo" maxlength="30" [(ngModel)]="f.caseTo" /></app-field>
            <app-field label="Group Company" [wide]="true">
              <input name="company" list="pt-companies" [(ngModel)]="f.company" />
              <datalist id="pt-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
            </app-field>
            <app-field label="Clinic">
              <input name="clinic" list="pt-clinics" [(ngModel)]="f.clinic" />
              <datalist id="pt-clinics">@for (c of clinics(); track c.code) { <option [value]="c.code + ' - ' + c.name"></option> }</datalist>
            </app-field>
            <app-field label="Member No"><input name="mbmNo" maxlength="30" [(ngModel)]="f.mbmNo" /></app-field>
            <app-field label="Invoice No From"><input name="invoiceFrom" maxlength="50" [(ngModel)]="f.invoiceFrom" /></app-field>
            <app-field label="Invoice No To"><input name="invoiceTo" maxlength="50" [(ngModel)]="f.invoiceTo" /></app-field>
            <app-field label="Invoice Date From"><app-date-input name="invDateFrom" ariaLabel="Invoice Date From" [(ngModel)]="f.invDateFrom" /></app-field>
            <app-field label="Invoice Date To"><app-date-input name="invDateTo" ariaLabel="Invoice Date To" [(ngModel)]="f.invDateTo" /></app-field>
            <app-field label="Bordx No From"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" /></app-field>
            <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
            <app-field label="OP Batch From"><input name="opFrom" maxlength="30" [(ngModel)]="f.opFrom" /></app-field>
            <app-field label="OP Batch To"><input name="opTo" maxlength="30" [(ngModel)]="f.opTo" /></app-field>
            <app-field label="PV No From"><input name="pvFrom" maxlength="30" [(ngModel)]="f.pvFrom" /></app-field>
            <app-field label="PV No To"><input name="pvTo" maxlength="30" [(ngModel)]="f.pvTo" /></app-field>
            <app-field label="PV Date From"><app-date-input name="pvDateFrom" ariaLabel="PV Date From" [(ngModel)]="f.pvDateFrom" /></app-field>
            <app-field label="PV Date To"><app-date-input name="pvDateTo" ariaLabel="PV Date To" [(ngModel)]="f.pvDateTo" /></app-field>
            <app-field label="Amount From"><input class="num" name="amountFrom" maxlength="20" [(ngModel)]="f.amountFrom" /></app-field>
            <app-field label="Amount To"><input class="num" name="amountTo" maxlength="20" [(ngModel)]="f.amountTo" /></app-field>
            <app-field label="A/C Month From"><select name="monthFrom" [(ngModel)]="f.monthFrom"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
            <app-field label="A/C Month To"><select name="monthTo" [(ngModel)]="f.monthTo"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select></app-field>
            <app-field label="A/C Year From"><input name="yearFrom" maxlength="4" [(ngModel)]="f.yearFrom" /></app-field>
            <app-field label="A/C Year To"><input name="yearTo" maxlength="4" [(ngModel)]="f.yearTo" /></app-field>
            <app-action-bar label="Tracking actions">
              <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
              <button appButton [disabled]="!rows().length" (click)="fullView.set(true)">View</button>
              <button appButton (click)="export()">Export</button>
              <button appButton (click)="clear()">Clear</button>
            </app-action-bar>
          </fieldset>
        } @else {
          <app-action-bar label="Tracking view actions"><button appButton (click)="fullView.set(false)">Back</button></app-action-bar>
        }
        <app-data-grid caption="Payment tracking" [columns]="columns()" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong> · Total Amount: <strong>{{ formatMoney(total()) }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class PaymentTrackingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  private readonly data = inject(ActivatedRoute).snapshot.data
  protected readonly title: string = this.data['screenTitle'] ?? 'Payment Tracking'
  protected readonly form: string = this.data['legacyForm'] ?? 'FrmPayment'

  protected readonly formatMoney = formatMoney
  protected readonly months = Array.from({ length: 12 }, (_, i) => String(i + 1))
  protected readonly selections = ['By Claim', 'By Bordx', 'By PV No', 'By Clinic', 'By Group Company', 'General Listing']
  protected readonly companies = signal<string[]>([])
  protected readonly clinics = signal<{ code: string; name: string }[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly total = signal(0)
  protected readonly busy = signal(false)
  protected readonly fullView = signal(false)
  protected f = this.empty()
  protected readonly key = (r: Row) => JSON.stringify(r)
  /** The columns of the listing returned (each selection has its own). */
  protected readonly columns = computed<GridColumn<Row>[]>(() =>
    Object.keys(this.rows()[0] ?? {}).map((k) => {
      const sample = this.rows().find((r) => r[k] != null)?.[k]
      if (isDate(sample)) return { key: k, label: k, value: (r: Row) => (r[k] ? formatDate(String(r[k])) : '') }
      if (typeof sample === 'number' && money.test(k)) return { key: k, label: k, align: 'right' as const, value: (r: Row) => (r[k] == null ? '' : formatMoney(Number(r[k]))) }
      return { key: k, label: k }
    }),
  )

  private empty() {
    return {
      reportType: null as number | null, selection: 'General Listing', payment: 0, caseFrom: '', caseTo: '', company: '', clinic: '', mbmNo: '', invoiceFrom: '',
      invoiceTo: '', invDateFrom: '', invDateTo: '', bordxFrom: '', bordxTo: '', opFrom: '', opTo: '', pvFrom: '', pvTo: '', pvDateFrom: '', pvDateTo: '',
      amountFrom: '', amountTo: '', monthFrom: '', monthTo: '', yearFrom: '', yearTo: '',
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ companies: string[]; clinics: { code: string; name: string }[] }>(`${BASE}/lookups`))
      this.companies.set(l.companies)
      this.clinics.set(l.clinics)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    if (this.f.reportType === null) return this.notify.notify('Please select a report type.', 'info')
    this.busy.set(true)
    this.rows.set([])
    this.total.set(0)
    try {
      const r = await firstValueFrom(this.http.post<{ rows: Row[]; totalPaid: number }>(`${BASE}/search`, this.f))
      this.rows.set(r.rows)
      this.total.set(r.totalPaid)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdExport_Click. */
  protected export(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    const cols = this.columns()
    saveCsv([cols.map((c) => c.label), ...this.rows().map((r) => cols.map((c) => (c.value ? c.value(r) : s(r[c.key]))))], 'PaymentTracking.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
    this.total.set(0)
  }
}
