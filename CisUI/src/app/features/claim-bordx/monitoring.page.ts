import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
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

const BASE = '/api/claim-bordx/monitoring'
type Row = Record<string, unknown>
type Kind = 'control' | 'outstanding' | 'bordx-worksheet'
interface Result {
  rows: Row[]
  total: number
  total14: number
  totalOver14: number
}
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })
const date = (key: string, label: string): GridColumn<Row> => ({ key, label, value: (r) => d(r[key]) })

const screens: Record<Kind, { title: string; form: string; columns: GridColumn<Row>[] }> = {
  control: {
    title: 'Bordx Control', form: 'FrmBordxControl',
    columns: [{ key: 'BordxRefNo', label: 'Bordx No' }, date('BordxDate', 'Bordx Date'), { key: 'BordxNoClaims', label: 'Claims', align: 'right' },
      money('BillAmt', 'Bill Amt'), money('CLMCoPayCLNColl', 'Co-Pay'), money('CLMFinalApprove', 'Approved'), money('CLMMCOProcFee', 'MCO Fee'),
      money('CLMBordxAmt', 'Bordx Amt'), money('CLMNonGpFee', 'Non GP Fee'), date('BordxPAYSentDate', 'Sent'), date('BordxPAYRecDocDate', 'Doc Returned'),
      { key: 'NoofDay', label: 'Days', align: 'right' }, date('BordxPAYRecDate', 'Paid'), { key: 'NoofDay2', label: 'Days Paid', align: 'right' },
      { key: 'MBMGroupCompany', label: 'Group Company' }, { key: 'Subsidiary', label: 'Subsidiary' }, { key: 'INVPayTo', label: 'Pay To' }],
  },
  outstanding: {
    title: 'Outstanding Bordereaux', form: 'FrmOSBordx',
    columns: [{ key: 'No', label: 'No', align: 'right' }, { key: 'BordxRefNo', label: 'Bordx No' }, { key: 'BordxPAYCode', label: 'Payor' },
      money('BordxClaimAmt', 'Amount'), date('BordxPAYSentDate', 'Date Sent'), { key: 'Days', label: 'Days', align: 'right' }, money('Within14', '<= 14 Days'),
      money('Over14', '> 14 Days')],
  },
  'bordx-worksheet': {
    title: 'Bordereaux & Worksheet', form: 'FrmBordxWks',
    columns: [{ key: 'BordxPAYCode', label: 'Payor' }, { key: 'GroupCompany', label: 'Group Company' }, { key: 'BordxRefNo', label: 'Bordx No' },
      money('BordxClaimAmt', 'Bordx Amt'), date('BordxDate', 'Bordx Date'), { key: 'CASNumber', label: 'Worksheet No' }, money('CLMTotalApproved', 'Approved'),
      date('ClaimOpendate', 'Open Date'), date('BordxPAYRecDate', 'Date Fr Payor'), { key: 'ReceiptNoPAY', label: 'Receipt No' }],
  },
}

/** Claim Bordx > Monitoring: Bordx Control, Outstanding Bordx, Bordx & Worksheet (route data `kind`). */
@Component({
  selector: 'app-bordx-monitoring-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="screen.title" [subtitle]="'Replaces ' + screen.form">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Group Company" [wide]="true">
            <input name="company" list="monitoring-companies" [(ngModel)]="f.company" (change)="companyChanged()" />
            <datalist id="monitoring-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          @switch (kind) {
            @case ('outstanding') {
              <app-field label="O/S Date"><app-date-input name="osDate" ariaLabel="O/S Date" [(ngModel)]="f.osDate" /></app-field>
            }
            @case ('control') {
              <app-field label="Subsidiary">
                <select name="subsidiary" [(ngModel)]="f.subsidiary"><option value=""></option>@for (x of subsidiaries(); track x) { <option [value]="x">{{ x }}</option> }</select>
              </app-field>
              <app-field label="Bordx No From"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" /></app-field>
              <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
              <app-field label="Bordx Date From"><app-date-input name="dateFrom" ariaLabel="Bordx Date From" [(ngModel)]="f.dateFrom" /></app-field>
              <app-field label="Bordx Date To"><app-date-input name="dateTo" ariaLabel="Bordx Date To" [(ngModel)]="f.dateTo" /></app-field>
              <app-field label="Sent to Grp. Company">
                <select name="sent" [(ngModel)]="f.sent"><option value=""></option><option value="1">Sent</option><option value="0">Not Sent</option></select>
              </app-field>
              <app-field label="Returned">
                <select name="returned" [(ngModel)]="f.returned"><option value=""></option><option value="1">Paid</option><option value="0">Partial</option></select>
              </app-field>
              <app-field label="Payment">
                <select name="paid" [(ngModel)]="f.paid"><option value=""></option><option value="F">Paid</option><option value="P">Partial</option><option value="U">Unpaid</option></select>
              </app-field>
              <app-field label="Cash Type">
                <select name="cashType" [(ngModel)]="f.cashType"><option value=""></option><option value="C">C - Cashless</option><option value="M">M - Member</option></select>
              </app-field>
              <app-field label="Entity Type">
                <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Health Connect</option><option [ngValue]="1">Sihat Johor</option></select>
              </app-field>
            }
            @default {
              <app-field label="Payor">
                <select name="payor" [(ngModel)]="f.payor"><option value=""></option>@for (p of payors(); track p.code) { <option [value]="p.code + ' - ' + p.name">{{ p.code }} - {{ p.name }}</option> }</select>
              </app-field>
              <app-field label="Bordx No From"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" /></app-field>
              <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
              <app-field label="Worksheet No From"><input name="caseFrom" maxlength="30" [(ngModel)]="f.caseFrom" /></app-field>
              <app-field label="Worksheet No To"><input name="caseTo" maxlength="30" [(ngModel)]="f.caseTo" /></app-field>
              <app-field label="Receipt No From"><input name="receiptFrom" maxlength="30" [(ngModel)]="f.receiptFrom" /></app-field>
              <app-field label="Receipt No To"><input name="receiptTo" maxlength="30" [(ngModel)]="f.receiptTo" /></app-field>
              <app-field label="Date Fr Payor From"><app-date-input name="recFrom" ariaLabel="Date Fr Payor From" [(ngModel)]="f.recFrom" /></app-field>
              <app-field label="Date Fr Payor To"><app-date-input name="recTo" ariaLabel="Date Fr Payor To" [(ngModel)]="f.recTo" /></app-field>
              <app-field label="No Date Fr Payor"><label class="check-field"><input type="checkbox" name="noRecDate" [(ngModel)]="f.noRecDate" /> Date from payor empty</label></app-field>
            }
          }
          <app-action-bar label="Monitoring actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid [caption]="screen.title" [columns]="screen.columns" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">
          Total: <strong>{{ rows().length }}</strong> ·
          @if (kind === 'outstanding') {
            &lt;= 14 Days: <strong>{{ money(result()?.total14 ?? 0) }}</strong> · &gt; 14 Days: <strong>{{ money(result()?.totalOver14 ?? 0) }}</strong> · Total Due:
          } @else {
            Total Amount:
          }
          <strong>{{ money(result()?.total ?? 0) }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class BordxMonitoringPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  protected readonly kind: Kind = inject(ActivatedRoute).snapshot.data['kind']
  protected readonly screen = screens[this.kind]

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly companies = signal<string[]>([])
  protected readonly payors = signal<{ code: string; name: string }[]>([])
  protected readonly subsidiaries = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly result = signal<Result | null>(null)
  protected readonly busy = signal(false)
  protected f = this.empty()
  protected readonly key = (r: Row) => `${s(r['BordxRefNo'])}-${s(r['CASNumber'])}-${s(r['INVPayTo'])}-${s(r['No'])}`

  private empty() {
    return {
      company: '', osDate: formatDate(new Date().toISOString()), subsidiary: '', bordxFrom: '', bordxTo: '', dateFrom: '', dateTo: '', sent: '', returned: '', paid: '',
      cashType: '', entity: 0, payor: '', caseFrom: '', caseTo: '', receiptFrom: '', receiptTo: '', recFrom: '', recTo: '', noRecDate: false,
    }
  }

  async ngOnInit(): Promise<void> {
    await this.lookups()
  }

  private async lookups(company = ''): Promise<void> {
    try {
      const l = await firstValueFrom(
        this.http.get<{ companies: string[]; payors: { code: string; name: string }[]; subsidiaries: string[] }>(`${BASE}/lookups`, { params: { company } }),
      )
      if (!company) {
        this.companies.set(l.companies)
        this.payors.set(l.payors)
      }
      this.subsidiaries.set(l.subsidiaries)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cboGRPCompany_LostFocus: BindSubsidary. */
  protected async companyChanged(): Promise<void> {
    if (this.kind === 'control' && this.f.company.trim()) await this.lookups(this.f.company.trim())
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.result.set(null)
    try {
      const f = this.f
      const r =
        this.kind === 'outstanding'
          ? await firstValueFrom(this.http.get<Result>(`${BASE}/outstanding`, { params: { company: f.company, osDate: f.osDate } }))
          : await firstValueFrom(this.http.post<Result>(`${BASE}/${this.kind}`, { ...f, sent: f.sent || null, returned: f.returned || null, paid: f.paid || null }))
      this.result.set(r)
      this.rows.set(r.rows)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdPrint_Click: the list as CSV. */
  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    const cols = this.screen.columns
    saveCsv([cols.map((c) => c.label), ...this.rows().map((r) => cols.map((c) => (c.value ? c.value(r) : s(r[c.key]))))], `${this.screen.form}.csv`)
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
    this.result.set(null)
  }
}
