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

const BASE = '/api/claim-bordx/delivery'
type Row = Record<string, unknown>
type Kind = 'payor' | 'returned' | 'ac' | 'mnrb' | 'acperiod'
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const flag = (v: unknown) => (v === true || v === 1 || v === '1' ? 'True' : 'False')

interface Screen {
  title: string
  form: string
  update: string
  dateLabel: string
  periods: boolean
  status: boolean
  columns: GridColumn<Row>[]
}

const common: GridColumn<Row>[] = [
  { key: 'BordxPAYCode', label: 'Payor' }, { key: 'GroupCompany', label: 'Group Company' }, { key: 'BordxRefNo', label: 'Bordx No' },
  { key: 'BordxDate', label: 'Bordx Date', value: (r) => d(r['BordxDate']) },
  { key: 'BordxClaimAmt', label: 'Amount', align: 'right', value: (r) => (r['BordxClaimAmt'] == null ? '' : formatMoney(Number(r['BordxClaimAmt']))) },
]
const period: GridColumn<Row> = { key: 'period', label: 'A/C Period', value: (r) => `${s(r['PAYToACPeriodMth'])}/${s(r['PAYToACPeriodYR'])}` }
const screens: Record<Kind, Screen> = {
  payor: {
    title: 'Send Bordereaux to Group Company', form: 'FrmSendtoPayor', update: 'To Grp. Company', dateLabel: 'Date Sent', periods: true, status: true,
    columns: [...common, { key: 'BordxPAYToStatus', label: 'Sent', value: (r) => flag(r['BordxPAYToStatus']) },
      { key: 'BordxPAYSentDate', label: 'Date Sent', value: (r) => d(r['BordxPAYSentDate']) }, period],
  },
  returned: {
    title: 'Bordereaux Returned from Group Company', form: 'FrmReturnPayor', update: 'Fr Grp. Company', dateLabel: 'Date Received', periods: true, status: true,
    columns: [...common, { key: 'BordxPAYDOCFRStatus', label: 'Returned', value: (r) => flag(r['BordxPAYDOCFRStatus']) },
      { key: 'BordxPAYRecDocDate', label: 'Date Received', value: (r) => d(r['BordxPAYRecDocDate']) }, period],
  },
  ac: {
    title: 'Send Bordereaux to A/C', form: 'FrmSendtoAC', update: 'To A/C', dateLabel: 'Date Sent to A/C', periods: false, status: false,
    columns: [...common, { key: 'BordxPAYSentDate', label: 'Sent to Payor', value: (r) => d(r['BordxPAYSentDate']) },
      { key: 'BordxPAYRecDocDate', label: 'Received', value: (r) => d(r['BordxPAYRecDocDate']) },
      { key: 'BordxMNRBSentDate', label: 'Sent to MNRB', value: (r) => d(r['BordxMNRBSentDate']) },
      { key: 'BordxMNRBRecDocDate', label: 'MNRB Received', value: (r) => d(r['BordxMNRBRecDocDate']) },
      { key: 'BordxSendACDate', label: 'Sent to A/C', value: (r) => d(r['BordxSendACDate']) },
      { key: 'BordxSendToACStatus', label: 'A/C Status', value: (r) => flag(r['BordxSendToACStatus']) }],
  },
  acperiod: {
    title: 'Bordereaux Accounting Period', form: 'FrmACPeriod', update: 'Update', dateLabel: '', periods: false, status: false,
    columns: [...common, { key: 'BordxPAYSentDate', label: 'Sent to Payor', value: (r) => d(r['BordxPAYSentDate']) },
      { key: 'BordxPAYRecDate', label: 'Received', value: (r) => d(r['BordxPAYRecDate']) },
      { key: 'BordxMNRBSentDate', label: 'Sent to MNRB', value: (r) => d(r['BordxMNRBSentDate']) }, period,
      { key: 'ACPeriodStatus', label: 'A/C Period Status', value: (r) => flag(s(r['ACPeriodStatus']) === 'True' ? 1 : r['ACPeriodStatus']) }],
  },
  mnrb: {
    title: 'Send Bordereaux to MNRB', form: 'FrmSendtoMNRB', update: 'To MNRB', dateLabel: 'Date Sent', periods: true, status: true,
    columns: [{ key: 'BordxPAYCode', label: 'Payor' }, { key: 'BordxRefNo', label: 'Bordx No' }, { key: 'BordxDate', label: 'Bordx Date', value: (r) => d(r['BordxDate']) },
      common[4], { key: 'BordxMNRBToStatus', label: 'Sent', value: (r) => flag(r['BordxMNRBToStatus']) },
      { key: 'BordxMNRBSentDate', label: 'Date Sent', value: (r) => d(r['BordxMNRBSentDate']) }, period],
  },
}

/** Claim Bordx > Bordx Delivery: Send to Grp. Company, Returned Fr Grp. Company, Send to A/C, Send to MNRB (route data `kind`). */
@Component({
  selector: 'app-bordx-delivery-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="screen.title" [subtitle]="'Replaces ' + screen.form">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          @if (kind === 'mnrb') {
            <app-field label="Payor">
              <select name="company" [(ngModel)]="f.company"><option value=""></option>@for (p of payors(); track p.code) { <option [value]="p.code + ' - ' + p.name">{{ p.code }} - {{ p.name }}</option> }</select>
            </app-field>
          } @else {
            <app-field label="Group Company" [wide]="true">
              <input name="company" list="delivery-companies" [(ngModel)]="f.company" />
              <datalist id="delivery-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
            </app-field>
          }
          <app-field label="Scan"><label class="check-field"><input type="checkbox" name="scan" [(ngModel)]="f.scan" (ngModelChange)="clearList()" /> Scan bordx no.</label></app-field>
          <app-field [label]="f.scan ? 'Bordx No' : 'Bordx No From'"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" (keydown.enter)="search()" /></app-field>
          @if (!f.scan) {
            <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
          }
          <app-field label="Bordx Date From"><app-date-input name="dateFrom" ariaLabel="Bordx Date From" [(ngModel)]="f.dateFrom" /></app-field>
          <app-field label="Bordx Date To"><app-date-input name="dateTo" ariaLabel="Bordx Date To" [(ngModel)]="f.dateTo" /></app-field>
          <app-field label="Amount From"><input class="num" name="amountFrom" maxlength="20" [(ngModel)]="f.amountFrom" /></app-field>
          <app-field label="Amount To"><input class="num" name="amountTo" maxlength="20" [(ngModel)]="f.amountTo" /></app-field>
          @if (screen.periods) {
            <app-field label="A/C Month From"><input name="monthFrom" maxlength="4" [(ngModel)]="f.monthFrom" /></app-field>
            <app-field label="A/C Month To"><input name="monthTo" maxlength="4" [(ngModel)]="f.monthTo" /></app-field>
            <app-field label="A/C Year From"><input name="yearFrom" maxlength="6" [(ngModel)]="f.yearFrom" /></app-field>
            <app-field label="A/C Year To"><input name="yearTo" maxlength="6" [(ngModel)]="f.yearTo" /></app-field>
          } @else if (kind === 'acperiod') {
            <app-field label="A/C Period From (mm/yyyy)"><input name="periodFrom" maxlength="7" placeholder="mm/yyyy" [(ngModel)]="f.periodFrom" /></app-field>
            <app-field label="A/C Period To (mm/yyyy)"><input name="periodTo" maxlength="7" placeholder="mm/yyyy" [(ngModel)]="f.periodTo" /></app-field>
            <app-field label="No A/C Period"><label class="check-field"><input type="checkbox" name="noPeriod" [(ngModel)]="f.noPeriod" /> Without A/C period</label></app-field>
          } @else {
            <app-field label="Date Sent From"><app-date-input name="sentFrom" ariaLabel="Date Sent From" [(ngModel)]="f.sentFrom" /></app-field>
            <app-field label="Date Sent To"><app-date-input name="sentTo" ariaLabel="Date Sent To" [(ngModel)]="f.sentTo" /></app-field>
          }
          @if (screen.status) {
            <app-field label="Status">
              <select name="status" [(ngModel)]="f.status"><option [ngValue]="0">0 - Not Sent</option><option [ngValue]="1">1 - Sent</option><option [ngValue]="2">2 - Both</option></select>
            </app-field>
          }
          @if (kind === 'acperiod') {
            <app-field label="Period Month" [required]="true">
              <select name="month" [(ngModel)]="month"><option value=""></option>@for (m of months; track m) { <option [value]="m">{{ m }}</option> }</select>
            </app-field>
            <app-field label="Period Year" [required]="true">
              <select name="year" [(ngModel)]="year"><option value=""></option>@for (y of years; track y) { <option [value]="y">{{ y }}</option> }</select>
            </app-field>
          } @else {
            <app-field [label]="screen.dateLabel" [required]="true"><app-date-input name="date" [ariaLabel]="screen.dateLabel" [(ngModel)]="date" /></app-field>
          }
          <app-action-bar label="Delivery actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [disabled]="!rows().length" (click)="checkAll()">Check All</button>
            <button appButton variant="primary" (click)="update()">{{ screen.update }}</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Bordereaux" [columns]="screen.columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()" (checkedChange)="check($event)" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong> · Total Amount: <strong>{{ money(total()) }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class BordxDeliveryPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  protected readonly kind: Kind = inject(ActivatedRoute).snapshot.data['kind']
  protected readonly screen = screens[this.kind]

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly companies = signal<string[]>([])
  protected readonly payors = signal<{ code: string; name: string }[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly total = signal(0)
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly busy = signal(false)
  protected date = ''
  protected month = ''
  protected year = ''
  protected readonly months = Array.from({ length: 12 }, (_, i) => String(i + 1))
  protected readonly years = Array.from({ length: 3000 - 1999 + 1 }, (_, i) => String(1999 + i))
  protected f = this.empty()
  private searched = this.empty()
  protected readonly key = (r: Row) => s(r['BordxRefNo'])

  private empty() {
    return {
      company: '', scan: false, bordxFrom: '', bordxTo: '', dateFrom: '', dateTo: '', amountFrom: '', amountTo: '', monthFrom: '', monthTo: '', yearFrom: '', yearTo: '',
      status: 0 as number | null, sentFrom: '', sentTo: '', periodFrom: '', periodTo: '', noPeriod: false,
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ companies: string[]; payors: { code: string; name: string }[] }>(`${BASE}/lookups`))
      this.companies.set(l.companies)
      this.payors.set(l.payors)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected clearList(): void {
    this.rows.set([])
    this.total.set(0)
    this.checked.set(new Set())
  }

  /** cmdSearch_Click: scan mode appends one bordereaux. */
  protected async search(): Promise<void> {
    this.busy.set(true)
    try {
      const body = { ...this.f, status: this.screen.status ? this.f.status : null }
      const r = await firstValueFrom(this.http.post<{ rows: Row[]; totalAmount: number }>(`${BASE}/${this.kind}/search`, body))
      this.searched = { ...this.f }
      const rows = this.f.scan ? [...this.rows(), ...r.rows.filter((x) => !this.rows().some((y) => this.key(y) === this.key(x)))] : r.rows
      if (!this.f.scan) this.checked.set(new Set())
      this.rows.set(rows)
      this.total.set(rows.reduce((t, x) => t + Number(x['BordxClaimAmt'] ?? 0), 0))
      if (this.f.scan) this.f.bordxFrom = ''
    } catch (e) {
      if (!this.f.scan) this.clearList()
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstBordxDelivery_ItemCheck: a bordereaux already done cannot be ticked. */
  protected check(keys: Set<string>): void {
    const before = this.checked()
    let refused = false
    for (const r of this.rows()) {
      const k = this.key(r)
      if (keys.has(k) && !before.has(k) && Number(r['Done']) === 1) {
        keys.delete(k)
        refused = true
      }
    }
    if (refused) this.notify.notify('Record has been Updated!', 'error')
    this.checked.set(keys)
  }

  protected checkAll(): void {
    this.check(new Set(this.rows().map(this.key)))
  }

  /** CmdToPayor_Click and the other update buttons. */
  protected async update(): Promise<void> {
    if (!this.checked().size) return this.notify.notify('Please select at least one record to be updated! ', 'error')
    if (this.kind === 'acperiod') {
      if (!this.month) return this.notify.notify('Month cannot be Empty.', 'error')
      if (!this.year) return this.notify.notify('Year cannot be empty!', 'error')
    } else if (!this.date) return this.notify.notify('Please Enter in date', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    try {
      const body = { bordxNos: [...this.checked()], date: this.date, month: this.month, year: this.year }
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/${this.kind}/update`, body))
      this.notify.notify(r.message, r.message === 'Record updated...' ? 'success' : 'error')
      this.clearList()
      if (!this.searched.scan) {
        this.f = { ...this.searched }
        await this.search()
      }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdPrint_Click: the list as CSV. */
  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    const cols = this.screen.columns
    saveCsv([cols.map((c) => c.label), ...this.rows().map((r) => cols.map((c) => (c.value ? c.value(r) : s(r[c.key]))))], `Bordx_${this.kind}.csv`)
  }

  protected clear(): void {
    this.f = this.empty()
    this.date = ''
    this.clearList()
  }
}
