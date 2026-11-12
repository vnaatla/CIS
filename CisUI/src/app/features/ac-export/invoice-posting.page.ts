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
import { saveBlob, saveCsv } from '../../shared/lib/download'
import { formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/ac-export/invoice-posting'
type Line = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const x = (k: string, label: string): GridColumn<Line> => ({ key: k, label, value: (r) => s((r['extra'] as Record<string, string> | null)?.[k]) })

/** A/C Export > Invoice Posting (frmInvPosting). */
@Component({
  selector: 'app-invoice-posting-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Invoice Posting" subtitle="Replaces frmInvPosting">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Payor" [wide]="true">
            <select name="payor" [(ngModel)]="f.payor"><option value=""></option>@for (p of payors(); track p.code) { <option [value]="p.code + ' - ' + p.name">{{ p.code }} - {{ p.name }}</option> }</select>
          </app-field>
          <app-field label="Invoice No From"><input name="invoiceFrom" maxlength="50" [(ngModel)]="f.invoiceFrom" /></app-field>
          <app-field label="Invoice No To"><input name="invoiceTo" maxlength="50" [(ngModel)]="f.invoiceTo" /></app-field>
          <app-field label="Register Date From"><app-date-input name="registerFrom" ariaLabel="Register Date From" [(ngModel)]="f.registerFrom" /></app-field>
          <app-field label="Register Date To"><app-date-input name="registerTo" ariaLabel="Register Date To" [(ngModel)]="f.registerTo" /></app-field>
          <app-field label="Post Status">
            <select name="posted" [(ngModel)]="f.posted"><option value=""></option><option value="1 - Yes">1 - Yes</option><option value="0 - No">0 - No</option></select>
          </app-field>
          <app-action-bar label="Posting actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [busy]="posting()" (click)="post()">Post</button>
            <button appButton (click)="print()">Print to File</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Invoice postings" [columns]="columns" [rows]="lines()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">
          Lines: <strong>{{ lines().length }}</strong> · Debit: <strong>{{ money(debit()) }}</strong> · Credit: <strong>{{ money(credit()) }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class InvoicePostingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly payors = signal<{ code: string; name: string }[]>([])
  protected readonly lines = signal<Line[]>([])
  protected readonly debit = signal(0)
  protected readonly credit = signal(0)
  protected readonly busy = signal(false)
  protected readonly posting = signal(false)
  protected f = this.empty()
  protected readonly key = (x: Line) => `${this.lines().indexOf(x)}`
  protected readonly columns: GridColumn<Line>[] = [
    { key: 'account', label: 'Account' }, { key: 'period', label: 'Period' },
    { key: 'amount', label: 'Amount', align: 'right', value: (x) => formatMoney(Number(x['amount'] ?? 0)) }, { key: 'debitCredit', label: 'D/C' },
    { key: 'date', label: 'Date' }, { key: 'journal', label: 'Journal' }, { key: 'reference', label: 'Reference' }, { key: 'description', label: 'Description' },
    { key: 'clinicCode', label: 'Clinic' }, { key: 'payCode', label: 'Payor' }, x('RegPeriod', 'Reg Period'), x('GrpCo', 'Grp Co'),
    x('Bordx', 'Bordx'), x('PostedText', 'Posted'), x('PostUser', 'Post User'), x('PostDate', 'Post Date'),
  ]

  private empty() {
    return { payor: '', invoiceFrom: '', invoiceTo: '', registerFrom: '', registerTo: '', posted: '0 - No' }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ payors: { code: string; name: string }[] }>('/api/payment/post/clinic/lookups'))
      this.payors.set(l.payors)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ lines: Line[]; debit: number; credit: number }>(`${BASE}/search`, this.f))
      this.lines.set(r.lines)
      this.debit.set(r.debit)
      this.credit.set(r.credit)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdPost_Click. */
  protected async post(): Promise<void> {
    if (!(await this.notify.ask('Do you want to post the listed invoices?'))) return
    this.posting.set(true)
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/post`, this.f, { responseType: 'blob' })), 'InvoicePosting.xlsx')
      await this.search()
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    } finally {
      this.posting.set(false)
    }
  }

  protected print(): void {
    if (!this.lines().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.lines().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'InvoicePosting.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.lines.set([])
    this.debit.set(0)
    this.credit.set(0)
  }
}
