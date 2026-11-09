import { HttpClient, HttpResponse } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { ActivatedRoute } from '@angular/router'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const CLINIC = '/api/ac-export/invoice-clinic'
type Line = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())

/** A/C Export > Invoice Posting > Clinic (frmPostingInvoiceCln). */
@Component({
  selector: 'app-posting-invoice-clinic-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="title" [subtitle]="'Replaces ' + form">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Clinic" [wide]="true"><input name="clinic" maxlength="300" placeholder="code - name" [(ngModel)]="f.clinic" /></app-field>
          <app-field label="Invoice Date From"><app-date-input name="invDateFrom" ariaLabel="Invoice Date From" [(ngModel)]="f.invDateFrom" /></app-field>
          <app-field label="Invoice Date To"><app-date-input name="invDateTo" ariaLabel="Invoice Date To" [(ngModel)]="f.invDateTo" /></app-field>
          <app-field label="Invoice No From"><input name="invoiceFrom" maxlength="100" [(ngModel)]="f.invoiceFrom" /></app-field>
          <app-field label="Invoice No To"><input name="invoiceTo" maxlength="100" [(ngModel)]="f.invoiceTo" /></app-field>
          @if (gst) {
            <app-field label="SUN Invoice No From"><input name="sunFrom" maxlength="20" [(ngModel)]="f.sunFrom" /></app-field>
            <app-field label="SUN Invoice No To"><input name="sunTo" maxlength="20" [(ngModel)]="f.sunTo" /></app-field>
          } @else {
            <app-field label="Bordx Date From"><app-date-input name="bordxDateFrom" ariaLabel="Bordx Date From" [(ngModel)]="f.bordxDateFrom" /></app-field>
            <app-field label="Bordx Date To"><app-date-input name="bordxDateTo" ariaLabel="Bordx Date To" [(ngModel)]="f.bordxDateTo" /></app-field>
            <app-field label="PV Date From"><app-date-input name="pvDateFrom" ariaLabel="PV Date From" [(ngModel)]="f.pvDateFrom" /></app-field>
            <app-field label="PV Date To"><app-date-input name="pvDateTo" ariaLabel="PV Date To" [(ngModel)]="f.pvDateTo" /></app-field>
          }
          <app-field label="Post Status">
            <select name="posted" [(ngModel)]="f.posted"><option value="0 - No">0 - No</option><option value="1 - Yes">1 - Yes</option></select>
          </app-field>
          <app-field label="Post Date From"><app-date-input name="postDateFrom" ariaLabel="Post Date From" [(ngModel)]="f.postDateFrom" /></app-field>
          <app-field label="Post Date To"><app-date-input name="postDateTo" ariaLabel="Post Date To" [(ngModel)]="f.postDateTo" /></app-field>
          <app-field label="Batch No From"><input name="batchFrom" maxlength="20" [(ngModel)]="f.batchFrom" /></app-field>
          <app-field label="Batch No To"><input name="batchTo" maxlength="20" [(ngModel)]="f.batchTo" /></app-field>
          <app-field label="A/C Period (posted)">
            <select name="invoicePeriod" [(ngModel)]="f.invoicePeriod"><option value=""></option>@for (p of periods(); track p) { <option [value]="p">{{ p }}</option> }</select>
          </app-field>
          <app-action-bar label="Posting search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search('G')">Search</button>
            <button appButton (click)="search('D')">Search Debit</button>
            <button appButton (click)="search('C')">Search Credit</button>
            <button appButton (click)="checkAll()">Check All</button>
            <button appButton (click)="print()">Print to File</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Post">
          <legend>Post</legend>
          <app-field label="Accounting Period" [required]="true">
            <select name="period" [(ngModel)]="period"><option value=""></option>@for (p of periods(); track p) { <option [value]="p">{{ p }}</option> }</select>
          </app-field>
          <app-field label="Batch No"><input name="batchNo" [value]="batchNo()" disabled /></app-field>
          <app-action-bar label="Post actions">
            <button appButton variant="primary" [busy]="posting()" (click)="post()">Post</button>
            @if (y2017) {
              <button appButton (click)="taxInvoice()">Generate Tax Invoice</button>
            }
          </app-action-bar>
        </fieldset>

        <app-data-grid [caption]="caption" [columns]="columns" [rows]="lines()" [rowKey]="key" [checkable]="kind() === 'G'" [checked]="checked()"
          (checkedChange)="check($event)" [selectable]="true" [selectedKey]="selectedKey()" (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">
          Records: <strong>{{ records() }}</strong> · Debit: <strong>{{ money(debit()) }}</strong> · Credit: <strong>{{ money(credit()) }}</strong> · Posted:
          <strong>{{ postedList() ? 'Yes' : 'No' }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class PostingInvoiceClinicPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  /** Invoice Posting > Clinic (GST) and (GST-2017) use the same screen (route data variant 'gst' / 'gst2017'). */
  private readonly variant = inject(ActivatedRoute).snapshot.data['variant'] as string | undefined
  protected readonly gst = this.variant === 'gst' || this.variant === 'gst2017' || this.variant === 'zeroised'
  protected readonly y2017 = this.variant === 'gst2017' || this.variant === 'zeroised'
  private readonly suffix = this.variant === 'zeroised' ? ' (GST-2017 Zeroised)' : this.y2017 ? ' (GST-2017)' : this.gst ? ' (GST)' : ''
  private readonly base = CLINIC + (({ gst: '-gst', gst2017: '-gst-2017', zeroised: '-gst-2017-zeroised' } as Record<string, string>)[this.variant ?? ''] ?? '')
  protected readonly title = 'Invoice Posting - Clinic' + this.suffix
  protected readonly form = this.variant === 'zeroised' ? 'frmPostingInvClnGST2017Zeroised' : this.y2017 ? 'frmPostingInvClnGST2017' : this.gst ? 'frmPostingInvoiceClnGST' : 'frmPostingInvoiceCln'
  protected readonly caption = 'Clinic invoice postings' + this.suffix
  protected readonly selectedKey = signal<string | null>(null)
  private selected: Line | null = null

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly periods = signal<string[]>([])
  protected readonly lines = signal<Line[]>([])
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly kind = signal('G')
  protected readonly debit = signal(0)
  protected readonly credit = signal(0)
  protected readonly records = signal(0)
  protected readonly postedList = signal(false)
  protected readonly batchNo = signal('')
  protected readonly busy = signal(false)
  protected readonly posting = signal(false)
  protected period = ''
  protected f = this.empty()
  /** A line's invoice: ticking one line of an invoice ticks its debit and credit lines together. */
  private readonly invoice = (x: Line) => `${s(x['reference'])}|${s(x['clinicCode'])}`
  protected readonly key = (x: Line) => `${this.invoice(x)}|${s(x['debitCredit'])}`
  protected readonly columns: GridColumn<Line>[] = [
    { key: 'posted', label: 'Post', value: (x) => (x['posted'] ? 'Post' : '') }, { key: 'journal', label: 'Journal' }, { key: 'ledger', label: 'Ledger' },
    { key: 'batchNo', label: 'Batch No' }, { key: 'period', label: 'Period' }, { key: 'date', label: 'Date' }, { key: 'reference', label: 'Invoice No' },
    { key: 'account', label: 'Account' }, { key: 'amount', label: 'Amount', align: 'right', value: (x) => formatMoney(Number(x['amount'] ?? 0)) },
    { key: 'debitCredit', label: 'D/C' }, { key: 'analysis', label: 'Analysis' }, { key: 'payCode', label: 'Payor' }, { key: 'clinicCode', label: 'Clinic' },
    { key: 'description', label: 'Description' },
  ]

  private empty() {
    return {
      clinic: '', invDateFrom: '', invDateTo: '', invoiceFrom: '', invoiceTo: '', bordxDateFrom: '', bordxDateTo: '', pvDateFrom: '', pvDateTo: '', posted: '0 - No',
      postDateFrom: '', postDateTo: '', batchFrom: '', batchTo: '', invoicePeriod: '', sunFrom: '', sunTo: '',
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      this.periods.set(await firstValueFrom(this.http.get<string[]>(`${CLINIC}/periods`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(kind: string): Promise<void> {
    this.busy.set(true)
    this.batchNo.set('')
    try {
      this.show(await firstValueFrom(this.http.post<{ lines: Line[]; debit: number; credit: number; records: number; posted: boolean }>(`${this.base}/search`, { ...this.f, kind })), kind)
    } catch (e) {
      this.show({ lines: [], debit: 0, credit: 0, records: 0, posted: false }, kind)
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private show(r: { lines: Line[]; debit: number; credit: number; records: number; posted: boolean }, kind: string): void {
    this.kind.set(kind)
    this.lines.set(r.lines)
    this.debit.set(r.debit)
    this.credit.set(r.credit)
    this.records.set(r.records)
    this.postedList.set(r.posted)
    this.checked.set(new Set())
  }

  /** lstClnInvList_ItemCheck: a posted list cannot be ticked; an invoice's lines are ticked together. */
  protected check(keys: Set<string>): void {
    if (this.postedList()) {
      this.notify.notify('Record(s) has been posted ! ', 'error')
      this.checked.set(new Set(this.checked()))
      return
    }
    const before = this.checked()
    const invoices = new Set<string>()
    for (const x of this.lines()) {
      const k = this.key(x)
      if (keys.has(k) !== before.has(k)) invoices.add(this.invoice(x))
    }
    const next = new Set(before)
    for (const x of this.lines())
      if (invoices.has(this.invoice(x))) {
        const on = this.lines().some((y) => this.invoice(y) === this.invoice(x) && keys.has(this.key(y)) && !before.has(this.key(y)))
        if (on) next.add(this.key(x))
        else next.delete(this.key(x))
      }
    this.checked.set(next)
  }

  /** CheckAll_Click: toggles every line. */
  protected checkAll(): void {
    if (this.postedList() || !this.lines().length) return
    const all = this.lines().every((x) => this.checked().has(this.key(x)))
    this.checked.set(all ? new Set() : new Set(this.lines().map(this.key)))
  }

  /** CmdPost_Click. */
  protected async post(): Promise<void> {
    const ticked = this.lines().filter((x) => x['postLine'] && this.checked().has(this.key(x)))
    if (this.kind() !== 'G' || !ticked.length) return this.notify.notify('Please select the record', 'error')
    if (this.postedList()) return this.notify.notify('Record(s) has been posted ! ', 'error')
    if (this.gst && (!this.f.invDateFrom || !this.f.invDateTo)) return this.notify.notify('Please Enter Invoice Date! ', 'error')
    if (!this.gst && (!this.f.bordxDateFrom || !this.f.bordxDateTo)) return this.notify.notify('Please Enter Bordx Date! ', 'error')
    if (!this.period) return this.notify.notify('Please Select Accounting Period ! ', 'error')
    if (!(await this.notify.ask('Do you want to Post record(s) ?'))) return
    this.posting.set(true)
    try {
      const body = { search: this.f, period: this.period, ticked: ticked.map((x) => ({ invoiceNo: s(x['reference']), clinicCode: s(x['clinicCode']), invoiceDate: x['invoiceDate'] })) }
      const res: HttpResponse<Blob> = await firstValueFrom(this.http.post(`${this.base}/post`, body, { responseType: 'blob', observe: 'response' }))
      const batch = res.headers.get('X-Batch-No') ?? ''
      saveBlob(res.body!, `${batch} RC-CLNInv.xlsx`)
      this.batchNo.set(batch)
      this.f = { ...this.empty(), posted: '1 - Yes', batchFrom: batch, batchTo: batch }
      await this.search('G')
      this.batchNo.set(batch)
      this.notify.notify('Record(s) Posted Sucessfully... ! ', 'success')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    } finally {
      this.posting.set(false)
    }
  }

  protected select(x: Line): void {
    this.selected = x
    this.selectedKey.set(this.key(x))
  }

  /** cmdgenerate_Click (GST-2017): the FFS tax invoice of the selected line's SUN invoice batch. */
  protected async taxInvoice(): Promise<void> {
    if (!this.lines().length) return this.notify.notify('Please select the record', 'error')
    const batch = s((this.selected ?? this.lines()[0])['batchNo'])
    if (!batch) return this.notify.notify('Invoice posting is not done,Generate is not possible', 'error')
    try {
      saveBlob(await firstValueFrom(this.http.get(`${this.base}/tax-invoice`, { params: { invBatch: batch }, responseType: 'blob' })), `${batch}.xlsx`)
      this.notify.notify('Tax Invoice Generated Successfully', 'success')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  /** CmdPrintFile_Click. */
  protected async print(): Promise<void> {
    if (!this.lines().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    try {
      saveBlob(await firstValueFrom(this.http.post(`${this.base}/export`, { ...this.f, kind: this.kind() }, { responseType: 'blob' })), 'PostingInvoiceClinic.xlsx')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  protected clear(): void {
    this.f = this.empty()
    this.period = ''
    this.batchNo.set('')
    this.show({ lines: [], debit: 0, credit: 0, records: 0, posted: false }, 'G')
  }
}
