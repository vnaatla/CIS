import { HttpClient, HttpResponse } from '@angular/common/http'
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
import { formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/ac-export/payment-clinic'
type Line = Record<string, unknown>
type Result = { lines: Line[]; debit: number; credit: number; records: number; posted: boolean }
const s = (v: unknown) => (v == null ? '' : String(v).trim())

/** A/C Export > Payment Posting > Clinic (FrmPostingPay2Cln). */
@Component({
  selector: 'app-posting-payment-clinic-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Payment Posting - Clinic" subtitle="Replaces FrmPostingPay2Cln">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Entity">
            <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Health Connect</option><option [ngValue]="1">Sihat Johor</option></select>
          </app-field>
          <app-field label="Clinic" [wide]="true"><input name="clinic" maxlength="300" placeholder="code - name" [(ngModel)]="f.clinic" /></app-field>
          <app-field label="PV Date From"><app-date-input name="pvDateFrom" ariaLabel="PV Date From" [(ngModel)]="f.pvDateFrom" /></app-field>
          <app-field label="PV Date To"><app-date-input name="pvDateTo" ariaLabel="PV Date To" [(ngModel)]="f.pvDateTo" /></app-field>
          <app-field label="PV No From"><input name="pvFrom" maxlength="30" [(ngModel)]="f.pvFrom" /></app-field>
          <app-field label="PV No To"><input name="pvTo" maxlength="30" [(ngModel)]="f.pvTo" /></app-field>
          <app-field label="Post Status">
            <select name="posted" [(ngModel)]="f.posted"><option value="0 - No">0 - No</option><option value="1 - Yes">1 - Yes</option></select>
          </app-field>
          <app-field label="Post Date From"><app-date-input name="postDateFrom" ariaLabel="Post Date From" [(ngModel)]="f.postDateFrom" /></app-field>
          <app-field label="Post Date To"><app-date-input name="postDateTo" ariaLabel="Post Date To" [(ngModel)]="f.postDateTo" /></app-field>
          <app-field label="Batch No From"><input name="batchFrom" maxlength="20" [(ngModel)]="f.batchFrom" /></app-field>
          <app-field label="Batch No To"><input name="batchTo" maxlength="20" [(ngModel)]="f.batchTo" /></app-field>
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
          <app-action-bar label="Post actions"><button appButton variant="primary" [busy]="posting()" (click)="post()">Post</button></app-action-bar>
        </fieldset>

        <app-data-grid caption="Clinic payment postings" [columns]="columns" [rows]="lines()" [rowKey]="key" [checkable]="kind() === 'G'" [checked]="checked()"
          (checkedChange)="check($event)" />
        <p class="totals" aria-label="Totals">
          Records: <strong>{{ records() }}</strong> · Debit: <strong>{{ money(debit()) }}</strong> · Credit: <strong>{{ money(credit()) }}</strong> · Posted:
          <strong>{{ postedList() ? 'Yes' : 'No' }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class PostingPaymentClinicPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

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
  /** A PV's lines are ticked together. */
  private readonly item = (x: Line) => s(x['reference'])
  protected readonly key = (x: Line) => `${this.item(x)}|${s(x['account'])}|${s(x['debitCredit'])}|${this.lines().indexOf(x)}`
  protected readonly columns: GridColumn<Line>[] = [
    { key: 'posted', label: 'Post', value: (x) => (x['posted'] ? 'Post' : '') }, { key: 'journal', label: 'Journal' }, { key: 'ledger', label: 'Ledger' },
    { key: 'reference', label: 'PV No' }, { key: 'period', label: 'Period' }, { key: 'date', label: 'PV Date' }, { key: 'description', label: 'Description' },
    { key: 'account', label: 'Account' }, { key: 'amount', label: 'Amount', align: 'right', value: (x) => formatMoney(Number(x['amount'] ?? 0)) },
    { key: 'debitCredit', label: 'D/C' }, { key: 'analysis', label: 'Analysis' }, { key: 'payCode', label: 'Payor' },
    { key: 'clinicCode', label: 'Clinic' }, { key: 'batchNo', label: 'Batch No' },
  ]

  private empty() {
    return {
      entity: 0, clinic: '', pvDateFrom: '', pvDateTo: '', pvFrom: '', pvTo: '', posted: '0 - No', postDateFrom: '', postDateTo: '', batchFrom: '', batchTo: '',
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      this.periods.set(await firstValueFrom(this.http.get<string[]>('/api/ac-export/invoice-clinic/periods')))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(kind: string): Promise<void> {
    this.busy.set(true)
    this.batchNo.set('')
    try {
      this.show(await firstValueFrom(this.http.post<Result>(`${BASE}/search`, { ...this.f, kind })), kind)
    } catch (e) {
      this.show({ lines: [], debit: 0, credit: 0, records: 0, posted: false }, kind)
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private show(r: Result, kind: string): void {
    this.kind.set(kind)
    this.lines.set(r.lines)
    this.debit.set(r.debit)
    this.credit.set(r.credit)
    this.records.set(r.records)
    this.postedList.set(r.posted)
    this.checked.set(new Set())
  }

  protected check(keys: Set<string>): void {
    if (this.postedList()) {
      this.notify.notify('Record(s) has been posted ! ', 'error')
      this.checked.set(new Set(this.checked()))
      return
    }
    const before = this.checked()
    const next = new Set(before)
    for (const x of this.lines()) {
      const k = this.key(x)
      if (keys.has(k) === before.has(k)) continue
      for (const y of this.lines().filter((y) => this.item(y) === this.item(x)))
        if (keys.has(k)) next.add(this.key(y))
        else next.delete(this.key(y))
    }
    this.checked.set(next)
  }

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
    if (!this.f.pvDateFrom || !this.f.pvDateTo) return this.notify.notify('Please Enter PV Date! ', 'error')
    if (!this.period) return this.notify.notify('Please Keyin for Accounting Period ! ', 'error')
    if (!(await this.notify.ask('Do you want to Post record(s) ?'))) return
    this.posting.set(true)
    try {
      const body = { search: this.f, period: this.period, ticked: ticked.map((x) => ({ pvNo: s(x['reference']), analysis: s(x['analysis']) })) }
      const res: HttpResponse<Blob> = await firstValueFrom(this.http.post(`${BASE}/post`, body, { responseType: 'blob', observe: 'response' }))
      const batch = res.headers.get('X-Batch-No') ?? ''
      saveBlob(res.body!, `${batch} CP-PayCLN.xlsx`)
      this.f = { ...this.empty(), entity: this.f.entity, posted: '1 - Yes', batchFrom: batch, batchTo: batch }
      await this.search('G')
      this.batchNo.set(batch)
      this.notify.notify('Record(s) Posted Sucessfully... ! ', 'success')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    } finally {
      this.posting.set(false)
    }
  }

  protected async print(): Promise<void> {
    if (!this.lines().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/export`, { ...this.f, kind: this.kind() }, { responseType: 'blob' })), 'PostingPaymentClinic.xlsx')
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
