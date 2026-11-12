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

type Line = Record<string, unknown>
type Result = { lines: Line[]; debit: number; credit: number; records: number; posted: boolean }
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const extra = (x: Line, k: string) => s((x['extra'] as Record<string, string> | null)?.[k])

/** A/C Export > Bordx Posting > Grp Company (frmPostingBordxtoGrpComp) and Grp Company (GST) (frmPostingBordxtoGrpGST, route variant gst). */
@Component({
  selector: 'app-posting-bordx-group-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="gst ? 'Bordx Posting - Grp Company (GST)' : 'Bordx Posting - Grp Company'" [subtitle]="gst ? 'Replaces frmPostingBordxtoGrpGST' : 'Replaces frmPostingBordxtoGrpComp'">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Group Company" [wide]="true">
            <input name="groupCompany" list="pim-companies" [(ngModel)]="f.groupCompany" />
            <datalist id="pim-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          @if (gst) {
            <app-field label="Entity">
              <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Health Connect</option><option [ngValue]="1">Sihat Johor</option></select>
            </app-field>
          }
          <app-field label="Bordx Date From"><app-date-input name="bordxDateFrom" ariaLabel="Bordx Date From" [(ngModel)]="f.bordxDateFrom" /></app-field>
          <app-field label="Bordx Date To"><app-date-input name="bordxDateTo" ariaLabel="Bordx Date To" [(ngModel)]="f.bordxDateTo" /></app-field>
          <app-field label="Bordx No From"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" /></app-field>
          <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
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

        <app-data-grid caption="Group company bordx postings" [columns]="columns" [rows]="lines()" [rowKey]="key" [checkable]="kind() === 'G'" [checked]="checked()"
          (checkedChange)="check($event)" />
        <p class="totals" aria-label="Totals">
          Records: <strong>{{ records() }}</strong> · Debit: <strong>{{ money(debit()) }}</strong> · Credit: <strong>{{ money(credit()) }}</strong> · Posted:
          <strong>{{ postedList() ? 'Yes' : 'No' }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class PostingBordxGroupPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  protected readonly gst = inject(ActivatedRoute).snapshot.data['variant'] === 'gst'
  private readonly base = this.gst ? '/api/ac-export/bordx-group-gst' : '/api/ac-export/bordx-group'

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly periods = signal<string[]>([])
  protected readonly companies = signal<string[]>([])
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
  /** A bordereaux's lines are ticked together. */
  private readonly item = (x: Line) => s(x['reference'])
  protected readonly key = (x: Line) => `${this.item(x)}|${s(x['account'])}|${s(x['debitCredit'])}|${this.lines().indexOf(x)}`
  protected readonly columns: GridColumn<Line>[] = [
    { key: 'posted', label: 'Post', value: (x) => (x['posted'] ? 'Post' : '') }, { key: 'journal', label: 'Journal' }, { key: 'ledger', label: 'Ledger' },
    { key: 'reference', label: 'Bordx No' }, { key: 'period', label: 'Period' }, { key: 'date', label: 'Bordx Date' }, { key: 'description', label: 'Description' },
    { key: 'account', label: 'Account' }, { key: 'amount', label: 'Amount', align: 'right', value: (x) => formatMoney(Number(x['amount'] ?? 0)) },
    { key: 'debitCredit', label: 'D/C' }, { key: 'payCode', label: 'Payor' },
    { key: 'bordxPeriod', label: 'Bordx Period', value: (x) => extra(x, 'BordxPeriod') }, { key: 'clinicCode', label: 'Grp Co' },
    { key: 'bordx', label: 'Bordx', value: (x) => extra(x, 'Bordx') },
    { key: 'batchNo', label: 'Batch No' },
  ]

  private empty() {
    return {
      entity: 0, groupCompany: '', bordxDateFrom: '', bordxDateTo: '', bordxFrom: '', bordxTo: '', posted: '0 - No', postDateFrom: '', postDateTo: '', batchFrom: '',
      batchTo: '',
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      this.periods.set(await firstValueFrom(this.http.get<string[]>('/api/ac-export/invoice-clinic/periods')))
      this.companies.set(await firstValueFrom(this.http.get<string[]>('/api/ac-export/invoice-member/companies')))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(kind: string): Promise<void> {
    this.busy.set(true)
    this.batchNo.set('')
    try {
      this.show(await firstValueFrom(this.http.post<Result>(`${this.base}/search`, { ...this.f, kind })), kind)
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
    if (!this.f.bordxDateFrom || !this.f.bordxDateTo) return this.notify.notify('Please Enter Bordx Date! ', 'error')
    if (!this.period) return this.notify.notify('Please Keyin for Accounting Period ! ', 'error')
    if (!(await this.notify.ask('Do you want to Post record(s) ?'))) return
    this.posting.set(true)
    try {
      const body = { search: this.f, period: this.period, ticked: ticked.map((x) => s(x['reference'])) }
      const res: HttpResponse<Blob> = await firstValueFrom(this.http.post(`${this.base}/post`, body, { responseType: 'blob', observe: 'response' }))
      const batch = res.headers.get('X-Batch-No') ?? ''
      saveBlob(res.body!, `${batch} BI-BordxGrp.xlsx`)
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
      saveBlob(await firstValueFrom(this.http.post(`${this.base}/export`, { ...this.f, kind: this.kind() }, { responseType: 'blob' })), 'PostingBordxGroup.xlsx')
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
