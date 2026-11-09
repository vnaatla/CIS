import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/claim-bordx/insert-invoice'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())

/** Claim Bordx > Worksheet > Insert Invoice (frmWSBordx2). */
@Component({
  selector: 'app-insert-invoice-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Insert Invoice" subtitle="Replaces frmWSBordx2">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Register Date From"><app-date-input name="registerFrom" ariaLabel="Register Date From" [(ngModel)]="f.registerFrom" /></app-field>
          <app-field label="Register Date To"><app-date-input name="registerTo" ariaLabel="Register Date To" [(ngModel)]="f.registerTo" /></app-field>
          <app-field label="Group Company" [wide]="true">
            <input name="groupCompany" list="ii-companies" [(ngModel)]="f.groupCompany" (change)="company()" />
            <datalist id="ii-companies">@for (c of companies(); track c.code) { <option [value]="c.code + ' - ' + c.name"></option> }</datalist>
          </app-field>
          <app-field label="Payor">
            <select name="payor" [(ngModel)]="f.payor"><option value=""></option>@for (p of payors(); track p.code) { <option [value]="p.code + ' - ' + p.name">{{ p.code }} - {{ p.name }}</option> }</select>
          </app-field>
          <app-action-bar label="Insert invoice actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [disabled]="!rows().length" (click)="checkAll()">Check All</button>
            <button appButton variant="primary" [busy]="saving()" (click)="bordx()">Bordx</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Invoices pending bordx" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="check($event)" />
        <p class="totals" aria-label="Totals">
          Total: <strong>{{ rows().length }}</strong> · Total Approved: <strong>{{ money(total()) }}</strong>
          @if (bordxNo()) { · Bordx No: <strong>{{ bordxNo() }}</strong> }
        </p>
      </div>
    </app-page-shell>
  `,
})
export class InsertInvoicePage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly companies = signal<{ code: string; name: string; payCode: string | null }[]>([])
  protected readonly payors = signal<{ code: string; name: string }[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly total = signal(0)
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly bordxNo = signal('')
  protected f = { registerFrom: '', registerTo: '', groupCompany: '', payor: '' }
  protected readonly key = (x: Row) => [x['InvoiceNo'], x['InvoiceDate'], x['RegisterDate'], x['GRPCoCode']].map(s).join('|')
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'GRPCoPaycode', label: 'Payor' }, { key: 'InvoiceNo', label: 'Invoice No' },
    { key: 'InvoiceDate', label: 'Invoice Date', value: (x) => (x['InvoiceDate'] ? formatDate(String(x['InvoiceDate'])) : '') },
    { key: 'GRPCoName', label: 'Group Company' },
    { key: 'CLMTotalApproved', label: 'Approved', align: 'right', value: (x) => (x['CLMTotalApproved'] == null ? '' : formatMoney(Number(x['CLMTotalApproved']))) },
  ]

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ companies: { code: string; name: string; payCode: string | null }[]; payors: { code: string; name: string }[] }>(`${BASE}/lookups`))
      this.companies.set(l.companies)
      this.payors.set(l.payors)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CboGrpCompany_LostFocus: the payor follows the group company. */
  protected company(): void {
    const code = this.f.groupCompany.split(' - ')[0].trim()
    const c = this.companies().find((x) => x.code === code)
    const p = c?.payCode ? this.payors().find((x) => x.code === c.payCode) : undefined
    if (p) this.f.payor = `${p.code} - ${p.name}`
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.bordxNo.set('')
    this.rows.set([])
    this.total.set(0)
    this.checked.set(new Set())
    try {
      const r = await firstValueFrom(this.http.post<{ rows: Row[]; totalApproved: number }>(`${BASE}/search`, this.f))
      this.rows.set(r.rows)
      this.total.set(r.totalApproved)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstWorksheetTrackList_ItemCheck: no invoice date / no group company. */
  protected check(keys: Set<string>): void {
    const before = this.checked()
    for (const x of this.rows()) {
      const k = this.key(x)
      if (keys.has(k) && !before.has(k) && x['TickError']) {
        keys.delete(k)
        this.notify.notify(String(x['TickError']), 'error')
      }
    }
    this.checked.set(keys)
  }

  protected checkAll(): void {
    this.check(new Set(this.rows().map(this.key)))
  }

  /** CmdBordx_Click. */
  protected async bordx(): Promise<void> {
    if (!this.rows().length) return this.notify.notify('Please select a record before proceeding!', 'error')
    if (!this.checked().size) return this.notify.notify('Please select at least one record to be updated! ', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    this.saving.set(true)
    try {
      const ticked = this.rows().filter((x) => this.checked().has(this.key(x)))
        .map((x) => ({ invoiceNo: s(x['InvoiceNo']), invoiceDate: x['InvoiceDate'], registerDate: x['RegisterDate'], grpCoCode: s(x['GRPCoCode']) }))
      const r = await firstValueFrom(this.http.post<{ bordxNo: string; message: string }>(`${BASE}/bordx`, { groupCompany: this.f.groupCompany, ticked }))
      this.rows.set([])
      this.total.set(0)
      this.checked.set(new Set())
      this.bordxNo.set(r.bordxNo)
      this.notify.notify(r.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  protected clear(): void {
    this.f = { registerFrom: '', registerTo: '', groupCompany: '', payor: '' }
    this.rows.set([])
    this.total.set(0)
    this.checked.set(new Set())
    this.bordxNo.set('')
  }
}
