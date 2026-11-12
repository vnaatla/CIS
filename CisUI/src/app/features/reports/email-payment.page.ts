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
import { saveCsv } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/email/payment'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())

/** Invoice > Email Notification (Payment) (frmEmailNotification). */
@Component({
  selector: 'app-email-payment-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Email Notification (Payment)" subtitle="Replaces frmEmailNotification">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Entity">
            <select name="entity" [(ngModel)]="f.entity"><option [ngValue]="0">Health Connect</option><option [ngValue]="1">Sihat Johor</option></select>
          </app-field>
          <app-field label="PV Date From"><app-date-input name="pvDateFrom" ariaLabel="PV Date From" [(ngModel)]="f.pvDateFrom" /></app-field>
          <app-field label="PV Date To"><app-date-input name="pvDateTo" ariaLabel="PV Date To" [(ngModel)]="f.pvDateTo" /></app-field>
          <app-field label="PV No From"><input name="pvFrom" maxlength="30" [(ngModel)]="f.pvFrom" /></app-field>
          <app-field label="PV No To"><input name="pvTo" maxlength="30" [(ngModel)]="f.pvTo" /></app-field>
          <app-field label="Status">
            <select name="status" [(ngModel)]="f.status"><option value=""></option><option>N - New</option><option>R - Resend</option></select>
          </app-field>
          <app-field label="Group Company" [wide]="true">
            <input name="groupCompany" list="ep-companies" [(ngModel)]="f.groupCompany" />
            <datalist id="ep-companies">@for (c of companies(); track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
          </app-action-bar>
        </fieldset>
        <fieldset aria-label="Send">
          <legend>Send</legend>
          <app-field label="Date Banked-in"><app-date-input name="sendDate" ariaLabel="Send Date" [(ngModel)]="sendDate" /></app-field>
          <app-field label="Extra BCC" [wide]="true"><input name="extraBcc" maxlength="500" placeholder="a@b.com; c@d.com" [(ngModel)]="extraBcc" /></app-field>
          <label class="check-field"><input type="checkbox" name="checkAll" [checked]="allChecked()" (change)="checkAll($any($event.target).checked)" /> Check all</label>
          <app-action-bar label="Send actions">
            <button appButton variant="primary" [busy]="sending()" (click)="send()">Send</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Payments to e-mail" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="onChecked($event)" />
        <p class="totals" aria-label="Totals">Records: <strong>{{ rows().length }}</strong> · Ticked: <strong>{{ checked().size }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class EmailPaymentPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly companies = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly checked = signal(new Set<string>())
  protected readonly allChecked = signal(false)
  protected readonly busy = signal(false)
  protected readonly sending = signal(false)
  protected f = { entity: 0, pvDateFrom: '', pvDateTo: '', pvFrom: '', pvTo: '', status: '', groupCompany: '' }
  protected sendDate = ''
  protected extraBcc = ''
  protected readonly key = (x: Row) => s(x['PVNo'])
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'MBMName', label: 'Member Name' }, { key: 'MBMIcBcPp', label: 'IC No' }, { key: 'MBMGroupCompany', label: 'Group Company' },
    { key: 'PatientName', label: 'Patient' }, { key: 'CLMTotalActual', label: 'Incurred', align: 'right', value: (x) => formatMoney(Number(x['CLMTotalActual'] ?? 0)) },
    { key: 'CLMBordxAmt', label: 'Paid', align: 'right', value: (x) => formatMoney(Number(x['CLMBordxAmt'] ?? 0)) }, { key: 'PVNo', label: 'PV No' },
    { key: 'PVDate', label: 'PV Date', value: (x) => (x['PVDate'] ? formatDate(String(x['PVDate'])) : '') }, { key: 'MBMBankACNo', label: 'Bank Account' },
    { key: 'InvoiceReceiptDate', label: 'Submitted', value: (x) => (x['InvoiceReceiptDate'] ? formatDate(String(x['InvoiceReceiptDate'])) : '') },
    { key: 'PVEmailStatus', label: 'E-mailed' }, { key: 'MBMnumber', label: 'Member No' }, { key: 'MBMEmail', label: 'E-mail' },
  ]

  async ngOnInit(): Promise<void> {
    try {
      this.companies.set(await firstValueFrom(this.http.get<string[]>(`${BASE}/companies`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.checked.set(new Set())
    this.allChecked.set(false)
    try {
      this.rows.set(await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, this.f)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstEmailListing_ItemCheck: a member without an e-mail cannot be ticked. */
  protected onChecked(keys: Set<string>): void {
    const noEmail = this.rows().filter((x) => keys.has(this.key(x)) && !s(x['MBMEmail']))
    if (noEmail.length) {
      this.notify.notify("Please check member's email!", 'error')
      noEmail.forEach((x) => keys.delete(this.key(x)))
    }
    this.checked.set(new Set(keys))
  }

  /** checkall_Click: ticks the rows with an e-mail. */
  protected checkAll(on: boolean): void {
    if (!this.rows().length) {
      this.allChecked.set(false)
      if (on) this.notify.notify('Please search a record.', 'error')
      return
    }
    this.allChecked.set(on)
    this.checked.set(new Set(on ? this.rows().filter((x) => s(x['MBMEmail'])).map(this.key) : []))
  }

  protected async send(): Promise<void> {
    if (!this.sendDate) return this.notify.notify('Please enter send date', 'error')
    const ticked = this.rows().filter((x) => this.checked().has(this.key(x)))
    if (!ticked.length) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to start sending email(s) ?'))) return
    this.sending.set(true)
    try {
      const body = {
        sendDate: this.sendDate, extraBcc: this.extraBcc,
        ticked: ticked.map((x) => ({
          pvNo: s(x['PVNo']), email: s(x['MBMEmail']), name: s(x['MBMName']), memberNo: s(x['MBMnumber']), bankAccount: s(x['MBMBankACNo']),
          submitted: x['InvoiceReceiptDate'] ? formatDate(String(x['InvoiceReceiptDate'])) : '',
        })),
      }
      const r = await firstValueFrom(this.http.post<{ sent: number; message: string }>(`${BASE}/send`, body))
      this.notify.notify(r.message, 'success')
      await this.search()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.sending.set(false)
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'EmailNotificationPayment.csv')
  }
}
