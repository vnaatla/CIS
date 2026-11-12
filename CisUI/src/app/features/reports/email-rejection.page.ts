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
import { saveBlob, saveCsv } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/email/rejection'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')

/** Invoice > Email Notification (Rejection) (frmEmailRejectedClaim). */
@Component({
  selector: 'app-email-rejection-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Email Notification (Rejection)" subtitle="Replaces frmEmailRejectedClaim">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Registration Date From"><app-date-input name="registerFrom" ariaLabel="Registration Date From" [(ngModel)]="f.registerFrom" /></app-field>
          <app-field label="Registration Date To"><app-date-input name="registerTo" ariaLabel="Registration Date To" [(ngModel)]="f.registerTo" /></app-field>
          <app-field label="Case No From"><input name="caseFrom" maxlength="20" [(ngModel)]="f.caseFrom" /></app-field>
          <app-field label="Case No To"><input name="caseTo" maxlength="20" [(ngModel)]="f.caseTo" /></app-field>
          <app-field label="Status">
            <select name="status" [(ngModel)]="f.status"><option value=""></option><option>N - New</option><option>R - Resend</option></select>
          </app-field>
          <app-field label="Group Company" [wide]="true">
            <input name="groupCompany" list="er-companies" [(ngModel)]="f.groupCompany" />
            <datalist id="er-companies">@for (c of companies(); track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [busy]="sending()" (click)="send()">Send</button>
            <button appButton (click)="print()">Print</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Rejected claims" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="onChecked($event)" />
        <p class="totals" aria-label="Totals">Records: <strong>{{ rows().length }}</strong> · Ticked: <strong>{{ checked().size }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class EmailRejectionPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly companies = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly checked = signal(new Set<string>())
  protected readonly busy = signal(false)
  protected readonly sending = signal(false)
  protected f = { registerFrom: '', registerTo: '', caseFrom: '', caseTo: '', status: '', groupCompany: '' }
  protected readonly key = (x: Row) => s(x['CASNumber'])
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'MBMName', label: 'Member Name' }, { key: 'MBMIcBcPp', label: 'IC No' }, { key: 'MBMGroupCompany', label: 'Group Company' },
    { key: 'PatientName', label: 'Patient' }, { key: 'CLMTotalActual', label: 'Incurred', align: 'right', value: (x) => formatMoney(Number(x['CLMTotalActual'] ?? 0)) },
    { key: 'CLMBordxAmt', label: 'Bordx Amount', align: 'right', value: (x) => formatMoney(Number(x['CLMBordxAmt'] ?? 0)) },
    { key: 'DateVisit', label: 'Visit Date', value: (x) => d(x['DateVisit']) }, { key: 'NPClinic', label: 'Non-Panel Clinic' }, { key: 'CASNumber', label: 'Case No' },
    { key: 'InvoiceReceiptDate', label: 'Submitted', value: (x) => d(x['InvoiceReceiptDate']) }, { key: 'CASEmailRejStatus', label: 'Notified' },
    { key: 'MBMnumber', label: 'Member No' }, { key: 'MBMEmail', label: 'E-mail' }, { key: 'RepudiationReason', label: 'Reason' }, { key: 'InvoiceNo', label: 'Invoice No' },
    { key: 'CASRepAction', label: 'Action' }, { key: 'CASRepAmt', label: 'Rejected Amount', align: 'right', value: (x) => formatMoney(Number(x['CASRepAmt'] ?? 0)) },
    { key: 'MBMPolicyNo', label: 'Policy No' }, { key: 'MBMEmployeeNo', label: 'Employee No' },
  ]

  async ngOnInit(): Promise<void> {
    try {
      this.companies.set(await firstValueFrom(this.http.get<string[]>('/api/email/payment/companies')))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.checked.set(new Set())
    try {
      this.rows.set(await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, this.f)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstEmailListing_ItemCheck: a claim already notified asks before it is ticked again. */
  protected async onChecked(keys: Set<string>): Promise<void> {
    const before = this.checked()
    const added = this.rows().filter((x) => keys.has(this.key(x)) && !before.has(this.key(x)) && s(x['CASEmailRejStatus']) === 'Y')
    for (const x of added)
      if (!(await this.notify.ask('Email has been sent to the member! Do you want to send again?'))) keys.delete(this.key(x))
    this.checked.set(new Set(keys))
  }

  protected async send(): Promise<void> {
    const ticked = this.rows().filter((x) => this.checked().has(this.key(x)))
    if (!ticked.length) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to start sending email(s) ?'))) return
    this.sending.set(true)
    try {
      const body = {
        ticked: ticked.map((x) => ({
          caseNo: s(x['CASNumber']), email: s(x['MBMEmail']), memberName: s(x['MBMName']), patientName: s(x['PatientName']), ic: s(x['MBMIcBcPp']),
          policy: s(x['MBMPolicyNo']), invoiceNo: s(x['InvoiceNo']), memberNo: s(x['MBMnumber']), submitted: d(x['InvoiceReceiptDate']), visitDate: d(x['DateVisit']),
          amount: s(x['CASRepAmt']), reason: s(x['RepudiationReason']), action: s(x['CASRepAction']), groupCompany: s(x['MBMGroupCompany']),
        })),
      }
      const r = await firstValueFrom(
        this.http.post<{ sent: number; letters: number; lettersFileName: string | null; lettersZip: string | null; message: string }>(`${BASE}/send`, body))
      if (r.lettersZip && r.lettersFileName) {
        const bytes = Uint8Array.from(atob(r.lettersZip), (c) => c.charCodeAt(0))
        saveBlob(new Blob([bytes], { type: 'application/zip' }), r.lettersFileName)
      }
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
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'EmailNotificationRejection.csv')
  }
}
