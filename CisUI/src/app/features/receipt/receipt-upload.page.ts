import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
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
import { formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/receipt/upload'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => (r[key] == null ? '' : formatMoney(Number(r[key]))) })

/** Receipt > Upload Receipt (frmReceiptUpload) and Upload Receipt (Multiple Receipts) (frmReceiptUpMultiple). */
@Component({
  selector: 'app-receipt-upload-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="multiple ? 'Upload Receipt (Multiple Receipts)' : 'Upload Receipt'" [subtitle]="'Replaces ' + (multiple ? 'frmReceiptUpMultiple' : 'frmReceiptUpload')">
      <div class="screen-form">
        <fieldset aria-label="Upload">
          <legend>Upload</legend>
          <app-field label="File" [wide]="true"><input type="file" name="file" accept=".xlsx" (change)="pick($event)" /></app-field>
          <app-field label="DB"><input name="db" value="UHC" disabled /></app-field>
          <app-action-bar label="Upload actions">
            <button appButton variant="primary" [busy]="busy()" (click)="upload()">Upload</button>
            <button appButton (click)="print()">Print</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Receipt">
          <legend>Receipt</legend>
          @if (!multiple) {
            <app-field label="Receipt No" [required]="true"><input name="receiptNo" maxlength="30" [(ngModel)]="r.receiptNo" /></app-field>
            <app-field label="Medix Receipt No" [required]="true"><input name="medixReceiptNo" maxlength="30" [(ngModel)]="r.medixReceiptNo" /></app-field>
          }
          <app-field label="Receipt Date"><app-date-input name="receiptDate" ariaLabel="Receipt Date" [(ngModel)]="r.receiptDate" (ngModelChange)="period()" /></app-field>
          <app-field label="Receipt Amount"><input class="num" name="receiptAmt" maxlength="20" [(ngModel)]="r.receiptAmt" /></app-field>
          <app-field label="Cheque No"><input name="chequeNo" maxlength="30" [(ngModel)]="r.chequeNo" /></app-field>
          <app-field label="A/C Period Month"><input name="month" maxlength="2" [(ngModel)]="r.month" /></app-field>
          <app-field label="A/C Period Year"><input name="year" maxlength="4" [(ngModel)]="r.year" /></app-field>
          <app-action-bar label="Receipt actions">
            <button appButton variant="primary" [busy]="saving()" (click)="save()">Update</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Uploaded receipts" [columns]="columns" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class ReceiptUploadPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  protected readonly multiple: boolean = inject(ActivatedRoute).snapshot.data['multiple'] === true

  protected readonly rows = signal<Row[]>([])
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected r = { receiptNo: '', medixReceiptNo: '', receiptDate: '', receiptAmt: '', chequeNo: '', month: '', year: '' }
  private file: File | null = null
  protected readonly key = (x: Row) => `${this.rows().indexOf(x)}|${s(x['claimNo'])}`
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'claimNo', label: 'Claim No (file)' }, { key: 'bordxNo', label: 'Bordx No (file)' }, money('amount', 'Amount (file)'), { key: 'claimNoCis', label: 'Claim No (CIS)' },
    { key: 'bordxNoCis', label: 'Bordx No (CIS)' }, money('amountCis', 'Amount (CIS)'), { key: 'caseStatus', label: 'Claim' }, { key: 'bordxStatus', label: 'Bordx' },
    money('difference', 'Difference'), { key: 'clientPvNo', label: 'Client PV No' },
    ...(this.multiple ? [{ key: 'receiptNo', label: 'Receipt No' }, { key: 'medixReceiptNo', label: 'Medix Receipt No' }] : []),
  ]

  protected pick(e: Event): void {
    this.file = (e.target as HTMLInputElement).files?.[0] ?? null
  }

  /** cmpUpload_Click. */
  protected async upload(): Promise<void> {
    if (!this.file) return this.notify.notify('Please select file to upload!', 'error')
    this.busy.set(true)
    try {
      const form = new FormData()
      form.append('file', this.file)
      this.rows.set(await firstValueFrom(this.http.post<Row[]>(BASE, form, { params: { multiple: this.multiple } })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** txtReceiptDate_LostFocus: the period from the receipt date. */
  protected period(): void {
    const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(this.r.receiptDate)
    this.r.month = m ? String(Number(m[2])) : ''
    this.r.year = m ? m[3] : ''
  }

  /** CmdUpdate_Click. */
  protected async save(): Promise<void> {
    if (!this.rows().length) return this.notify.notify('Please Upload file..! ', 'error')
    if (!this.multiple && !this.r.receiptNo.trim()) return this.notify.notify('Please Enter Receipt Number! ', 'error')
    if (!this.multiple && !this.r.medixReceiptNo.trim()) return this.notify.notify('Please Enter Medix Receipt Number! ', 'error')
    if (!(await this.notify.ask('Do you want to update the record ?'))) return
    this.saving.set(true)
    try {
      const res = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/save`, { ...this.r, multiple: this.multiple, rows: this.rows() }))
      this.notify.notify(res.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'ReceiptUpload.csv')
  }
}
