import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/claim-bordx/update-claims-data'
type Kind = 'Ffs' | 'Copay' | 'BordxAmt' | 'BordxDate'

interface ClaimDataError {
  row: number
  bordxNo: string
  appCode: string
  appDate: string
  amount: string
  date: string
}
interface ClaimDataResult {
  message: string
  rows: number
  errors: ClaimDataError[]
  updated: number
}

const KINDS: [Kind, string, string][] = [
  ['Ffs', 'Verify FFS', 'FFS'],
  ['Copay', 'Verify Copay', 'Copay'],
  ['BordxAmt', 'Verify Bordx Amt', 'Bordx Amt'],
  ['BordxDate', 'Verify Bordx Date', 'Bordx Date'],
]

/** Claim Bordx > Update Claims Data (frmUpdateClaimsData): verify a spreadsheet, then upload FFS, co-pay, bordereaux amounts or dates. */
@Component({
  selector: 'app-update-claims-data-page',
  imports: [PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Update Claims Data" subtitle="Replaces frmUpdateClaimsData">
      <div class="screen-form">
        <fieldset aria-label="File">
          <legend>File</legend>
          <app-field label="File" [wide]="true">
            <input type="file" name="file" accept=".xlsx,.xlsm" (change)="picked($any($event.target).files)" />
          </app-field>
          @for (k of kinds; track k[0]) {
            <app-action-bar [label]="k[2] + ' actions'">
              <button appButton [busy]="busy()" (click)="verify(k[0])">{{ k[1] }}</button>
              <button appButton variant="primary" [busy]="busy()" [disabled]="verified() !== k[0]" (click)="upload(k[0])">{{ k[2] }}</button>
            </app-action-bar>
          }
          <app-action-bar label="Print">
            <button appButton (click)="print()">Print</button>
          </app-action-bar>
        </fieldset>
        @if (result(); as r) {
          <p class="totals" aria-label="Result">{{ r.message }} · Rows: <strong>{{ r.rows }}</strong>
            @if (r.updated) { · Updated: <strong>{{ r.updated }}</strong> }</p>
        }
        <app-data-grid caption="Rows with errors" [columns]="columns" [rows]="errors()" [rowKey]="key" />
      </div>
    </app-page-shell>
  `,
})
export class UpdateClaimsDataPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly kinds = KINDS
  protected readonly busy = signal(false)
  protected readonly verified = signal<Kind | null>(null)
  protected readonly result = signal<ClaimDataResult | null>(null)
  protected readonly errors = signal<ClaimDataError[]>([])
  protected readonly key = (r: ClaimDataError) => String(r.row)
  protected readonly columns: GridColumn<ClaimDataError>[] = [
    { key: 'row', label: 'Row', align: 'right' }, { key: 'bordxNo', label: 'Bordx No' }, { key: 'appCode', label: 'Approval Code' },
    { key: 'appDate', label: 'Approval Date' }, { key: 'amount', label: 'Amount', align: 'right' }, { key: 'date', label: 'Date' },
  ]
  private file: File | null = null

  protected picked(files: FileList | null): void {
    this.file = files?.[0] ?? null
    this.verified.set(null)
    this.result.set(null)
    this.errors.set([])
  }

  /** cmdVerify*: the upload button of the same kind is enabled after a clean verification. */
  protected async verify(kind: Kind): Promise<void> {
    const r = await this.send(kind, 'verify')
    if (!r) return
    this.verified.set(r.errors.length === 0 ? kind : null)
    this.notify.notify(r.message, r.errors.length === 0 ? 'success' : 'error')
  }

  /** cmdUpload / cmdReturnDate / cmdUploadSend / cmdUpdateBordxDate. */
  protected async upload(kind: Kind): Promise<void> {
    if (!(await this.notify.ask('Do you want to Upload the records?'))) return
    const r = await this.send(kind, 'upload')
    if (!r) return
    this.verified.set(null)
    this.notify.notify(r.message, r.errors.length === 0 ? 'success' : 'error')
  }

  protected print(): void {
    if (this.errors().length === 0) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([['Row', 'Bordx No', 'Approval Code', 'Approval Date', 'Amount', 'Date'],
      ...this.errors().map((e) => [e.row, e.bordxNo, e.appCode, e.appDate, e.amount, e.date])], 'UpdateClaimsDataErrors.csv')
  }

  private async send(kind: Kind, action: 'verify' | 'upload'): Promise<ClaimDataResult | null> {
    if (!this.file) {
      this.notify.notify('Please select file to upload!', 'error')
      return null
    }
    const body = new FormData()
    body.append('file', this.file)
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<ClaimDataResult>(`${BASE}/${kind}/${action}`, body))
      this.result.set(r)
      this.errors.set(r.errors)
      return r
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
      return null
    } finally {
      this.busy.set(false)
    }
  }
}
