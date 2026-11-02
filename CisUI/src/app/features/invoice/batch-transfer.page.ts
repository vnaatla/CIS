import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { FieldComponent } from '../../shared/field.component'
import { formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/invoice/batch-transfer'

interface BatchInfo {
  batchNo: string
  cases: number
  invoiceNo: string
  invoiceAmt: number
  clinic: string
  warning: string | null
}

/** Invoice > Clinic > Batch Transfer (FrmInvoiceBatchNoUpd): move the cases of OP batch 1 to OP batch 2 (same invoice). */
@Component({
  selector: 'app-batch-transfer-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Invoice Batch Number Transfer" subtitle="Replaces FrmInvoiceBatchNoUpd">
      <div class="screen-form">
        @for (n of [0, 1]; track n) {
          <fieldset [attr.aria-label]="'Batch ' + (n + 1)">
            <legend>{{ n === 0 ? 'From' : 'To' }} Batch</legend>
            <app-field [label]="'Batch No. ' + (n + 1)" [required]="true">
              <input [attr.name]="'batch' + (n + 1)" maxlength="20" [attr.data-autofocus]="n === 0 ? '' : null" [ngModel]="text[n]"
                (ngModelChange)="text[n] = $event" (blur)="show(n)" (keydown.enter)="show(n)" />
            </app-field>
            <app-field label="No. of Cases"><input class="num" [attr.name]="'cases' + (n + 1)" [value]="info()[n]?.cases ?? ''" disabled /></app-field>
            <app-field label="Invoice No."><input [attr.name]="'invoiceNo' + (n + 1)" [value]="info()[n]?.invoiceNo ?? ''" disabled /></app-field>
            <app-field label="Invoice Amt"><input class="num" [attr.name]="'invoiceAmt' + (n + 1)" [value]="amt(info()[n]?.invoiceAmt)" disabled /></app-field>
            <app-field label="Clinic Name"><input [attr.name]="'clinic' + (n + 1)" [value]="info()[n]?.clinic ?? ''" disabled /></app-field>
          </fieldset>
        }
        <fieldset aria-label="Final">
          <legend>Final</legend>
          <app-field label="Final Batch No."><input name="finalBatch" [value]="info()[1]?.batchNo ?? ''" disabled /></app-field>
          <app-field label="Total Cases"><input class="num" name="finalCases" [value]="finalCases()" disabled /></app-field>
          <app-field label="Total Amt"><input class="num" name="finalAmt" [value]="finalAmt()" disabled /></app-field>
          <app-action-bar label="Transfer actions">
            <button appButton variant="primary" [busy]="busy()" (click)="save()">Transfer</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
      </div>
    </app-page-shell>
  `,
})
export class BatchTransferPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly info = signal<(BatchInfo | null)[]>([null, null])
  protected readonly busy = signal(false)
  protected text = ['', '']
  private shown = ['', '']

  /** FinalBatchCal. */
  protected readonly finalCases = computed(() => (this.info()[0] || this.info()[1] ? (this.info()[0]?.cases ?? 0) + (this.info()[1]?.cases ?? 0) : ''))
  protected readonly finalAmt = computed(() =>
    this.info()[0] || this.info()[1] ? formatMoney((this.info()[0]?.invoiceAmt ?? 0) + (this.info()[1]?.invoiceAmt ?? 0)) : '')

  protected amt(v: number | undefined): string {
    return v == null ? '' : formatMoney(v)
  }

  /** txtBatchNo_LostFocus → ShowInfo: an unknown batch is cleared ("Invalid OP Batch Number ! "). */
  protected async show(n: number): Promise<void> {
    this.text[n] = this.text[n].trim().toUpperCase()
    if (this.text[n] === this.shown[n]) return
    this.shown[n] = this.text[n]
    this.set(n, null)
    if (!this.text[n]) return
    try {
      const b = await firstValueFrom(this.http.get<BatchInfo>(`${BASE}/${encodeURIComponent(this.text[n])}`))
      this.set(n, b)
      if (b.warning) this.notify.notify(b.warning, 'info')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
      this.text[n] = ''
      this.shown[n] = ''
    }
  }

  /** CmdSave_Click. */
  protected async save(): Promise<void> {
    const [a, b] = this.info()
    if (!this.text[0].trim() || !this.text[1].trim()) return this.notify.notify('Please confirm your Batch Number ! ', 'error')
    if ((a?.invoiceNo ?? '') !== (b?.invoiceNo ?? '')) return this.notify.notify('Different Invoice No. Must be SAME Invoice No. ! ', 'error')
    if (this.text[0].trim().toUpperCase() === this.text[1].trim().toUpperCase()) return this.notify.notify('Same Batch Number ! ', 'error')
    if (!(await this.notify.ask('Are you sure?'))) return
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ message: string }>(BASE, { fromBatch: this.text[0], toBatch: this.text[1] }))
      this.notify.notify(r.message, 'success')
      this.clear()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected clear(): void {
    this.text = ['', '']
    this.shown = ['', '']
    this.info.set([null, null])
  }

  private set(n: number, b: BatchInfo | null): void {
    this.info.update((list) => list.map((x, i) => (i === n ? b : x)))
  }
}
