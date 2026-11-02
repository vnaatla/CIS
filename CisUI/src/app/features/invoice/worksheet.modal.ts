import { ChangeDetectionStrategy, Component, input, output } from '@angular/core'
import { formatMoney } from '../../shared/lib/format'
import { ModalComponent } from '../../shared/modal.component'

/** A benefit row of the worksheet (api WorksheetRow); null = "-". */
export interface WorksheetRow {
  code: string
  benefit: string
  actual: number | null
  contractedRate: number | null
  payableToClinic: number | null
  notPayableToClinic: number | null
  benefitLimit: number | null
  approved: number | null
  excess: number | null
}

/** WorksheetResult of the API. */
export interface WorksheetResult {
  rows: WorksheetRow[]
  totalActual: number
  totalPayableToClinic: number
  totalNotPayableToClinic: number
  totalApproved: number
  totalExcess: number
  coPay: number
  coPayCollected: number
  visitLimit: string
  balanceLimit: string
  mcoProcFee: number
  finalPayableToClinic: number
  finalApproved: number | null
  finalExcess: number
  finalBalanceLimit: number
}

const m = (v: number | null | undefined) => (v == null || v === 0 ? '-' : formatMoney(v))

/** frmWorksheet: the benefit grid of an invoice line (read-only; the invoice screen saves it). */
@Component({
  selector: 'app-worksheet-modal',
  imports: [ModalComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-modal title="Worksheet" [wide]="true" (closed)="closed.emit()">
      @let w = worksheet();
      <div class="table-wrap">
        <table class="grid" aria-label="Worksheet">
          <thead>
            <tr>
              <th>Benfits</th><th class="num">Actual</th><th class="num">Contracted Rate</th><th class="num">Payable to CLN</th>
              <th class="num">Not Payable to CLN</th><th class="num">Benefit Lmt</th><th class="num">Approved Benefit</th><th class="num">Excess</th>
            </tr>
          </thead>
          <tbody>
            @for (r of w.rows; track r.code) {
              <tr>
                <td>{{ r.benefit }}</td><td class="num">{{ m(r.actual) }}</td><td class="num">{{ m(r.contractedRate) }}</td>
                <td class="num">{{ m(r.payableToClinic) }}</td><td class="num">{{ m(r.notPayableToClinic) }}</td><td class="num">{{ m(r.benefitLimit) }}</td>
                <td class="num">{{ m(r.approved) }}</td><td class="num">{{ m(r.excess) }}</td>
              </tr>
            }
            <tr class="total">
              <td>Total</td><td class="num">{{ m(w.totalActual) }}</td><td></td><td class="num">{{ m(w.totalPayableToClinic) }}</td>
              <td class="num">{{ m(w.totalNotPayableToClinic) }}</td><td></td><td class="num">{{ m(w.totalApproved) }}</td><td class="num">{{ m(w.totalExcess) }}</td>
            </tr>
            <tr><td>Less Co-Payment</td><td class="num">{{ m(w.coPay) }}</td><td></td><td class="num">{{ m(-w.coPayCollected) }}</td><td></td><td></td>
              <td class="num">{{ m(-w.coPayCollected) }}</td><td></td></tr>
            <tr><td>Visit Limit</td><td class="num">{{ w.visitLimit }}</td><td colspan="6"></td></tr>
            <tr><td>Balance Limit</td><td class="num">{{ w.balanceLimit }}</td><td colspan="6"></td></tr>
            <tr><td>MCO Proc. Fee</td><td></td><td></td><td class="num">{{ m(-w.mcoProcFee) }}</td><td colspan="4"></td></tr>
            <tr class="total">
              <td>Final Approved</td><td class="num">{{ m(w.totalActual) }}</td><td></td><td class="num" data-final-cln>{{ m(w.finalPayableToClinic) }}</td><td></td><td></td>
              <td class="num" data-final-approved>{{ m(w.finalApproved) }}</td><td class="num">{{ m(w.finalExcess) }}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </app-modal>
  `,
})
export class WorksheetModalComponent {
  readonly worksheet = input.required<WorksheetResult>()
  readonly closed = output<void>()
  protected readonly m = m
}
