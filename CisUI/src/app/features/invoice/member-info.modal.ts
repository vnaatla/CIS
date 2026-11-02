import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, input, output, signal } from '@angular/core'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { FieldComponent } from '../../shared/field.component'
import { formatDate } from '../../shared/lib/format'
import { ModalComponent } from '../../shared/modal.component'

/** MemberSuppInfo's output (api/invoice/member-info). */
export interface MemberSupp {
  mbmNumber: string
  coverId: string
  patientName: string
  patientIc: string
  patientDob: string | null
  policyNo: string
  effDate: string | null
  expDate: string | null
  dataStatus: string
  policyStatus: string
  cardType: string
  visitLimit: string
  balanceLimit: string
  dentalStatus: string
  planCode: string
  insType: string
  dentalLimit: string
  specLimit: string
  mepLimit: string
  groupCompany: string
  [key: string]: unknown
}

/** FrmINVMBMinfo: the member details of the invoice row's member / cover. */
@Component({
  selector: 'app-member-info-modal',
  imports: [ModalComponent, FieldComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-modal title="Member Info" [wide]="true" (closed)="closed.emit()">
      @if (info(); as m) {
        <div class="screen-form">
          <fieldset aria-label="Member">
            <legend>{{ m.mbmNumber }} - {{ m.coverId }}</legend>
            <app-field label="Name"><input name="miName" [value]="m.patientName" disabled /></app-field>
            <app-field label="Member Status"><input name="miStatus" [value]="m.dataStatus" disabled /></app-field>
            <app-field label="Policy No"><input name="miPolicy" [value]="m.policyNo" disabled /></app-field>
            <app-field label="Policy Status"><input name="miPolStatus" [value]="m.policyStatus" disabled /></app-field>
            <app-field label="Eff Date"><input name="miEff" [value]="fmt(m.effDate)" disabled /></app-field>
            <app-field label="Exp Date"><input name="miExp" [value]="fmt(m.expDate)" disabled /></app-field>
            <app-field label="IC"><input name="miIc" [value]="m.patientIc" disabled /></app-field>
            <app-field label="DOB"><input name="miDob" [value]="fmt(m.patientDob)" disabled /></app-field>
            <app-field label="Card Type"><input name="miCard" [value]="m.cardType" disabled /></app-field>
            <app-field label="Group Company"><input name="miGroup" [value]="m.groupCompany" disabled /></app-field>
            <app-field label="Plan Code"><input name="miPlan" [value]="m.planCode" disabled /></app-field>
            <app-field label="Ins Type"><input name="miIns" [value]="m.insType" disabled /></app-field>
          </fieldset>
          <fieldset aria-label="Limits">
            <legend>Limits</legend>
            <app-field label="Visit Limit"><input class="num" name="miVisit" [value]="m.visitLimit" disabled /></app-field>
            <app-field label="Balance Limit"><input class="num" name="miBalance" [value]="m.balanceLimit" disabled /></app-field>
            <app-field label="Dental Status"><input name="miDentStatus" [value]="m.dentalStatus" disabled /></app-field>
            <app-field label="Dental Limit"><input class="num" name="miDent" [value]="m.dentalLimit" disabled /></app-field>
            <app-field label="Specialist Limit"><input class="num" name="miSpec" [value]="m.specLimit" disabled /></app-field>
            <app-field label="MEP Limit"><input class="num" name="miMep" [value]="m.mepLimit" disabled /></app-field>
          </fieldset>
        </div>
      }
    </app-modal>
  `,
})
export class MemberInfoModalComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  readonly mbmNumber = input.required<string>()
  readonly coverId = input('00')
  readonly closed = output<void>()
  protected readonly info = signal<MemberSupp | null>(null)
  protected readonly fmt = formatDate

  async ngOnInit(): Promise<void> {
    try {
      const url = `/api/invoice/member-info/${encodeURIComponent(this.mbmNumber())}/${encodeURIComponent(this.coverId() || '00')}`
      this.info.set(await firstValueFrom(this.http.get<MemberSupp>(url)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
      this.closed.emit()
    }
  }
}
