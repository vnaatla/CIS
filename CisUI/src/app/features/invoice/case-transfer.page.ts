import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { FieldComponent } from '../../shared/field.component'
import { formatDate } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/invoice/case-transfer'

interface TransferCase {
  casNumber: string
  mbmNumber: string
  covId: string
  patient: string
  visitDate: string | null
}

/** Invoice > Other > Case Transfer (frmMembershipTransfer): move a case to another member / cover. */
@Component({
  selector: 'app-case-transfer-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Case Transfer" subtitle="Replaces frmMembershipTransfer">
      <div class="screen-form">
        <fieldset aria-label="Case">
          <legend>Case</legend>
          <app-field label="Case No." [required]="true">
            <input name="casNumber" data-autofocus maxlength="20" [(ngModel)]="casNumber" (ngModelChange)="canTransfer.set(false)" (keydown.enter)="go()" />
          </app-field>
          <app-action-bar label="Case actions">
            <button appButton variant="primary" [busy]="busy()" (click)="go()">Go</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Current Member">
          <legend>Current Member</legend>
          <app-field label="Member No."><input name="currentMember" [value]="current()?.mbmNumber ?? ''" disabled /></app-field>
          <app-field label="Cov ID"><input name="currentCovId" [value]="current()?.covId ?? ''" disabled /></app-field>
          <app-field label="Patient"><input name="currentPatient" [value]="current()?.patient ?? ''" disabled /></app-field>
          <app-field label="Visit Date"><input name="visitDate" [value]="fmt(current()?.visitDate)" disabled /></app-field>
        </fieldset>

        <fieldset aria-label="Transfer Member">
          <legend>Transfer Member</legend>
          <app-field label="Member No." [required]="true">
            <input name="mbmNumber" maxlength="50" [(ngModel)]="mbmNumber" (ngModelChange)="newMember()" (keydown.enter)="search()" />
          </app-field>
          <app-field label="Cov ID" [required]="true">
            <input name="covId" maxlength="2" [(ngModel)]="covId" (ngModelChange)="newMember()" (keydown.enter)="search()" />
          </app-field>
          <app-field label="Patient"><input name="transferPatient" [value]="patient()" disabled /></app-field>
          <app-action-bar label="Transfer member actions">
            <button appButton (click)="search()">Search</button>
          </app-action-bar>
        </fieldset>

        <app-action-bar label="Case transfer actions">
          <button appButton variant="primary" [disabled]="!canTransfer()" [busy]="busy()" (click)="transfer()">Transfer</button>
        </app-action-bar>
      </div>
    </app-page-shell>
  `,
})
export class CaseTransferPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly current = signal<TransferCase | null>(null)
  protected readonly patient = signal('')
  /** cmdSave is enabled only after Go found the case and Search found the transfer member; any change disables it. */
  protected readonly canTransfer = signal(false)
  protected readonly busy = signal(false)
  protected casNumber = ''
  protected mbmNumber = ''
  protected covId = ''

  protected fmt(v: string | null | undefined): string {
    return formatDate(v)
  }

  protected newMember(): void {
    this.canTransfer.set(false)
    this.patient.set('')
  }

  /** CmdGo_Click: ClearForm (the case number is kept), then ListCases. */
  protected async go(): Promise<void> {
    this.current.set(null)
    this.mbmNumber = ''
    this.covId = ''
    this.newMember()
    if (!this.casNumber.trim()) return this.notify.notify('Please select Case Number ! ', 'error')
    this.busy.set(true)
    try {
      this.current.set(await firstValueFrom(this.http.get<TransferCase>(`${BASE}/${encodeURIComponent(this.casNumber.trim())}`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdSearch_Click: the transfer member's name; Transfer is enabled when it is found. */
  protected async search(): Promise<void> {
    this.newMember()
    try {
      const params = new HttpParams({ fromObject: { mbmNumber: this.mbmNumber, covId: this.covId } })
      this.patient.set((await firstValueFrom(this.http.get<{ name: string }>(`${BASE}/member`, { params }))).name)
      this.canTransfer.set(!!this.current())
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdSave_Click: CISCAS member number and cover id → "Record Updated!". */
  protected async transfer(): Promise<void> {
    if (!(await this.notify.ask(`Transfer case ${this.casNumber.trim()} to ${this.mbmNumber.trim()}-${this.covId.trim()}?`))) return
    this.busy.set(true)
    try {
      const body = { casNumber: this.casNumber, mbmNumber: this.mbmNumber, covId: this.covId }
      this.notify.notify((await firstValueFrom(this.http.post<{ message: string }>(BASE, body))).message, 'success')
      this.canTransfer.set(false)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }
}
