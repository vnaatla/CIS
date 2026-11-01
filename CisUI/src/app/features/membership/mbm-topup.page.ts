import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { FieldComponent } from '../../shared/field.component'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { MemberRow, MemberSearchModalComponent, MemberSearchMode } from './member-search.modal'

const BASE = '/api/membership/topup'

interface TopupMember {
  mbmNumber: string
  covId: string
  name: string
  icNo: string
  plan: string
  insType: string
  effDate: string | null
  expDate: string | null
  status: string
  groupCompany: string
  department: string
  balanceLimit: number | null
  limitCode: string
  limitLabel: string
  topupEnabled: boolean
}

/** Membership > MBM Topup Adj (FrmMBMTopupLmt): "Member Topup Limit and Amendment". */
@Component({
  selector: 'app-mbm-topup-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, MemberSearchModalComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Member Topup Limit" subtitle="Member Topup Limit and Amendment · replaces FrmMBMTopupLmt">
      <div class="screen-form">
        <fieldset aria-label="Member">
          <legend>Member</legend>
          <app-field label="Member No" hint="NUMBER or NUMBER-ID; 8–9 characters list the matching members">
            <input name="member" data-autofocus [(ngModel)]="memberText" (keydown.enter)="go()" (blur)="memberText = memberText.toUpperCase()" />
          </app-field>
          <app-action-bar label="Member actions">
            <button appButton variant="primary" [busy]="busy()" (click)="go()">Go</button>
            <button appButton (click)="step(false)">Prev</button>
            <button appButton (click)="step(true)">Next</button>
            <button appButton (click)="openSearch('search')">Search</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Details">
          <legend>Details</legend>
          <app-field label="MBM Name"><input name="name" [value]="m()?.name ?? ''" disabled /></app-field>
          <app-field label="IC Number"><input name="icNo" maxlength="30" [disabled]="!m()" [(ngModel)]="icNo" (blur)="icNo = icNo.trim().toUpperCase()" /></app-field>
          <app-field label="Group Comp."><input name="groupCompany" [value]="m()?.groupCompany ?? ''" disabled /></app-field>
          <app-field label="Department"><input name="department" [value]="m()?.department ?? ''" disabled /></app-field>
          <app-field label="Plan Code"><input name="plan" [value]="m()?.plan ?? ''" disabled /></app-field>
          <app-field label="Ins Type"><input name="insType" [value]="m()?.insType ?? ''" disabled /></app-field>
          <app-field label="Eff Date"><input name="effDate" [value]="fmtDate(m()?.effDate)" disabled /></app-field>
          <app-field label="Exp Date"><input name="expDate" [value]="fmtDate(m()?.expDate)" disabled /></app-field>
          <app-field label="Status"><input name="status" [value]="m()?.status ?? ''" disabled /></app-field>
          <app-field label="Limit Type"><input name="limitType" [value]="m()?.limitLabel ?? ''" disabled /></app-field>
        </fieldset>

        <fieldset aria-label="Topup">
          <legend>Topup</legend>
          <div class="radio-row" role="radiogroup" aria-label="Limit">
            <label><input type="radio" name="limit" [value]="true" [(ngModel)]="gp" (ngModelChange)="reloadBalance()" /> GP</label>
            <label><input type="radio" name="limit" [value]="false" [(ngModel)]="gp" (ngModelChange)="reloadBalance()" /> SP</label>
          </div>
          <app-field label="Balance Lmt"><input class="num" name="balance" [value]="money(m()?.balanceLimit)" disabled /></app-field>
          <app-field label="Topup Amt">
            <input class="num" name="topupAmt" inputmode="decimal" [disabled]="!m() || !m()!.topupEnabled" [(ngModel)]="topupAmt" (input)="onlyNumber()" />
          </app-field>
          <app-field label="New Bal."><input class="num" name="newBalance" [value]="newBalance()" disabled /></app-field>
        </fieldset>

        <app-action-bar label="Topup actions">
          <button appButton variant="primary" [disabled]="!m()" [busy]="busy()" (click)="save()">Save</button>
          <button appButton variant="ghost" (click)="clear()">Clear</button>
        </app-action-bar>
      </div>
    </app-page-shell>

    @if (searchMode(); as mode) {
      <app-member-search-modal [mode]="mode" [initialText]="memberText" (picked)="picked($event)" (closed)="searchMode.set(null)" />
    }
  `,
})
export class MbmTopupPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly m = signal<TopupMember | null>(null)
  protected readonly busy = signal(false)
  protected readonly searchMode = signal<MemberSearchMode | null>(null)
  protected memberText = ''
  protected icNo = ''
  protected topupAmt = ''
  /** Form_Load: GP. */
  protected gp = true

  protected fmtDate(v: string | null | undefined): string {
    return formatDate(v)
  }

  protected money(v: number | null | undefined): string {
    return v == null ? '' : formatMoney(v)
  }

  /** txtTopupAmt_LostFocus: balance + topup. */
  protected newBalance(): string {
    if (!this.m()) return ''
    const add = Number(this.topupAmt)
    return formatMoney((this.m()!.balanceLimit ?? 0) + (Number.isFinite(add) ? add : 0))
  }

  /** KeyNumber: digits and one decimal point. */
  protected onlyNumber(): void {
    const cleaned = this.topupAmt.replace(/[^0-9.-]/g, '').replace(/(\..*)\./g, '$1')
    if (cleaned !== this.topupAmt) this.topupAmt = cleaned
  }

  /** cmdSearch_Click: frmSearchMBM. */
  protected openSearch(mode: MemberSearchMode): void {
    this.clear()
    this.searchMode.set(mode)
  }

  /** frmSearchMBM / frmSearchMBM05 double-click: "NUMBER-ID", then Go. */
  protected async picked(r: MemberRow): Promise<void> {
    this.searchMode.set(null)
    this.memberText = `${r.mbmNumber}-${r.coverId ?? '00'}`
    await this.find()
  }

  /** CmdGo_Click: 8 or 9 characters open frmSearchMBM05 (one match is taken directly); otherwise FindMBM with GP. */
  protected async go(): Promise<void> {
    this.memberText = this.memberText.trim().toUpperCase()
    if (!this.memberText) return this.notify.notify('Please Enter Membership No', 'error')
    if (this.memberText.length === 8 || this.memberText.length === 9) {
      try {
        const params = new HttpParams({ fromObject: { memberNo: this.memberText } })
        const r = await firstValueFrom(this.http.get<{ rows: MemberRow[] }>('/api/membership/search/prefix', { params }))
        if (r.rows.length === 1) return this.picked(r.rows[0])
        this.reset()
        return this.searchMode.set('prefix')
      } catch (e) {
        this.reset()
        return this.notify.notify(errorMessage(e), 'info')
      }
    }
    this.gp = true
    await this.find()
  }

  /** CmdNext_Click / CmdPreview_Click. */
  protected async step(next: boolean): Promise<void> {
    try {
      const params = new HttpParams({ fromObject: { member: this.memberText, next } })
      this.memberText = (await firstValueFrom(this.http.get<{ member: string }>(`${BASE}/step`, { params }))).member
      this.gp = true
      await this.find()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** optGPlmt_Click / optSPlmt_Click: the balance of the chosen limit. */
  protected async reloadBalance(): Promise<void> {
    if (this.m()) await this.find(false)
  }

  /** CmdSave_Click. */
  protected async save(): Promise<void> {
    const m = this.m()
    if (!m) return
    this.busy.set(true)
    try {
      const body = { member: this.memberText, gp: this.gp, icNo: this.icNo, prevIcNo: m.icNo, balanceLimit: String(m.balanceLimit ?? ''), topupAmt: this.topupAmt, limitCode: m.limitCode }
      this.notify.notify((await firstValueFrom(this.http.post<{ message: string }>(BASE, body))).message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** CmdClear_Click: every box. */
  protected clear(): void {
    this.memberText = ''
    this.reset()
  }

  private async find(resetEntries = true): Promise<void> {
    this.busy.set(true)
    if (resetEntries) this.reset()
    try {
      const params = new HttpParams({ fromObject: { member: this.memberText, gp: this.gp } })
      const m = await firstValueFrom(this.http.get<TopupMember>(BASE, { params }))
      this.m.set(m)
      if (resetEntries) this.icNo = m.icNo
    } catch (e) {
      this.m.set(null)
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  private reset(): void {
    this.m.set(null)
    this.icNo = ''
    this.topupAmt = ''
  }
}
