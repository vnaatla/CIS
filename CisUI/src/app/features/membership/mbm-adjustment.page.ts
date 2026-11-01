import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { parseDmy } from '../../shared/lib/masks'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { MemberRow, MemberSearchModalComponent, MemberSearchMode } from './member-search.modal'

const BASE = '/api/membership/adjustment'

interface Cover {
  coverId: string
  name: string
  ic: string
  relationship: string
  status: string
  exclusion: string
  remarks: string
  dob: string | null
  plan: string
  cardType: string
  premium: number
  premReduction: number | null
  premEffective: number | null
  mco: number
  mcoRefund: number | null
  mcoEffective: number | null
  annualLimit: string
  accumAmt: number
  balanceLimit: string
  visitLimit: string
  mepLimit: number
  dentalLimit: number
  dentalExtLimit: number
  specLimit: number
  firstSpecLimit: number
  ageCode: string
  adultChild: string
  cancel?: boolean
}
interface Member {
  mbmNumber: string
  status: string
  expiryStatus: string
  dataStatus: string
  payCode: string
  name: string
  policyNo: string
  ic: string
  insType: string
  plan: string
  dob: string | null
  adultChild: string
  groupCompany: string
  department: string
  limitStatus: string
  limitStatusText: string
  effDate: string | null
  expDate: string | null
  ageCode: string
  cardType: string
  exclusion: string
  remarks: string
  cisRemarks: string
  premium: number
  mcoFee: number
  accumAmt: number
  email: string
  mobile: string
  bankName: string
  bankAccount: string
  annualLimit: string
  visitLimit: string
  balanceLimit: string
  mepLimit: number
  dentalLimit: number
  specLimit: number
  dentalExtLimit: number
  firstSpecLimit: number
  covers: Cover[]
  canEdit: boolean
  notice: string | null
}

const ENDORSE: [string, string][] = [
  ['A', 'A - Amendment'], ['C', 'C - Cancellation'], ['K', 'K - Cancel Supplementary'], ['X', 'X - Suspend Membership'], ['Y', 'Y - Re-Activate Membership'], ['D', 'D - Delete Policy'],
]
const money = (v: number | null | undefined) => (v == null ? '' : formatMoney(v))

/** Membership > MBM Adjustment (FrmMBMCLNAdj): view a member and its covers, then endorse (amend, cancel, cancel supplementary, suspend, re-activate, delete). */
@Component({
  selector: 'app-mbm-adjustment-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent, TabsComponent, TabDirective,
    MemberSearchModalComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="MBM Adjustment" subtitle="Replaces FrmMBMCLNAdj">
      <div class="screen-form">
        <fieldset aria-label="Member">
          <legend>Member</legend>
          <app-field label="Member No">
            <input name="member" data-autofocus [(ngModel)]="memberText" (keydown.enter)="show()" (blur)="memberText = memberText.toUpperCase()" />
          </app-field>
          <app-action-bar label="Member actions">
            <button appButton variant="primary" [busy]="busy()" (click)="show()">Show</button>
            <button appButton (click)="step(false)">Prev</button>
            <button appButton (click)="step(true)">Next</button>
            <button appButton (click)="searchMode.set('search')">Search</button>
            <button appButton [disabled]="!m()?.canEdit" (click)="edit()">Edit</button>
          </app-action-bar>
        </fieldset>

        @if (m(); as mm) {
          <fieldset aria-label="Endorsement">
            <legend>Endorsement</legend>
            <app-field label="Endorse Status">
              <select name="endtType" [disabled]="!editing()" [(ngModel)]="endtType" (ngModelChange)="endorseChanged()">
                @for (e of endorse; track e[0]) {
                  <option [value]="e[0]">{{ e[1] }}</option>
                }
              </select>
            </app-field>
            <app-field label="Endt. Effective Date" [required]="true">
              <app-date-input name="endtEff" ariaLabel="Endt. Effective Date" [disabled]="!editing()" [(ngModel)]="endtEff" (focusout)="checkDate(endtEff)" />
            </app-field>
            <app-field label="Endt. Bordx Date" [required]="true">
              <app-date-input name="endtBordx" ariaLabel="Endt. Bordx Date" [disabled]="!editing()" [(ngModel)]="endtBordx" (focusout)="checkDate(endtBordx)" />
            </app-field>
            <app-action-bar label="Save">
              <button appButton variant="primary" [disabled]="!editing() || !endtEff || !endtBordx" [busy]="busy()" (click)="save()">Save</button>
            </app-action-bar>
          </fieldset>

          <app-tabs label="MBM Adjustment" [(active)]="tab" (activeChange)="tabChanged($event)">
            <ng-template appTab="Principal">
              <fieldset aria-label="Principal">
                <legend>Principal</legend>
                <app-field label="Name"><input name="name" [value]="mm.name" disabled /></app-field>
                <app-field label="IC No"><input name="ic" [value]="mm.ic" disabled /></app-field>
                <app-field label="Policy No"><input name="policyNo" [disabled]="!amend()" [(ngModel)]="p.policyNo" (blur)="p.policyNo = p.policyNo.toUpperCase()" /></app-field>
                <app-field label="Payor"><input name="payCode" [value]="mm.payCode" disabled /></app-field>
                <app-field label="Insured Type"><input name="insType" [value]="mm.insType" disabled /></app-field>
                <app-field label="Plan"><input name="plan" [disabled]="!amend()" [(ngModel)]="p.plan" /></app-field>
                <app-field label="Group Company"><input name="groupCompany" [value]="mm.groupCompany" disabled /></app-field>
                <app-field label="Department"><input name="department" [value]="mm.department" disabled /></app-field>
                <app-field label="Status">
                  <select name="status" disabled [ngModel]="p.status">
                    <option value="A">A - Active</option><option value="C">C - Cancel</option><option value="D">D - Deleted</option><option value="S">S - Suspended</option>
                  </select>
                </app-field>
                <app-field label="Exp. Status"><input name="expStatus" [value]="mm.expiryStatus === 'N' ? 'Y - Yes' : 'N - No'" disabled /></app-field>
                <app-field label="Data Status">
                  <select name="dataStatus" [disabled]="!amend()" [(ngModel)]="p.dataStatus"><option value="A">A - Active</option><option value="D">D - Deleted</option></select>
                </app-field>
                <app-field label="Eff. Date">
                  <app-date-input name="effDate" ariaLabel="Eff. Date" [disabled]="!amend()" [(ngModel)]="p.effDate" (focusout)="periodChanged()" />
                </app-field>
                <app-field label="Exp. Date">
                  <app-date-input name="expDate" ariaLabel="Exp. Date" [disabled]="!amend()" [(ngModel)]="p.expDate" (focusout)="periodChanged()" />
                </app-field>
                <app-field label="DOB / Age Band"><input name="dob" [value]="fmt(mm.dob) + ' / ' + p.ageCode + ' (' + p.adultChild + ')'" disabled /></app-field>
                <app-field label="Card Type">
                  <select name="cardType" [disabled]="!amend()" [(ngModel)]="p.cardType">
                    <option value=""></option>
                    @for (c of cardTypes(); track c) {
                      <option [value]="c">{{ c }}</option>
                    }
                  </select>
                </app-field>
                <app-field label="Limit Status"><input name="limitStatus" [value]="mm.limitStatusText" disabled /></app-field>
                <app-field label="E-mail"><input name="email" [disabled]="!amend()" [(ngModel)]="p.email" /></app-field>
                <app-field label="Mobile"><input name="mobile" [disabled]="!amend()" [(ngModel)]="p.mobile" /></app-field>
                <app-field label="Bank Name"><input name="bankName" [disabled]="!amend()" [(ngModel)]="p.bankName" /></app-field>
                <app-field label="Bank Acc. No"><input name="bankAccount" [disabled]="!amend()" [(ngModel)]="p.bankAccount" /></app-field>
                <app-field label="Exclusion" [wide]="true"><textarea name="exclusion" rows="2" [disabled]="!amend()" [(ngModel)]="p.exclusion"></textarea></app-field>
                <app-field label="Remarks" [wide]="true"><textarea name="remarks" rows="2" [value]="mm.remarks" disabled></textarea></app-field>
                <app-field label="CIS Remarks" [wide]="true"><textarea name="cisRemarks" rows="2" [disabled]="!amend()" [(ngModel)]="p.cisRemarks"></textarea></app-field>
              </fieldset>
              <fieldset aria-label="Fees and limits">
                <legend>Fees and Limits</legend>
                @for (f of feeFields(mm); track f[0]) {
                  <app-field [label]="f[0]"><input class="num" [attr.name]="f[0]" [value]="f[1]" disabled /></app-field>
                }
              </fieldset>
            </ng-template>

            <ng-template appTab="Supplementary">
              @if (mm.insType !== 'F' && mm.insType !== 'H') {
                <p class="muted">This Plan does not have supplementary</p>
              } @else {
                <app-data-grid caption="Covered persons" [columns]="coverColumns" [rows]="covers()" [rowKey]="coverKey" [selectable]="true"
                  [selectedKey]="cover()?.coverId ?? null" (rowSelected)="pickCover($event)" [pageSize]="0" />
                @if (cover(); as c) {
                  <fieldset aria-label="Covered person">
                    <legend>Covered Person {{ c.coverId }}</legend>
                    <app-field label="Name"><input name="covName" [value]="c.name" disabled /></app-field>
                    <app-field label="IC"><input name="covIc" [value]="c.ic" disabled /></app-field>
                    <app-field label="Relationship"><input name="covRel" [value]="c.relationship" disabled /></app-field>
                    <app-field label="Plan"><input name="covPlan" [disabled]="!supp()" [(ngModel)]="ce.plan" /></app-field>
                    <app-field label="Card Type">
                      <select name="covCard" [disabled]="!supp()" [(ngModel)]="ce.cardType">
                        <option value=""></option>
                        @for (t of cardTypes(); track t) {
                          <option [value]="t">{{ t }}</option>
                        }
                      </select>
                    </app-field>
                    <app-field label="Exclusion" [wide]="true"><textarea name="covExclusion" rows="2" [disabled]="!supp()" [(ngModel)]="ce.exclusion"></textarea></app-field>
                    <app-field label="Remarks" [wide]="true"><textarea name="covRemarks" rows="2" [disabled]="!supp()" [(ngModel)]="ce.remarks"></textarea></app-field>
                    <app-action-bar label="Cover actions">
                      <button appButton [disabled]="endtType !== 'A' || !editing()" (click)="updateCover()">Update</button>
                      <button appButton variant="danger" [disabled]="endtType !== 'K' || !editing()" (click)="cancelCover()">Cancel Supp</button>
                    </app-action-bar>
                  </fieldset>
                }
              }
            </ng-template>

            <ng-template appTab="Adjustment History">
              <app-data-grid caption="Adjustment history" [columns]="amendColumns" [rows]="amends()" [rowKey]="indexKey" />
            </ng-template>
            <ng-template appTab="MBM History">
              <app-data-grid caption="MBM history" [columns]="historyColumns" [rows]="history()" [rowKey]="indexKey" />
            </ng-template>
          </app-tabs>
        }
      </div>
    </app-page-shell>

    @if (searchMode(); as mode) {
      <app-member-search-modal [mode]="mode" [initialText]="memberText" (picked)="picked($event)" (closed)="searchMode.set(null)" />
    }
  `,
})
export class MbmAdjustmentPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly endorse = ENDORSE
  protected readonly m = signal<Member | null>(null)
  protected readonly covers = signal<Cover[]>([])
  protected readonly cover = signal<Cover | null>(null)
  protected readonly amends = signal<Record<string, unknown>[]>([])
  protected readonly history = signal<Record<string, unknown>[]>([])
  protected readonly cardTypes = signal<string[]>([])
  protected readonly editing = signal(false)
  protected readonly busy = signal(false)
  protected readonly searchMode = signal<MemberSearchMode | null>(null)
  protected readonly coverKey = (c: Cover) => c.coverId
  protected readonly indexKey = (r: Record<string, unknown>) => String(r['_i'])
  protected memberText = ''
  protected tab = 0
  protected endtType = 'A'
  protected endtEff = ''
  protected endtBordx = ''
  protected p = this.emptyPrincipal()
  protected ce = { plan: '', cardType: '', exclusion: '', remarks: '' }

  protected readonly coverColumns: GridColumn<Cover>[] = [
    { key: 'name', label: 'Name' }, { key: 'coverId', label: 'ID' }, { key: 'ic', label: 'IC' }, { key: 'relationship', label: 'Rel' },
    { key: 'status', label: 'Status', value: (c) => (c.cancel ? 'C (to cancel)' : c.status) }, { key: 'exclusion', label: 'Exclusion' },
    { key: 'remarks', label: 'Remarks' }, { key: 'dob', label: 'DOB', value: (c) => formatDate(c.dob) }, { key: 'plan', label: 'Plan' },
    { key: 'cardType', label: 'Card Type' }, { key: 'premium', label: 'Premium', align: 'right', value: (c) => money(c.premium) },
    { key: 'premReduction', label: 'Prem. Red.', align: 'right', value: (c) => money(c.premReduction) },
    { key: 'premEffective', label: 'Prem. Eff.', align: 'right', value: (c) => money(c.premEffective) },
    { key: 'mco', label: 'MCO', align: 'right', value: (c) => money(c.mco) }, { key: 'mcoRefund', label: 'MCO Ref.', align: 'right', value: (c) => money(c.mcoRefund) },
    { key: 'mcoEffective', label: 'MCO Eff.', align: 'right', value: (c) => money(c.mcoEffective) }, { key: 'annualLimit', label: 'Ann. Limit', align: 'right' },
    { key: 'accumAmt', label: 'Accum.', align: 'right', value: (c) => money(c.accumAmt) }, { key: 'balanceLimit', label: 'Bal. Limit', align: 'right' },
    { key: 'visitLimit', label: 'Visit Limit', align: 'right' }, { key: 'ageCode', label: 'Age' }, { key: 'adultChild', label: 'A/C' },
  ]
  protected readonly amendColumns: GridColumn<Record<string, unknown>>[] = [
    { key: 'BordxDate', label: 'Endt. Bordx Date', value: (r) => formatDate(r['BordxDate']) }, { key: 'AmendDate', label: 'Amend Date', value: (r) => formatDate(r['AmendDate']) },
    { key: 'EndtType', label: 'Endt. Type' }, { key: 'EndtEffDate', label: 'Endt. Eff. Date', value: (r) => formatDate(r['EndtEffDate']) }, { key: 'AmendUser', label: 'User' },
  ]
  protected readonly historyColumns: GridColumn<Record<string, unknown>>[] = [
    { key: 'MBMNumber', label: 'Member No' }, { key: 'MBMPolNumber', label: 'Policy No' }, { key: 'MBMPolEffDate', label: 'Eff. Date', value: (r) => formatDate(r['MBMPolEffDate']) },
    { key: 'MBMPolExpDate', label: 'Exp. Date', value: (r) => formatDate(r['MBMPolExpDate']) },
    { key: 'MBMCLNEndtEffDate', label: 'Endt. Eff. Date', value: (r) => (r['MBMCLNStatus'] === 'C' ? formatDate(r['MBMCLNEndtEffDate']) : '') },
    { key: 'MBMCLNDataStatus', label: 'Data St.' }, { key: 'MBMCLNStatus', label: 'Status' }, { key: 'MBMBordxDate', label: 'Bordx Date', value: (r) => formatDate(r['MBMBordxDate']) },
    { key: 'MBMCLNEndtBordxDate', label: 'Endt. Bordx', value: (r) => formatDate(r['MBMCLNEndtBordxDate']) },
    { key: 'MBMValueDate', label: 'Value Date', value: (r) => formatDate(r['MBMValueDate']) },
  ]

  async ngOnInit(): Promise<void> {
    try {
      this.cardTypes.set((await firstValueFrom(this.http.get<{ cardTypes: string[]; relationships: LookupOption[] }>(`${BASE}/setup`))).cardTypes)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected fmt(v: string | null | undefined): string {
    return formatDate(v)
  }

  protected amend(): boolean {
    return this.editing() && this.endtType === 'A'
  }

  protected supp(): boolean {
    return this.editing() && (this.endtType === 'A' || this.endtType === 'K')
  }

  protected feeFields(mm: Member): [string, string][] {
    return [
      ['Premium', money(this.p.premium)], ['MCO Fee', money(this.p.mcoFee)], ['Accum. Amt', money(mm.accumAmt)], ['Annual Limit', mm.annualLimit],
      ['Balance Limit', mm.balanceLimit], ['Visit Limit', mm.visitLimit], ['MEP Limit', money(mm.mepLimit)], ['Dental Limit', money(mm.dentalLimit)],
      ['Dental Ext. Limit', money(mm.dentalExtLimit)], ['Spec. Limit', money(mm.specLimit)], ['First Spec. Limit', money(mm.firstSpecLimit)],
    ]
  }

  protected checkDate(v: string): void {
    if (v && !parseDmy(v)) this.notify.notify('Invalid Date Format! ', 'error')
  }

  /** cmdShow_Click: 8–9 characters open frmSearchMBM05 (one match is taken), else FindMBM. */
  protected async show(): Promise<void> {
    this.memberText = this.memberText.trim().toUpperCase()
    if (!this.memberText) return this.notify.notify('Please Enter Membership No', 'error')
    if (this.memberText.length === 8 || this.memberText.length === 9) {
      try {
        const r = await firstValueFrom(this.http.get<{ rows: MemberRow[] }>('/api/membership/search/prefix', { params: new HttpParams({ fromObject: { memberNo: this.memberText } }) }))
        if (r.rows.length === 1) this.memberText = r.rows[0].mbmNumber
        else return this.searchMode.set('prefix')
      } catch (e) {
        return this.notify.notify(errorMessage(e), 'info')
      }
    }
    await this.load()
  }

  protected async picked(r: MemberRow): Promise<void> {
    this.searchMode.set(null)
    this.memberText = r.mbmNumber
    await this.load()
  }

  protected async step(next: boolean): Promise<void> {
    try {
      const params = new HttpParams({ fromObject: { member: this.memberText, next } })
      this.memberText = (await firstValueFrom(this.http.get<{ member: string }>(`${BASE}/step`, { params }))).member
      await this.load()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdEditMBM_Click: the endorsement status becomes available. */
  protected edit(): void {
    this.editing.set(true)
    this.endorseChanged()
  }

  /** cboEndorseStatus_LostFocus: the desktop's warnings (it does not block). */
  protected endorseChanged(): void {
    const status = this.m()?.status
    if (this.endtType === 'Y' && status !== 'C') this.notify.notify('No Re-Activate allowed ! ', 'error')
    if (this.endtType !== 'Y' && status === 'C') this.notify.notify('Only Re-Activate allow ! ', 'error')
  }

  /** EffdateExpDateChg: new premium / MCO; a different MCO asks " Do you want to change member's MCO Fee?". */
  protected async periodChanged(): Promise<void> {
    if (!this.amend() || !parseDmy(this.p.effDate) || !parseDmy(this.p.expDate)) return
    try {
      const r = await firstValueFrom(this.http.post<{ fees: { coverId: string; ageCode: string; adultChild: string; premium: number; mco: number }[]; warnings: string[] }>(
        `${BASE}/fees`, { member: this.m()!.mbmNumber, effDate: this.p.effDate, expDate: this.p.expDate }))
      for (const w of r.warnings) this.notify.notify(w, 'error')
      const principal = r.fees.find((f) => f.coverId === '00')
      if (!principal) return
      this.p.ageCode = principal.ageCode
      this.p.adultChild = principal.adultChild
      this.p.premium = principal.premium
      const changeMco = principal.mco !== this.p.mcoFee && (await this.notify.ask(" Do you want to change member's MCO Fee?"))
      if (changeMco) this.p.mcoFee = principal.mco
      this.covers.update((list) =>
        list.map((c) => {
          const f = r.fees.find((x) => x.coverId === c.coverId)
          return f ? { ...c, ageCode: f.ageCode, adultChild: f.adultChild, premium: f.premium, mco: changeMco ? f.mco : c.mco } : c
        }))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected pickCover(c: Cover): void {
    this.cover.set(c)
    this.ce = { plan: c.plan, cardType: c.cardType, exclusion: c.exclusion, remarks: c.remarks }
  }

  /** CmdUpdate_Click: the cover's exclusion, remarks, plan and card type (saved with the amendment). */
  protected async updateCover(): Promise<void> {
    const c = this.cover()
    if (!c) return this.notify.notify('Please select covered persons to update !', 'error')
    if (!c.ic.trim()) return this.notify.notify('Please Enter Covered Person Identification', 'error')
    if (!(await this.notify.ask('Are you sure?'))) return
    this.covers.update((list) => list.map((x) => (x.coverId === c.coverId ? { ...x, ...this.ce, exclusion: this.ce.exclusion.toUpperCase(), remarks: this.ce.remarks.toUpperCase() } : x)))
    this.cover.set(null)
  }

  /** CmdCancelSupp_Click: flags the cover for cancellation (K). */
  protected async cancelCover(): Promise<void> {
    const c = this.cover()
    if (!c) return this.notify.notify('Please Choose a Record', 'error')
    if (!(await this.notify.ask('Are you sure?'))) return
    this.covers.update((list) => list.map((x) => (x.coverId === c.coverId ? { ...x, cancel: true } : x)))
  }

  /** SSTab1_Click: the history tabs load on first view. */
  protected async tabChanged(i: number): Promise<void> {
    const m = this.m()
    if (!m) return
    try {
      if (i === 2) {
        const rows = (await firstValueFrom(this.http.get<Record<string, unknown>[]>(`${BASE}/amend-history`, { params: new HttpParams({ fromObject: { member: m.mbmNumber } }) }))) ?? []
        this.amends.set(rows.map((r, n) => ({ ...r, _i: n })))
      }
      if (i === 3 && this.history().length === 0) {
        const rows = (await firstValueFrom(this.http.get<Record<string, unknown>[]>(`${BASE}/mbm-history`, { params: new HttpParams({ fromObject: { ic: m.ic } }) }))) ?? []
        this.history.set(rows.map((r, n) => ({ ...r, _i: n })))
      }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CmdSave_Click. */
  protected async save(): Promise<void> {
    const m = this.m()
    if (!m) return this.notify.notify('Please select a record to edit!', 'error')
    if (!(await this.notify.ask('Are you sure?'))) return
    this.busy.set(true)
    try {
      const body = {
        member: m.mbmNumber, endtType: this.endtType, endtBordxDate: this.endtBordx, endtEffDate: this.endtEff, ...this.p,
        covers: this.covers().map((c) => ({ coverId: c.coverId, exclusion: c.exclusion, remarks: c.remarks, plan: c.plan, cardType: c.cardType, premium: c.premium, mco: c.mco, cancel: !!c.cancel })),
      }
      this.notify.notify((await firstValueFrom(this.http.post<{ message: string }>(BASE, body))).message, 'success')
      await this.load()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private async load(): Promise<void> {
    this.busy.set(true)
    this.reset()
    try {
      const m = await firstValueFrom(this.http.get<Member>(BASE, { params: new HttpParams({ fromObject: { member: this.memberText } }) }))
      this.m.set(m)
      this.covers.set(m.covers)
      this.p = {
        policyNo: m.policyNo, plan: m.plan, status: m.status, dataStatus: m.dataStatus || 'A', effDate: formatDate(m.effDate), expDate: formatDate(m.expDate),
        cardType: m.cardType, email: m.email, mobile: m.mobile, bankName: m.bankName, bankAccount: m.bankAccount, exclusion: m.exclusion, cisRemarks: m.cisRemarks,
        premium: m.premium, mcoFee: m.mcoFee, ageCode: m.ageCode, adultChild: m.adultChild,
      }
      if (m.notice) this.notify.notify(m.notice, 'error')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  private reset(): void {
    this.m.set(null)
    this.covers.set([])
    this.cover.set(null)
    this.amends.set([])
    this.history.set([])
    this.editing.set(false)
    this.endtType = 'A'
    this.endtEff = ''
    this.endtBordx = ''
    this.tab = 0
    this.p = this.emptyPrincipal()
  }

  private emptyPrincipal() {
    return {
      policyNo: '', plan: '', status: 'A', dataStatus: 'A', effDate: '', expDate: '', cardType: '', email: '', mobile: '', bankName: '', bankAccount: '',
      exclusion: '', cisRemarks: '', premium: 0, mcoFee: 0, ageCode: '', adultChild: '',
    }
  }
}
