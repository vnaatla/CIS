import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { NgTemplateOutlet } from '@angular/common'
import { FormsModule } from '@angular/forms'
import { Router } from '@angular/router'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { openBlob } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { MemberRow, MemberSearchModalComponent, MemberSearchMode } from './member-search.modal'

const BASE = '/api/membership/enquiry'

interface Benefit {
  limit: string
  utilised: string
  balance: string
}
interface Cover {
  coverId: string
  name: string
  ic: string
  relationship: string
  dob: string | null
  plan: string
  status: string
  cardType: string
  premium: number
  mco: number
  exclusion: string
  remarks: string
  premReduction: number | null
  premEffective: number | null
  mcoRefund: number | null
  mcoEffective: number | null
  gp: Benefit
  specialist: Benefit
  dental: Benefit
  optical: Benefit
  visitLimit: string
  mepLimit: number
  dentalLimit: number
  dentalExtLimit: number
  specLimit: number
  firstSpecLimit: number
  ageBand: string
  adultChild: string
  effDate: string | null
  expDate: string | null
  ffs: string
  topupAmt: string
  spTopupAmt: string
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
  address: string[]
  country: string
  postCode: string
  state: string
  subsidiary: string
  department: string
  groupCompany: string
  position: string
  hisRemarks: string
  specialRemark: string
  dob: string | null
  adultChild: string
  effDate: string | null
  expDate: string | null
  ageBand: string
  cardType: string
  exclusion: string
  cisRemarks: string
  invoiceDate: string | null
  invoiceNo: string
  premium: number
  mcoFee: number
  product: string
  limitStatus: string
  limitStatusText: string
  topupAmt: string
  spTopupAmt: string
  ffsAmt: string
  accumAmt: string
  annualLimit: string
  balanceLimit: string
  visitLimit: string
  gp: Benefit
  specialist: Benefit
  maternity: Benefit
  dental: Benefit
  optical: Benefit
  mepLimit: number
  dentalLimit: number
  specLimit: number
  dentalExtLimit: number
  firstSpecLimit: number
  covers: Cover[]
  flexCare: boolean
  warnings: string[]
  principalNote: string
  dependantNote: string
}
interface FlexRow {
  item: string
  capped: string
  utilised: string
  balance: string
  highlight: boolean
}
interface Flex {
  found: boolean
  coverId: string
  cappedItems: FlexRow[]
  items: FlexRow[]
  totals: FlexRow[]
  family: FlexRow[]
  outpatient: FlexRow[]
  otherItems: FlexRow[]
}
interface Case {
  caseNo: string
  mbmNumber: string
  name: string
  relationship: string
  visitDate: string | null
  claimability: string
  reason: string
  status: string
  bill: number
  bordx: number
  ffs: number
}
interface Cases {
  gp: Case[]
  gpVisits: number
  gpBill: number
  gpBordx: number
  gpFfs: number
  specialist: Case[]
  spVisits: number
  spBill: number
  spBordx: number
}

const STATUS: Record<string, string> = { A: 'A - Active', C: 'C - Cancel', D: 'D - Deleted', S: 'S - Suspended' }
const money = (v: number | null | undefined) => (v == null ? '' : formatMoney(v))

/** Membership > MBM Enquiry (FrmMBMCLNEnq): a member's details, limits and utilisation, covers, histories, cases and benefit schedule (read-only). */
@Component({
  selector: 'app-mbm-enquiry-page',
  imports: [FormsModule, NgTemplateOutlet, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent, TabsComponent, TabDirective, MemberSearchModalComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="MBM Enquiry" subtitle="Replaces FrmMBMCLNEnq">
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
            <button appButton [disabled]="!m()" (click)="exclusions()">Exclusions</button>
          </app-action-bar>
        </fieldset>

        @if (m(); as mm) {
          <app-tabs label="MBM Enquiry" [(active)]="tab" (activeChange)="tabChanged($event)">
            <ng-template appTab="Principal Details">
              @if (mm.principalNote) {
                <p class="notice limit-note" aria-label="Limit note">{{ mm.principalNote }}</p>
              }
              <fieldset aria-label="Principal">
                <legend>Principal</legend>
                <app-field label="Name"><input name="name" [value]="mm.name" disabled /></app-field>
                <app-field label="IC No"><input name="ic" [value]="mm.ic" disabled /></app-field>
                <app-field label="Policy No"><input name="policyNo" [value]="mm.policyNo" disabled /></app-field>
                <app-field label="Payor"><input name="payCode" [value]="mm.payCode" disabled /></app-field>
                <app-field label="Insured Type"><input name="insType" [value]="mm.insType" disabled /></app-field>
                <app-field label="Plan"><input name="plan" [value]="mm.plan" disabled /></app-field>
                <app-field label="Product"><input name="product" [value]="mm.product" disabled /></app-field>
                <app-field label="Status"><input name="status" [value]="statusText(mm.status)" disabled /></app-field>
                <app-field label="Exp. Status"><input name="expStatus" [value]="mm.expiryStatus === 'N' ? 'Y - Yes' : 'N - No'" disabled /></app-field>
                <app-field label="Data Status"><input name="dataStatus" [value]="mm.dataStatus === 'D' ? 'D - Deleted' : 'A - Active'" disabled /></app-field>
                <app-field label="Eff. Date"><input name="effDate" [value]="fmt(mm.effDate)" disabled /></app-field>
                <app-field label="Exp. Date"><input name="expDate" [value]="fmt(mm.expDate)" disabled /></app-field>
                <app-field label="DOB / Age Band"><input name="dob" [value]="fmt(mm.dob) + ' / ' + mm.ageBand + ' (' + mm.adultChild + ')'" disabled /></app-field>
                <app-field label="Card Type"><input name="cardType" [value]="mm.cardType" disabled /></app-field>
                <app-field label="Limit Status"><input name="limitStatus" [value]="mm.limitStatusText" disabled /></app-field>
                <app-field label="Group Company"><input name="groupCompany" [value]="mm.groupCompany" disabled /></app-field>
                <app-field label="Subsidiary"><input name="subsidiary" [value]="mm.subsidiary" disabled /></app-field>
                <app-field label="Department"><input name="department" [value]="mm.department" disabled /></app-field>
                <app-field label="Position"><input name="position" [value]="mm.position" disabled /></app-field>
                <app-field label="Invoice No / Date"><input name="invoice" [value]="mm.invoiceNo + (mm.invoiceDate ? ' / ' + fmt(mm.invoiceDate) : '')" disabled /></app-field>
                <app-field label="Address" [wide]="true"><input name="address" [value]="addressText(mm)" disabled /></app-field>
                <app-field label="Exclusion" [wide]="true"><textarea name="exclusion" rows="2" [value]="mm.exclusion" disabled></textarea></app-field>
                <app-field label="HIS Remarks" [wide]="true"><textarea name="hisRemarks" rows="2" [value]="mm.hisRemarks" disabled></textarea></app-field>
                <app-field label="CIS Remarks" [wide]="true"><textarea name="cisRemarks" rows="2" [value]="mm.cisRemarks" disabled></textarea></app-field>
                <app-field label="Special Remark" [wide]="true">
                  <textarea name="specialRemark" rows="2" [class.highlight]="!!mm.specialRemark" [value]="mm.specialRemark" disabled></textarea>
                </app-field>
              </fieldset>
              <fieldset aria-label="Fees and limits">
                <legend>Fees and Limits</legend>
                @for (f of feeFields(mm); track f[0]) {
                  <app-field [label]="f[0]"><input class="num" [attr.name]="f[0]" [value]="f[1]" disabled /></app-field>
                }
              </fieldset>
              <div class="table-wrap limits-grid"><table class="grid" aria-label="Limits and utilisation">
                <thead><tr><th scope="col">Item</th>@for (b of benefits(mm); track b[0]) {<th scope="col" class="num">{{ b[0] }}</th>}</tr></thead>
                <tbody>
                  <tr><th scope="row">Limit</th>@for (b of benefits(mm); track b[0]) {<td class="num">{{ b[1].limit }}</td>}</tr>
                  <tr><th scope="row">Utilised</th>@for (b of benefits(mm); track b[0]) {<td class="num">{{ b[1].utilised }}</td>}</tr>
                  <tr><th scope="row">Balance</th>@for (b of benefits(mm); track b[0]) {<td class="num">{{ b[1].balance }}</td>}</tr>
                </tbody>
              </table></div>
            </ng-template>

            <ng-template appTab="Supplementary">
              @if (mm.dependantNote) {
                <p class="notice limit-note" aria-label="Dependant limit note">{{ mm.dependantNote }}</p>
              }
              <app-data-grid caption="Covered persons" [columns]="coverColumns" [rows]="mm.covers" [rowKey]="coverKey" [pageSize]="0" [selectable]="true"
                [selectedKey]="cover()?.coverId ?? null" (rowSelected)="pickCover($event)" />
              @if (cover(); as c) {
                <fieldset aria-label="Covered person">
                  <legend>Covered Person {{ c.coverId }}</legend>
                  <app-field label="Name"><input name="covName" [value]="c.name" disabled /></app-field>
                  <app-field label="IC"><input name="covIc" [value]="c.ic" disabled /></app-field>
                  <app-field label="Relationship"><input name="covRel" [value]="c.relationship" disabled /></app-field>
                  <app-field label="DOB / Age Band"><input name="covDob" [value]="fmt(c.dob) + ' / ' + c.ageBand + ' (' + c.adultChild + ')'" disabled /></app-field>
                  <app-field label="Plan"><input name="covPlan" [value]="c.plan" disabled /></app-field>
                  <app-field label="Card Type"><input name="covCard" [value]="c.cardType" disabled /></app-field>
                  <app-field label="Eff. / Exp. Date"><input name="covPeriod" [value]="fmt(c.effDate) + ' - ' + fmt(c.expDate)" disabled /></app-field>
                  @for (f of coverFields(c); track f[0]) {
                    <app-field [label]="f[0]"><input class="num" [attr.name]="'cov ' + f[0]" [value]="f[1]" disabled /></app-field>
                  }
                  <app-field label="Exclusion" [wide]="true"><textarea name="covExclusion" rows="2" [value]="c.exclusion" disabled></textarea></app-field>
                  <app-field label="Remarks" [wide]="true"><textarea name="covRemarks" rows="2" [value]="c.remarks" disabled></textarea></app-field>
                </fieldset>
                <div class="table-wrap limits-grid cover-limits"><table class="grid" aria-label="Covered person limits">
                  <thead><tr><th scope="col">Item</th>@for (b of coverBenefits(c); track b[0]) {<th scope="col" class="num">{{ b[0] }}</th>}</tr></thead>
                  <tbody>
                    <tr><th scope="row">Limit</th>@for (b of coverBenefits(c); track b[0]) {<td class="num">{{ b[1].limit }}</td>}</tr>
                    <tr><th scope="row">Utilised</th>@for (b of coverBenefits(c); track b[0]) {<td class="num">{{ b[1].utilised }}</td>}</tr>
                    <tr><th scope="row">Balance</th>@for (b of coverBenefits(c); track b[0]) {<td class="num">{{ b[1].balance }}</td>}</tr>
                  </tbody>
                </table></div>
              }
            </ng-template>

            <ng-template appTab="Adjustment History">
              <app-data-grid caption="Adjustment history" [columns]="amendColumns" [rows]="amends()" [rowKey]="indexKey" />
            </ng-template>

            <ng-template appTab="MBM History">
              <app-data-grid caption="MBM history" [columns]="historyColumns" [rows]="history()" [rowKey]="indexKey" />
            </ng-template>

            <ng-template appTab="Case">
              @if (cases(); as cs) {
                <h3 class="section-title">GP Cases</h3>
                <app-data-grid caption="GP cases" [columns]="caseColumns" [rows]="cs.gp" [rowKey]="caseKey" />
                <p class="totals" aria-label="GP totals">
                  Total Visits: <strong>{{ cs.gpVisits }}</strong> · Bill Amount: <strong>{{ money(cs.gpBill) }}</strong> ·
                  Bordx Amount: <strong>{{ money(cs.gpBordx) }}</strong> · FFS: <strong>{{ money(cs.gpFfs) }}</strong>
                </p>
                <h3 class="section-title">Specialist Cases</h3>
                <app-data-grid caption="Specialist cases" [columns]="caseColumns" [rows]="cs.specialist" [rowKey]="caseKey" />
                <p class="totals" aria-label="Specialist totals">
                  Total Visits: <strong>{{ cs.spVisits }}</strong> · Bill Amount: <strong>{{ money(cs.spBill) }}</strong> ·
                  Bordx Amount: <strong>{{ money(cs.spBordx) }}</strong>
                </p>
              } @else {
                <p class="muted">Loading cases…</p>
              }
            </ng-template>

            <ng-template appTab="Benefit Schedule">
              @if (!mm.groupCompany) {
                <p class="muted">The member has no group company.</p>
              } @else if (!filesConfigured()) {
                <p class="muted">The policy wording folder (CLNPolicyWording) is not configured on the API server (Files:Folders:PolicyWording).</p>
              } @else if (files().length === 0) {
                <p class="muted">No benefit schedule files for {{ mm.groupCompany }}.</p>
              } @else {
                <ul class="file-list" aria-label="Benefit schedule files">
                  @for (f of files(); track f) {
                    <li><button type="button" class="link" (click)="open(f)">{{ f }}</button></li>
                  }
                </ul>
              }
            </ng-template>
            @if (mm.flexCare) {
              <ng-template appTab="FlexCare">
                @if (flex(); as fx) {
                  @if (!fx.found) {
                    <p class="muted">No FlexCare figures for this member.</p>
                  } @else {
                    <ng-container *ngTemplateOutlet="flexTable; context: { caption: 'Capped items', rows: fx.cappedItems, capped: true }" />
                    <ng-container *ngTemplateOutlet="flexTable; context: { caption: 'Items utilised', rows: fx.items, capped: false }" />
                    <ng-container *ngTemplateOutlet="flexTable; context: { caption: 'FlexCare totals', rows: fx.totals, capped: false }" />
                  }
                } @else {
                  <p class="muted">Loading FlexCare…</p>
                }
              </ng-template>
              <ng-template appTab="Family Benefits">
                @if (!cover()) {
                  <p class="muted">Select a covered person on the Supplementary tab.</p>
                } @else if (family(); as fm) {
                  <p class="notice">Member Name :{{ cover()!.name }} - CoverID:{{ cover()!.coverId }}</p>
                  @if (fm.coverId === '01') {
                    <ng-container *ngTemplateOutlet="flexTable; context: { caption: 'Family benefits', rows: fm.family, capped: false }" />
                  } @else {
                    <p class="muted">Please select cover id 01 for available Limits &amp; Balance</p>
                  }
                  <ng-container *ngTemplateOutlet="flexTable; context: { caption: 'Out-patient utilised', rows: fm.outpatient, capped: false }" />
                  <ng-container *ngTemplateOutlet="flexTable; context: { caption: 'Other items utilised', rows: fm.otherItems, capped: false }" />
                } @else {
                  <p class="muted">Loading family benefits…</p>
                }
              </ng-template>
            }
          </app-tabs>
        }
      </div>
    </app-page-shell>

    <ng-template #flexTable let-caption="caption" let-rows="rows" let-capped="capped">
      <div class="table-wrap limits-grid flex-grid"><table class="grid" [attr.aria-label]="caption">
        <thead><tr><th scope="col">No.</th><th scope="col">Item</th>@if (capped) {<th scope="col" class="num">Capped Limit</th>}<th scope="col" class="num">{{ capped ? 'Utilised' : 'Amount' }}</th>@if (capped) {<th scope="col" class="num">Balance</th>}</tr></thead>
        <tbody>
          @for (r of rows; track $index) {
            <tr><td>{{ $index + 1 }}</td><th scope="row">{{ r.item }}</th>@if (capped) {<td class="num">{{ r.capped }}</td>}<td class="num" [class.highlight]="r.highlight">{{ r.utilised }}</td>@if (capped) {<td class="num">{{ r.balance }}</td>}</tr>
          }
        </tbody>
      </table></div>
    </ng-template>

    @if (searchMode(); as mode) {
      <app-member-search-modal [mode]="mode" [initialText]="memberText" (picked)="picked($event)" (closed)="searchMode.set(null)" />
    }
  `,
})
export class MbmEnquiryPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  private readonly router = inject(Router)

  protected readonly money = money
  protected readonly m = signal<Member | null>(null)
  protected readonly amends = signal<Record<string, unknown>[]>([])
  protected readonly history = signal<Record<string, unknown>[]>([])
  protected readonly cases = signal<Cases | null>(null)
  protected readonly cover = signal<Cover | null>(null)
  protected readonly flex = signal<Flex | null>(null)
  protected readonly family = signal<Flex | null>(null)
  protected readonly files = signal<string[]>([])
  protected readonly filesConfigured = signal(true)
  protected readonly busy = signal(false)
  protected readonly searchMode = signal<MemberSearchMode | null>(null)
  protected readonly coverKey = (c: Cover) => c.coverId
  protected readonly caseKey = (c: Case) => c.caseNo
  protected readonly indexKey = (r: Record<string, unknown>) => String(r['_i'])
  protected memberText = ''
  protected tab = 0
  private loaded = new Set<number>()

  protected readonly coverColumns: GridColumn<Cover>[] = [
    { key: 'name', label: 'Name' }, { key: 'coverId', label: 'ID' }, { key: 'ic', label: 'IC' }, { key: 'relationship', label: 'Rel' },
    { key: 'dob', label: 'DOB', value: (c) => formatDate(c.dob) }, { key: 'plan', label: 'Plan' }, { key: 'status', label: 'Status' },
    { key: 'cardType', label: 'Card Type' }, { key: 'premium', label: 'Premium', align: 'right', value: (c) => money(c.premium) },
    { key: 'mco', label: 'MCO', align: 'right', value: (c) => money(c.mco) }, { key: 'exclusion', label: 'Exclusion' }, { key: 'remarks', label: 'Remarks' },
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
  protected readonly caseColumns: GridColumn<Case>[] = [
    { key: 'caseNo', label: 'Case No' }, { key: 'mbmNumber', label: 'Member No' }, { key: 'name', label: 'Name' }, { key: 'relationship', label: 'Rel' },
    { key: 'visitDate', label: 'Visit Date', value: (c) => formatDate(c.visitDate) }, { key: 'claimability', label: 'Claimability' }, { key: 'reason', label: 'Reason' },
    { key: 'status', label: 'Status' }, { key: 'bill', label: 'Bill Amt', align: 'right', value: (c) => money(c.bill) },
    { key: 'bordx', label: 'Bordx Amt', align: 'right', value: (c) => money(c.bordx) }, { key: 'ffs', label: 'FFS', align: 'right', value: (c) => money(c.ffs) },
  ]

  protected fmt(v: string | null | undefined): string {
    return formatDate(v)
  }

  protected statusText(s: string): string {
    return STATUS[s] ?? s
  }

  protected addressText(mm: Member): string {
    return [...mm.address, mm.postCode, mm.state, mm.country].filter((x) => x).join(', ')
  }

  protected feeFields(mm: Member): [string, string][] {
    return [
      ['Premium', money(mm.premium)], ['MCO Fee', money(mm.mcoFee)], ['Topup Amt', mm.topupAmt], ['SP Topup Amt', mm.spTopupAmt], ['FFS Amt', mm.ffsAmt],
      ['Accum. Amt', mm.accumAmt], ['Annual Limit', mm.annualLimit], ['Balance Limit', mm.balanceLimit], ['Visit Limit', mm.visitLimit],
      ['MEP Limit', money(mm.mepLimit)], ['Dental Limit', money(mm.dentalLimit)], ['Dental Ext. Limit', money(mm.dentalExtLimit)], ['Spec. Limit', money(mm.specLimit)],
      ['First Spec. Limit', money(mm.firstSpecLimit)],
    ]
  }

  protected coverFields(c: Cover): [string, string][] {
    return [
      ['Premium', money(c.premium)], ['Prem. Reduction', money(c.premReduction)], ['Prem. Effective', money(c.premEffective)], ['MCO', money(c.mco)],
      ['MCO Refund', money(c.mcoRefund)], ['MCO Effective', money(c.mcoEffective)], ['FFS Amt', c.ffs], ['Topup Amt', c.topupAmt], ['SP Topup Amt', c.spTopupAmt],
      ['Visit Limit', c.visitLimit], ['MEP Limit', money(c.mepLimit)], ['Dental Limit', money(c.dentalLimit)], ['Dental Ext. Limit', money(c.dentalExtLimit)],
      ['Spec. Limit', money(c.specLimit)], ['First Spec. Limit', money(c.firstSpecLimit)],
    ]
  }

  /** lstCovAdd_Click's FGridTopup: maternity is always N/A for a covered person. */
  protected coverBenefits(c: Cover): [string, Benefit][] {
    return [['GP', c.gp], ['SP', c.specialist], ['Maternity', { limit: 'N/A', utilised: 'N/A', balance: 'N/A' }], ['Dental', c.dental], ['Optical', c.optical]]
  }

  protected benefits(mm: Member): [string, Benefit][] {
    return [['GP', mm.gp], ['SP', mm.specialist], ['Maternity', mm.maternity], ['Dental', mm.dental], ['Optical', mm.optical]]
  }

  /** cmdShow_Click: at least 8 characters; numbers not starting DC / SL (and not "…*nn") go through frmSearchMBM05 first. */
  protected async show(): Promise<void> {
    this.memberText = this.memberText.trim().toUpperCase()
    const t = this.memberText
    if (t.length < 8) return this.notify.notify('Please Enter Membership No. Atleast first 8 characters required.', 'error')
    if (!t.startsWith('DC') && !t.startsWith('SL') && t.length <= 11 && t.charAt(11) !== '*') {
      try {
        const r = await firstValueFrom(this.http.get<{ rows: MemberRow[] }>('/api/membership/search/prefix', { params: new HttpParams({ fromObject: { memberNo: t } }) }))
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

  /** lstCovAdd_Click: the cover's panel; FlexCare members reload the Family Benefits tab for it (ShowFlexiMBM). */
  protected pickCover(c: Cover): void {
    this.cover.set(c)
    this.family.set(null)
    this.loaded.delete(7)
  }

  /** cmdExclusions_Click: General Exclusions (Enquiry) for the member's group company. */
  protected exclusions(): void {
    const grp = this.m()?.groupCompany.trim()
    if (!grp) return this.notify.notify('Please Check Group Company! ', 'error')
    void this.router.navigate(['/invoice/general-exclusions/enquiry'], { queryParams: { groupCompany: grp } })
  }

  /** SSTab1_Click: I / G plans have no Supplementary tab; the other tabs load on first view. */
  protected async tabChanged(i: number): Promise<void> {
    const m = this.m()
    if (!m) return
    if (i === 1 && (m.insType === 'I' || m.insType === 'G')) {
      this.notify.notify('This Plan does not have supplementary', 'error')
      this.tab = 0
      return
    }
    if (this.loaded.has(i)) return
    this.loaded.add(i)
    try {
      if (i === 2) this.amends.set(this.indexed(await firstValueFrom(this.http.get<Record<string, unknown>[]>(`${BASE}/amend-history`, { params: this.p({ member: m.mbmNumber }) }))))
      if (i === 3) this.history.set(this.indexed(await firstValueFrom(this.http.get<Record<string, unknown>[]>(`${BASE}/mbm-history`, { params: this.p({ ic: m.ic }) }))))
      if (i === 6) this.flex.set(await firstValueFrom(this.http.get<Flex>(`${BASE}/flex`, { params: this.p({ member: m.mbmNumber, coverId: '00' }) })))
      if (i === 7) {
        const c = this.cover()
        if (!c) this.loaded.delete(7)
        else this.family.set(await firstValueFrom(this.http.get<Flex>(`${BASE}/flex`, { params: this.p({ member: m.mbmNumber, coverId: c.coverId }) })))
      }
      if (i === 4) this.cases.set(await firstValueFrom(this.http.get<Cases>(`${BASE}/cases`, { params: this.p({ member: m.mbmNumber }) })))
      if (i === 5 && m.groupCompany) {
        const r = await firstValueFrom(this.http.get<{ configured: boolean; files: string[] }>(`${BASE}/policy-files`, { params: this.p({ groupCompany: m.groupCompany }) }))
        this.filesConfigured.set(r.configured)
        this.files.set(r.files)
      }
    } catch (e) {
      this.loaded.delete(i)
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** lstPolicyWording_Click: open the file. */
  protected async open(file: string): Promise<void> {
    try {
      const params = this.p({ groupCompany: this.m()!.groupCompany, fileName: file })
      openBlob(await firstValueFrom(this.http.get(`${BASE}/policy-file`, { params, responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  private async load(): Promise<void> {
    this.busy.set(true)
    this.m.set(null)
    this.amends.set([])
    this.history.set([])
    this.cases.set(null)
    this.cover.set(null)
    this.flex.set(null)
    this.family.set(null)
    this.files.set([])
    this.loaded = new Set()
    this.tab = 0
    try {
      const m = await firstValueFrom(this.http.get<Member>(BASE, { params: this.p({ member: this.memberText }) }))
      this.m.set(m)
      for (const w of m.warnings) this.notify.notify(w, 'error')
      if (m.specialRemark) this.notify.notify(m.specialRemark, 'info')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  private p(o: Record<string, string>): HttpParams {
    return new HttpParams({ fromObject: o })
  }

  private indexed(rows: Record<string, unknown>[] | null): Record<string, unknown>[] {
    return (rows ?? []).map((r, n) => ({ ...r, _i: n }))
  }
}
