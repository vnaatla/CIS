import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { ActivatedRoute, Router } from '@angular/router'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { CasePendingModalComponent, PendingDetails } from './case-pending.modal'
import { GP, IMMUNIZATION, Item, LABORATORY, OTHERS, PROCEDURES, SP } from './clinic-invoice.page'
import { CaseListModalComponent, CaseListRow, IcdRow, IcdSearchModalComponent } from './invoice-popups'
import { MemberInfoModalComponent } from './member-info.modal'
import { WorksheetModalComponent, WorksheetResult, WorksheetRow } from './worksheet.modal'

const BASE = '/api/invoice/adjustment'
type Row = Record<string, unknown>

interface AdjustmentCase {
  case: Row
  claim: Row | null
  paid: boolean
  posted: boolean
  bordxed: boolean
  editMode: 'Full' | 'DiagnosisOnly' | 'None'
  editMessage: string | null
  pendingReasonIds: string[]
  pendingReasons: LookupOption[]
  specialistCoPay: number | null
  email: string
  specialRemark: string
  btbg: boolean
  member: { member: { mbmNumber: string; coverId: string; patientName: string }; annualLimit: string; visitLimit: string; balanceLimit: string;
    accumAmount: number; specLimit: string; specBalance: string; asoInd: string } | null
  worksheet: { rows: WorksheetRow[]; claim: Row | null }
  messages: string[]
}

/** SearchCASOP columns of A7..A33 (A29 co-pay and A31 room & board come from the claim). */
const A_COLUMNS: Record<number, string> = {
  7: 'CLMDental', 8: 'CLMMaternity', 9: 'CLMPRVCheckup', 10: 'CLMXRay', 11: 'CLMUltraSound', 12: 'CLMECG', 13: 'CLMDressing', 14: 'CLMNebulizer', 15: 'CLMID',
  16: 'CLMTSSTO', 17: 'CLMMinorSurgery', 18: 'CLMPOthers', 19: 'CLMBloodDrawing', 20: 'CLMBloodGlucose', 21: 'CLMUrine', 22: 'CLMPapSmear', 23: 'CLMLOthers',
  24: 'CLMDTPPolio', 25: 'CLMDTPPolioHIB', 26: 'CLMHepatitisB', 27: 'CLMMMR', 28: 'CLMIOthers', 30: 'CLMVision', 32: 'CLMTyphiadJabNoDed', 33: 'CLMOthersNoDed',
}
const B_COLUMNS = ['', 'GPFee', 'GPLTMed', 'GPInj', 'GPMed', 'SPFee', 'SPLTMed', 'SPInj', 'SPMed']
const money = (v: number | null | undefined) => (v == null ? '' : formatMoney(v))
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const n = (v: unknown): number | null => (v == null || s(v) === '' || Number.isNaN(Number(v)) ? null : Number(v))
const dmy = (v: unknown) => (v ? formatDate(String(v)) : '')

/** Invoice > Clinic > Adjustment (frmInvoiceAdj, frmCasePending, frmWorksheet): change a saved case. */
@Component({
  selector: 'app-invoice-adjustment-page',
  imports: [
    FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, CaseListModalComponent, IcdSearchModalComponent,
    MemberInfoModalComponent, WorksheetModalComponent, CasePendingModalComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="afterBordx ? 'Invoice Adjustment (After Bordx)' : afterPost ? 'Invoice Adjustment (After Posting)' : 'Invoice Adjustment'" [subtitle]="afterBordx ? 'Replaces frmInvoiceAdjAftBordx' : afterPost ? 'Replaces frmInvoiceAdjAfterPost' : 'Replaces frmInvoiceAdj, frmCasePending'">
      <div class="screen-form">
        <fieldset aria-label="Case">
          <legend>Case</legend>
          <app-field label="Case No"><input name="caseNo" data-autofocus maxlength="20" [(ngModel)]="caseNo" (keydown.enter)="go()" /></app-field>
          <app-action-bar label="Case actions">
            <button appButton variant="primary" [busy]="busy()" (click)="go()">Go</button>
            <button appButton (click)="searching.set(true)">Search</button>
            <button appButton [disabled]="!c()?.member" (click)="info.set(true)">MBM Info</button>
            <button appButton [disabled]="!c()" (click)="exclusions()">Exclusions</button>
          </app-action-bar>
        </fieldset>

        @if (c(); as x) {
          <fieldset aria-label="Status" [disabled]="mode() !== 'Full'">
            <legend>{{ x.case['CASNumber'] }} · {{ x.case['MBMnumber'] }}-{{ x.case['CoverID'] }} · {{ x.member?.member?.patientName }}</legend>
            <app-field label="Status">
              <select name="status" [(ngModel)]="xe.status"><option value="">--Select--</option><option value="O">O - Open</option><option value="C">C - Close</option><option value="D">D - Delete</option></select>
            </app-field>
            <app-field label="Claimability">
              <select name="claimability" [(ngModel)]="xe.claimability"><option value="">--Select--</option><option value="A">A - Accepted</option><option value="R">R - Repudiated</option><option value="C">C - Cancelled</option><option value="S">S - Special Approve</option></select>
            </app-field>
            <app-field label="Pending">
              <select name="pending" [(ngModel)]="xe.pending" (ngModelChange)="pendingChanged()"><option value="Y">Y - Yes</option><option value="N">N - No</option></select>
            </app-field>
            @if (!reduced) {
              <app-action-bar label="Pending actions"><button appButton [disabled]="xe.pending !== 'Y'" (click)="openPending('')">Pending Reasons</button></app-action-bar>
            }
            <app-field label="Pay To">
              <select name="payTo" [(ngModel)]="xe.payTo"><option value="C">C - CLINIC</option><option value="M">M - MEMBER</option></select>
            </app-field>
            <app-field label="Accepted"><select name="memoClaim" [(ngModel)]="xe.memoClaim"><option value=""></option><option value="Y">Y-Yes</option><option value="N">N-No</option></select></app-field>
            <app-field label="Referral Letter"><select name="refLetter" [(ngModel)]="xe.refLetter"><option value=""></option><option value="N">N-No</option><option value="Y">Y-Yes</option></select></app-field>
            <app-field label="Pharmacy"><select name="pharmacy" [(ngModel)]="xe.pharmacy"><option value=""></option><option value="Y">Y - Yes</option><option value="N">N - No</option></select></app-field>
            <app-field label="Repudiation Reason" [wide]="true"><input name="repReason" maxlength="1000" [(ngModel)]="xe.repReason" /></app-field>
            <app-field label="Repudiation Amt"><input class="num" name="repAmount" [(ngModel)]="xe.repAmount" /></app-field>
            <app-field label="Mgmt Notes" [wide]="true"><input name="mgmtNotes" maxlength="2000" [(ngModel)]="xe.mgmtNotes" /></app-field>
            <app-field label="OP Batch"><input name="opBatch" [value]="s(x.case['OPBatchNo'])" disabled /></app-field>
            <app-field label="Bordx"><input name="bordx" [value]="x.bordxed ? 'True' : 'False'" disabled /></app-field>
            <app-field label="Payment"><input name="paid" [value]="x.paid ? 'True' : 'False'" disabled /></app-field>
          </fieldset>

          <fieldset aria-label="Invoice" [disabled]="mode() !== 'Full'">
            <legend>Invoice</legend>
            <app-field label="Inv. No"><input name="invoiceNo" maxlength="50" [(ngModel)]="e.invoiceNo" /></app-field>
            <app-field label="Inv Date"><app-date-input name="invoiceDate" ariaLabel="Invoice Date" [(ngModel)]="e.invoiceDate" /></app-field>
            <app-field label="Rec Date"><app-date-input name="receiptDate" ariaLabel="Receipt Date" [(ngModel)]="e.receiptDate" /></app-field>
            <app-field label="Clinic"><input name="clinicCode" maxlength="10" [(ngModel)]="e.clinicCode" /></app-field>
            <app-field label="Clinic Name"><input name="clinicName" [value]="s(x.case['CLNName'])" disabled /></app-field>
            @if (['NP', 'NPS'].includes(e.clinicCode.trim().toUpperCase())) {
              <app-field label="NP Clinic"><input name="npClinic" [(ngModel)]="xe.npClinic" /></app-field>
            }
            <app-field label="Vst Date"><app-date-input name="visitDate" ariaLabel="Visit Date" [(ngModel)]="e.visitDate" /></app-field>
            <app-field label="Vst Time"><input name="visitTime" maxlength="5" [(ngModel)]="e.visitTime" /></app-field>
            <app-field label="Bill No"><input name="billNo" [(ngModel)]="e.billNo" /></app-field>
            <app-field label="MC"><input class="num" name="mcDays" [ngModel]="e.mcDays" (ngModelChange)="e.mcDays = num($event)" /></app-field>
            <app-field label="FFS"><select name="ffs" [(ngModel)]="e.ffs"><option value="Y">Y - Yes</option><option value="N">N - No</option></select></app-field>
            <app-field label="Mth"><input name="processMonth" [(ngModel)]="e.processMonth" /></app-field>
          </fieldset>

          <fieldset aria-label="Diagnosis" [disabled]="mode() === 'None'">
            <legend>Diagnosis &amp; Remark</legend>
            @for (i of [0, 1, 2, 3, 4]; track i) {
              <app-field [label]="'Diagnosis ' + (i + 1)">
                <input [attr.name]="'diag' + i" maxlength="20" [ngModel]="e.diagnoses[i]" (ngModelChange)="e.diagnoses[i] = $event" />
                <button appButton type="button" (click)="icdFor.set(i)">…</button>
              </app-field>
            }
            <app-field label="Others Diagnosis"><input name="diagnosisOther" [(ngModel)]="e.diagnosisOther" /></app-field>
            <app-field label="Remarks" [wide]="true"><input name="remark" maxlength="500" [(ngModel)]="e.remark" /></app-field>
          </fieldset>

          <fieldset aria-label="Items" [disabled]="mode() !== 'Full'">
            <legend>Items</legend>
            @for (it of fees; track it[0]) {
              <app-field [label]="it[1]"><input class="num" [attr.name]="'b' + it[0]" [ngModel]="e.b[it[0]]" (ngModelChange)="e.b[it[0]] = num($event)" /></app-field>
            }
            @for (it of items; track it[0]) {
              <app-field [label]="it[1]"><input class="num" [attr.name]="'a' + it[0]" [ngModel]="e.a[it[0]]" (ngModelChange)="e.a[it[0]] = num($event)" /></app-field>
            }
            <app-field label="Co-Pymt"><input class="num" name="a29" [ngModel]="e.a[29]" (ngModelChange)="e.a[29] = num($event)" /></app-field>
            <app-field label="Total Bill"><input class="num" name="totalBill" [value]="money(total())" disabled /></app-field>
            <app-field label="Bill Amt"><input class="num" name="clinicBillAmount" [(ngModel)]="e.clinicBillAmount" /></app-field>
            <app-field label="Adj. Ann. Lmt"><input class="num" name="adjAnnLmt" [ngModel]="e.adjustedAnnualLimit" (ngModelChange)="e.adjustedAnnualLimit = num($event)" /></app-field>
            <app-field label="Limit Exceeded"><input class="num" name="limitExceeded" [ngModel]="e.limitExceeded" (ngModelChange)="e.limitExceeded = num($event)" /></app-field>
            @if (x.btbg) {
              <app-field label="BTBG"><input class="num" name="btbg" [ngModel]="xe.btbgAmount" (ngModelChange)="xe.btbgAmount = num($event)" /></app-field>
              <app-field label="BTBG Remark"><input name="btbgRemark" [(ngModel)]="xe.btbgRemark" /></app-field>
            }
          </fieldset>

          <app-action-bar label="Adjustment actions">
            <button appButton variant="primary" [disabled]="editing()" (click)="edit()">Edit</button>
            <button appButton [disabled]="mode() !== 'Full'" [busy]="busy()" (click)="compute()">Compute</button>
            <button appButton variant="primary" [disabled]="!editing()" [busy]="busy()" (click)="update()">{{ afterPost ? 'Create' : 'Update' }}</button>
            @if (!reduced) {
              <button appButton [disabled]="!c()" (click)="voidBordx()">Void Bordx</button>
              <button appButton [disabled]="!c()" (click)="claimsApp()">Claims App</button>
            }
          </app-action-bar>
          @if (result()) {
            <p class="totals" aria-label="Result" data-result>{{ result() }}</p>
          }

          @if (!reduced) {
          <fieldset aria-label="Audit review">
            <legend>Audit Review</legend>
            <app-field label="Query" [wide]="true"><input name="auditQuery" [(ngModel)]="audit.query" /></app-field>
            <app-field label="Answer" [wide]="true"><input name="auditAnswer" [(ngModel)]="audit.answer" /></app-field>
            <app-field label="Cost Avoidance"><input name="costAvoidance" [(ngModel)]="audit.costAvoidance" /></app-field>
            <app-field label="Audit Remarks" [wide]="true"><input name="auditRemarks" [(ngModel)]="audit.remarks" /></app-field>
            <app-field label="Audit By"><input name="auditBy" [(ngModel)]="audit.auditBy" /></app-field>
            <app-field label="Audit Date"><app-date-input name="auditDate" ariaLabel="Audit Date" [(ngModel)]="audit.auditDate" /></app-field>
            <app-action-bar label="Audit actions"><button appButton (click)="saveAudit()">Update Audit</button></app-action-bar>
          </fieldset>
          }
        }
      </div>
    </app-page-shell>

    @if (searching()) {
      <app-case-list-modal (picked)="picked($event)" (closed)="searching.set(false)" />
    }
    @if (info() && c()?.member; as mb) {
      <app-member-info-modal [mbmNumber]="mb.member.mbmNumber" [coverId]="mb.member.coverId" (closed)="info.set(false)" />
    }
    @if (icdFor() !== null) {
      <app-icd-search-modal (picked)="icdPicked($event)" (closed)="icdFor.set(null)" />
    }
    @if (worksheet(); as w) {
      <app-worksheet-modal [worksheet]="w" (closed)="worksheet.set(null)" />
    }
    @if (pendingOpen()) {
      <app-case-pending-modal [reasons]="c()?.pendingReasons ?? []" [details]="pending" [editable]="editing()" [note]="pendingNote()"
        (saved)="pendingSaved($event)" (closed)="pendingOpen.set(false)" />
    }
  `,
})
export class InvoiceAdjustmentPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  private readonly router = inject(Router)
  /** frmInvoiceAdjAftBordx (route data afterBordx): its own API and no pending / audit / void / claims approval. */
  protected readonly afterBordx = inject(ActivatedRoute).snapshot.data['afterBordx'] === true
  /** frmInvoiceAdjAfterPost (route data afterPost): Create makes a sub-case of a case posted to SUN. */
  protected readonly afterPost = inject(ActivatedRoute).snapshot.data['afterPost'] === true
  /** The reduced screens have no pending / audit / void / claims approval. */
  protected readonly reduced = this.afterBordx || this.afterPost
  private readonly base = this.afterBordx ? '/api/invoice/adjustment-after-bordx' : this.afterPost ? '/api/invoice/adjustment-after-posting' : BASE

  protected readonly money = money
  protected readonly s = s
  protected readonly fees: Item[] = [...GP, ...SP]
  protected readonly items: Item[] = [...PROCEDURES, ...LABORATORY, ...IMMUNIZATION, ...OTHERS]
  protected readonly c = signal<AdjustmentCase | null>(null)
  protected readonly editing = signal(false)
  protected readonly mode = signal<'Full' | 'DiagnosisOnly' | 'None' | 'View'>('View')
  protected readonly busy = signal(false)
  protected readonly searching = signal(false)
  protected readonly info = signal(false)
  protected readonly icdFor = signal<number | null>(null)
  protected readonly worksheet = signal<WorksheetResult | null>(null)
  protected readonly pendingOpen = signal(false)
  protected readonly pendingNote = signal('')
  protected readonly result = signal('')
  protected readonly version = signal(0)
  protected caseNo = ''
  protected e = this.emptyEntry()
  protected xe = this.emptyExtra()
  protected pending: PendingDetails = { reasonIds: [], reasonNames: [], questions: [], remark: '' }
  protected audit = { query: '', answer: '', costAvoidance: '', remarks: '', auditBy: '', auditDate: '' }

  /** getTotal: every amount except the co-pay. */
  protected readonly total = computed(() => {
    this.version()
    let sum = 0
    for (let i = 1; i <= 8; i++) sum += this.e.b[i] ?? 0
    for (let i = 7; i <= 33; i++) if (i !== 29) sum += this.e.a[i] ?? 0
    return Math.round(sum * 100) / 100
  })

  protected num(v: unknown): number | null {
    this.version.update((x) => x + 1)
    return n(String(v ?? '').replace(/,/g, ''))
  }

  /** CmdGo_Click: InvoiceINFO into the form. */
  protected async go(): Promise<void> {
    this.caseNo = this.caseNo.trim().toUpperCase()
    if (!this.caseNo) return
    this.busy.set(true)
    this.result.set('')
    this.editing.set(false)
    this.mode.set('View')
    try {
      const x = await firstValueFrom(this.http.get<AdjustmentCase>(`${this.base}/${encodeURIComponent(this.caseNo)}`))
      this.fill(x)
      this.c.set(x)
      for (const m of x.messages) this.notify.notify(m, 'info')
    } catch (err) {
      this.c.set(null)
      this.notify.notify(errorMessage(err), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private fill(x: AdjustmentCase): void {
    const r = x.case
    const cl = x.claim ?? {}
    const e = this.emptyEntry()
    e.invoiceNo = s(r['InvoiceNo'])
    e.invoiceDate = dmy(r['InvoiceDate'])
    e.receiptDate = dmy(r['InvoiceReceiptDate'])
    e.receiptMode = s(r['InvoiceReceiptMode'])
    e.clinicCode = s(r['ClinicCode'])
    e.mbmNumber = s(r['MBMnumber'])
    e.coverId = s(r['CoverID'])
    e.visitDate = dmy(r['DateVisit'])
    e.visitTime = s(r['TimeVisit']).slice(0, 5)
    e.billNo = s(r['BillNo'])
    e.mcDays = n(r['MCDays'])
    e.diagnoses = [1, 2, 3, 4, 5].map((i) => s(r[`CASTreatReason${i}`]))
    e.diagnosisOther = s(r['CASDiagnosis5'])
    e.remark = s(r['CasRemark'])
    for (let i = 1; i <= 8; i++) e.b[i] = n(r[B_COLUMNS[i]])
    for (const [i, col] of Object.entries(A_COLUMNS)) e.a[Number(i)] = n(r[col])
    e.a[31] = n(cl['ClmNonGPFee'])
    e.a[29] = x.specialistCoPay ?? n(cl['CLMCoPayCLNColl'])
    e.clinicBillAmount = s(r['INVBillAmt']) ? String(Number(r['INVBillAmt'])) : ''
    e.totalInvoiceAmount = s(r['INVAmount'])
    e.ffs = s(cl['CLMMCOProcStatus']) || 'N'
    e.processMonth = s(cl['CLMProcMth']) ? `${s(cl['CLMProcMth'])}/${s(cl['CLMProcYR'])}` : ''
    e.adjustedAnnualLimit = n(cl['CLMAdjAnnLmt'])
    e.limitExceeded = n(cl['CLMLimitExceeded'])
    e.pOthDesc = s(r['CLMPOthersDesc'])
    e.lOthDesc = s(r['CLMLOthersDesc'])
    e.iOthDesc = s(r['CLMIOthersDesc'])
    e.medications = [1, 2, 3, 4, 5, 6, 7].map((i) => ({ type: s(r[`MedInjType${i}`]), days: s(r[`MedInjDays${i}`]), qty: s(r[`MedInjQty${i}`]), rate: '' }))
    this.e = e
    const xe = this.emptyExtra()
    xe.status = s(r['CASStatus'])
    xe.claimability = s(r['Claimability'])
    xe.pending = s(r['Pending']) || 'N'
    xe.payTo = s(r['INVPayTo']) || 'C'
    xe.memoClaim = s(cl['CLMBordxMemo'])
    xe.refLetter = s(cl['CLMRefLetter'])
    xe.pharmacy = s(r['IsPharmacy'])
    xe.npClinic = s(r['NPClinic'])
    xe.repReason = s(r['RepudiationReason'])
    xe.repAmount = s(r['CASRepAmt'])
    xe.mgmtNotes = s(r['CASMgmtNotes'])
    xe.btbgAmount = n(r['CLMBTBGAmt'])
    xe.btbgRemark = s(r['BTBGRemarks'])
    for (const [i, col] of Object.entries(A_COLUMNS)) xe.aGst[Number(i)] = n(r[`${col}GSTAmt`])
    for (let i = 5; i <= 8; i++) xe.bGst[i] = n(r[`${B_COLUMNS[i]}GSTAmt`])
    xe.medCharges = [1, 2, 3, 4, 5, 6, 7].map((i) => n(r[`MEDCHARGES${i}`]))
    this.xe = xe
    this.pending = {
      reasonIds: [...x.pendingReasonIds],
      reasonNames: x.pendingReasons.filter((p) => x.pendingReasonIds.includes(p.code)).map((p) => p.name),
      questions: s(r['qatohos']) ? s(r['qatohos']).split('|') : [],
      remark: s(r['addpendnote']),
    }
    this.audit = {
      query: s(r['AuditQuery']), answer: s(r['AuditAnswer']), costAvoidance: s(r['AuditDiscountReceived']), remarks: s(r['AuditRemarks']),
      auditBy: s(r['AuditBy']), auditDate: dmy(r['AuditDate']),
    }
    this.version.update((v) => v + 1)
  }

  /** CmdEdit_Click. */
  protected edit(): void {
    const x = this.c()
    if (!x) return
    if (x.editMode === 'None') return this.notify.notify(x.editMessage ?? '', 'info')
    if (x.editMessage) this.notify.notify(x.editMessage, 'info')
    this.mode.set(x.editMode)
    this.editing.set(true)
    if (x.editMode === 'Full' && x.specialRemark) this.notify.notify(x.specialRemark, 'info')
  }

  /** cboPending_Click: back to N with pending details left. */
  protected pendingChanged(): void {
    const p = this.pending
    if (this.xe.pending === 'N' && (p.reasonIds.length || p.questions.length || p.remark))
      this.openPending("Case pending status changed to 'N', but pending details not cleared! Please check")
  }

  protected openPending(note: string): void {
    this.pendingNote.set(note)
    this.pendingOpen.set(true)
  }

  protected pendingSaved(d: PendingDetails): void {
    this.pending = d
    this.pendingOpen.set(false)
  }

  protected async compute(): Promise<void> {
    this.busy.set(true)
    try {
      this.worksheet.set(await firstValueFrom(this.http.post<WorksheetResult>(`${this.base}/compute`, this.request(null, null))))
    } catch (err) {
      this.notify.notify(errorMessage(err), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** CmdUpdate_Click (+ ChkData's "Are you sure ? " and the rejection / pending e-mail prompts). */
  protected async update(): Promise<void> {
    const x = this.c()
    if (!x) return
    if (this.mode() === 'Full' && !(await this.notify.ask('Are you sure ? '))) return
    let notes: string | null = null
    let email: string | null = null
    const ubbs = !this.reduced && /^(UB|BS)/.test(s(x.case['MBMnumber']))
    const rejection = this.xe.status === 'D' && this.xe.claimability === 'C' && this.xe.pending === 'N'
    const pendingOpen = this.xe.status === 'O' && this.xe.claimability === 'A' && this.xe.pending === 'Y'
    const changed = (rejection && this.xe.repReason.trim() !== s(x.case['RepudiationReason']))
      || (pendingOpen && this.pending.reasonIds.join(',') !== x.pendingReasonIds.join(','))
    if (this.mode() === 'Full' && ubbs && this.xe.payTo === 'M' && changed) {
      const kind = rejection ? 'Rejection' : 'Pending'
      notes = window.prompt(`Please Provide Additional Notes If Any.\nClick OK Button to Send ${kind} Email to member. Cancel to Abort..!`)
      if (notes !== null && !x.email && !s(x.case['OPBatchNo']))
        email = window.prompt('Member Email is not found in our Records. Please Provide Email Id If Any.\nClick OK Button to add Email Address. Cancel to Abort..!')
    }
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ message: string; opBatchNo: string; warnings: string[]; newCaseNo: string | null }>(this.base, this.request(notes, email)))
      for (const w of r.warnings) this.notify.notify(w, 'info')
      const done = r.newCaseNo ? `${r.message} New Case No : ${r.newCaseNo}` : r.message || 'Diagnosis and remark updated.'
      this.result.set(done)
      this.notify.notify(done, 'success')
      this.editing.set(false)
      this.mode.set('View')
      await this.go()
      this.result.set(done)
    } catch (err) {
      this.notify.notify(errorMessage(err), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected async saveAudit(): Promise<void> {
    if (!(await this.notify.ask('Are you sure want update Audit Details? '))) return
    try {
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${this.base}/${encodeURIComponent(this.caseNo)}/audit`, this.audit))
      this.notify.notify(r.message, 'success')
    } catch (err) {
      this.notify.notify(errorMessage(err), 'error')
    }
  }

  protected async voidBordx(): Promise<void> {
    try {
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${this.base}/${encodeURIComponent(this.caseNo)}/void-bordx`, {}))
      this.notify.notify(r.message, 'success')
      await this.go()
    } catch (err) {
      this.notify.notify(errorMessage(err), 'error')
    }
  }

  protected async claimsApp(): Promise<void> {
    try {
      await firstValueFrom(this.http.post(`${this.base}/${encodeURIComponent(this.caseNo)}/claims-app`, {}))
      this.notify.notify('Claims approval set.', 'success')
    } catch (err) {
      this.notify.notify(errorMessage(err), 'error')
    }
  }

  protected exclusions(): void {
    const grp = s(this.c()?.case['MBMGroupCompany'])
    if (!grp) return this.notify.notify('Please Check Group Company! ', 'info')
    void this.router.navigate(['/invoice/general-exclusions/enquiry'], { queryParams: { groupCompany: grp } })
  }

  protected picked(r: CaseListRow): void {
    this.searching.set(false)
    this.caseNo = r.casNumber
    void this.go()
  }

  protected icdPicked(r: IcdRow): void {
    const i = this.icdFor()
    this.icdFor.set(null)
    if (i !== null) this.e.diagnoses[i] = r.reasonCode
  }

  private request(notifyNotes: string | null, notifyEmail: string | null) {
    return {
      caseNo: this.caseNo,
      entry: { ...this.e, medications: this.e.medications },
      extra: {
        ...this.xe, pendingReasonIds: this.pending.reasonIds, pendingReasonNames: this.pending.reasonNames, questions: this.pending.questions,
        addPendRemark: this.pending.remark, notifyNotes, notifyEmail,
      },
    }
  }

  private emptyEntry() {
    return {
      invoiceNo: '', invoiceDate: '', receiptDate: '', receiptMode: '', clinicCode: '', totalInvoiceAmount: '', mbmNumber: '', coverId: '00', visitDate: '',
      visitTime: '', billNo: '', mcDays: null as number | null, diagnoses: ['', '', '', '', ''], diagnosisDescriptions: ['', '', '', '', ''], diagnosisOther: '',
      a: Array.from({ length: 34 }, (): number | null => null), b: Array.from({ length: 9 }, (): number | null => null), clinicBillAmount: '', ffs: 'N',
      processMonth: '', claimability: '', remark: '', adjustedAnnualLimit: null as number | null, excess: null as number | null,
      limitExceeded: null as number | null, pOthDesc: '', lOthDesc: '', iOthDesc: '',
      medications: Array.from({ length: 7 }, () => ({ type: '', days: '', qty: '', rate: '' })),
    }
  }

  private emptyExtra() {
    return {
      status: '', claimability: '', pending: 'N', payTo: 'C', memoClaim: '', refLetter: '', pharmacy: '', npClinic: '', repReason: '', repAmount: '',
      mgmtNotes: '', btbgAmount: null as number | null, btbgRemark: '', aGst: Array.from({ length: 34 }, (): number | null => null),
      bGst: Array.from({ length: 9 }, (): number | null => null), medCharges: Array.from({ length: 7 }, (): number | null => null),
    }
  }
}
