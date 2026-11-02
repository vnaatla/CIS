import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { MemberRow, MemberSearchModalComponent } from '../membership/member-search.modal'
import { IcdRow, IcdSearchModalComponent } from './invoice-popups'
import { MemberInfoModalComponent, MemberSupp } from './member-info.modal'
import { WorksheetModalComponent, WorksheetResult } from './worksheet.modal'

const BASE = '/api/invoice/clinic'

interface ClinicInfo {
  code: string
  name: string
  clinicType: string
  ffsStatus: string
  panel: boolean
}
interface InvoiceMember {
  member: MemberSupp & { cardType: string; dentalStatus: string; [k: string]: unknown }
  memberType: string
  annualLimit: string
  visitLimit: string
  balanceLimit: string
  accumAmount: number
  gpLimit: number
  spLimit: number
  specLimit: string
  specAccum: number
  specBalance: string
  coPay: number
  asoInd: string
  warnings: string[]
}
interface ClinicLine {
  amounts: number[]
  total: number
  clinicBill: number
  actCoPay: number
  worksheet: WorksheetResult
}

/** An item field of the form: A index and label (the desktop's captions). */
type Item = [number, string]
const PROCEDURES: Item[] = [[10, 'X-Ray'], [11, 'Utr Sound'], [12, 'ECG'], [13, 'Dressing'], [14, 'Nebulizer'], [15, 'I&D'], [16, 'T&S/Sto'], [17, 'M. Surg.'], [18, 'Oth']]
const LABORATORY: Item[] = [[19, 'B. Drawing'], [20, 'B. Glucose'], [21, 'Urine'], [22, 'Pap Smear'], [23, 'Oth']]
const IMMUNIZATION: Item[] = [[24, 'DTP&Polio'], [25, 'DTP&Polio&HIB'], [26, 'Hepatitis B'], [27, 'MMR'], [28, 'Oth']]
const OTHERS: Item[] = [[7, 'Dental'], [8, 'Maternity'], [9, 'PRV Chk.'], [30, 'Vision'], [31, 'Room & Board'], [32, 'Typhoid Jab'], [33, 'Others']]
const GP: Item[] = [[1, 'GP Fee + LT Med.'], [2, 'GP Med.'], [3, 'GP Inj.'], [4, 'GP LT Med.']]
const SP: Item[] = [[5, 'Specialist Fee + LT Med.'], [6, 'Specialist Med.'], [7, 'Specialist Inj.'], [8, 'Specialist LT Med.']]

const money = (v: number | null | undefined) => (v == null ? '' : formatMoney(v))

/** Invoice > Clinic > Clinic (frmInvoice + frmWorksheet): enter a clinic's invoice case by case, verify and save. */
@Component({
  selector: 'app-clinic-invoice-page',
  imports: [
    FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, MemberSearchModalComponent,
    IcdSearchModalComponent, MemberInfoModalComponent, WorksheetModalComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Invoice Batch Processing - Clinic" subtitle="Replaces frmInvoice, frmWorksheet">
      <div class="screen-form">
        <fieldset aria-label="Invoice">
          <legend>Invoice</legend>
          <app-field label="Inv. No" [required]="true"><input name="invoiceNo" data-autofocus maxlength="50" [(ngModel)]="e.invoiceNo" /></app-field>
          <app-field label="Inv Date" [required]="true"><app-date-input name="invoiceDate" ariaLabel="Invoice Date" [(ngModel)]="e.invoiceDate" /></app-field>
          <app-field label="Rec Date" [required]="true"><app-date-input name="receiptDate" ariaLabel="Receipt Date" [(ngModel)]="e.receiptDate" /></app-field>
          <app-field label="Rec Mode">
            <select name="receiptMode" [(ngModel)]="e.receiptMode">
              <option value=""></option>
              @for (r of modes(); track r.code) {
                <option [value]="r.code">{{ r.code }} - {{ r.name }}</option>
              }
            </select>
          </app-field>
          <app-field label="Clinic" [required]="true"><input name="clinicCode" maxlength="10" [(ngModel)]="e.clinicCode" (blur)="loadClinic()" /></app-field>
          <app-field label="Clinic Name"><input name="clinicName" [value]="clinic()?.name ?? ''" disabled /></app-field>
          <app-field label="Tot. Amt" [required]="true"><input class="num" name="totalInvoiceAmount" [(ngModel)]="e.totalInvoiceAmount" /></app-field>
        </fieldset>

        <fieldset aria-label="Member">
          <legend>Member</legend>
          <app-field label="Member No" [required]="true">
            <input name="mbmNumber" maxlength="50" [(ngModel)]="mbmText" (keydown.enter)="go()" />
          </app-field>
          <app-action-bar label="Member actions">
            <button appButton variant="primary" [busy]="busy()" (click)="go()">Go</button>
            <button appButton (click)="searching.set(true)">Search</button>
            <button appButton [disabled]="!member()" (click)="info.set(true)">MBM Info</button>
          </app-action-bar>
          @if (member(); as mb) {
            <app-field label="Name"><input name="patientName" [value]="mb.member.patientName" disabled /></app-field>
            <app-field label="IC No"><input name="ic" [value]="mb.member.patientIc" disabled /></app-field>
            <app-field label="Policy No"><input name="policyNo" [value]="mb.member.policyNo" disabled /></app-field>
            <app-field label="Status"><input name="status" [value]="mb.member.dataStatus" disabled /></app-field>
            <app-field label="Pol. Sta"><input name="polStatus" [value]="mb.member.policyStatus" disabled /></app-field>
            <app-field label="Eff Date"><input name="effDate" [value]="fmt(mb.member.effDate)" disabled /></app-field>
            <app-field label="Exp Date"><input name="expDate" [value]="fmt(mb.member.expDate)" disabled /></app-field>
            <app-field label="Plan Code"><input name="planCode" [value]="mb.member.planCode" disabled /></app-field>
            <app-field label="Ann. Lmt"><input class="num" name="annualLimit" [value]="mb.annualLimit" disabled /></app-field>
            <app-field label="Vst Lmt"><input class="num" name="visitLimit" [value]="mb.visitLimit" disabled /></app-field>
            <app-field label="Acc Amt"><input class="num" name="accumAmount" [value]="money(mb.accumAmount)" disabled /></app-field>
            <app-field label="Bal. Lmt"><input class="num" name="balanceLimit" [value]="mb.balanceLimit" disabled /></app-field>
            <app-field label="Specialist Lmt"><input class="num" name="specLimit" [value]="mb.specLimit" disabled /></app-field>
            <app-field label="Specialist Bal"><input class="num" name="specBalance" [value]="mb.specBalance" disabled /></app-field>
            <app-field label="SP/GP Limit"><input class="num" name="gpLimit" [value]="money(mb.gpLimit) + ' / ' + money(mb.spLimit)" disabled /></app-field>
          }
        </fieldset>

        @if (member()) {
          <fieldset aria-label="Case">
            <legend>Case</legend>
            <app-field label="Vst Date" [required]="true"><app-date-input name="visitDate" ariaLabel="Visit Date" [(ngModel)]="e.visitDate" /></app-field>
            <app-field label="Vst Time"><input name="visitTime" placeholder="hh:mm" maxlength="5" [(ngModel)]="e.visitTime" /></app-field>
            <app-field label="Bill No"><input name="billNo" maxlength="50" [(ngModel)]="e.billNo" /></app-field>
            <app-field label="MC"><input class="num" name="mcDays" [ngModel]="e.mcDays" (ngModelChange)="e.mcDays = num($event)" /></app-field>
            @for (n of [0, 1, 2, 3, 4]; track n) {
              <app-field [label]="'Diagnosis ' + (n + 1)">
                <input [attr.name]="'diag' + n" maxlength="20" [ngModel]="e.diagnoses[n]" (ngModelChange)="e.diagnoses[n] = $event" (blur)="diagnosis(n)" />
                <button appButton type="button" (click)="icdFor.set(n)">…</button>
              </app-field>
              <app-field label="Description"><input [attr.name]="'diagDesc' + n" [value]="e.diagnosisDescriptions[n]" disabled /></app-field>
            }
            @if (otherDiagnosis()) {
              <app-field label="Others Diagnosis" [required]="true"><input name="diagnosisOther" maxlength="200" [(ngModel)]="e.diagnosisOther" /></app-field>
            }
            <app-field label="FFS">
              <select name="ffs" [(ngModel)]="e.ffs"><option value=""></option><option value="Y">Y - Yes</option><option value="N">N - No</option></select>
            </app-field>
            <app-field label="Mth">
              <select name="processMonth" [(ngModel)]="e.processMonth">
                <option value=""></option>
                @for (mo of months(); track mo) {
                  <option [value]="mo">{{ mo }}</option>
                }
              </select>
            </app-field>
            <app-field label="Clm'lity" [required]="true">
              <select name="claimability" [(ngModel)]="e.claimability">
                <option value=""></option><option value="A">A - Accepted</option><option value="M">M - Memo</option><option value="R">R - Rejected</option>
              </select>
            </app-field>
            <app-field label="Remarks" [wide]="true"><input name="remark" maxlength="500" [(ngModel)]="e.remark" /></app-field>
          </fieldset>

          <fieldset aria-label="Items">
            <legend>Items</legend>
            @for (g of [gp, sp]; track $index) {
              @for (it of g; track it[0]) {
                <app-field [label]="it[1]">
                  <input class="num" [attr.name]="'b' + it[0]" [ngModel]="e.b[it[0]]" (ngModelChange)="e.b[it[0]] = num($event)" />
                </app-field>
              }
            }
            @for (group of groups; track group[0]) {
              @for (it of group[1]; track it[0]) {
                <app-field [label]="($index === 0 ? group[0] + ': ' : '') + it[1]">
                  <input class="num" [attr.name]="'a' + it[0]" [ngModel]="e.a[it[0]]" (ngModelChange)="e.a[it[0]] = num($event)" />
                </app-field>
              }
            }
            <app-field label="Proc. Oth Desc"><input name="pOthDesc" [(ngModel)]="e.pOthDesc" /></app-field>
            <app-field label="Lab. Oth Desc"><input name="lOthDesc" [(ngModel)]="e.lOthDesc" /></app-field>
            <app-field label="Imm. Oth Desc"><input name="iOthDesc" [(ngModel)]="e.iOthDesc" /></app-field>
            <app-field label="Co-Pymt"><input class="num" name="a29" [ngModel]="e.a[29]" (ngModelChange)="e.a[29] = num($event)" /></app-field>
            <app-field label="Total Bill"><input class="num" name="totalBill" [value]="money(total())" disabled /></app-field>
            <app-field label="Bill Amt" [required]="true"><input class="num" name="clinicBillAmount" [(ngModel)]="e.clinicBillAmount" /></app-field>
            <app-field label="Adj. Ann. Lmt"><input class="num" name="adjAnnLmt" [ngModel]="e.adjustedAnnualLimit" (ngModelChange)="e.adjustedAnnualLimit = num($event)" /></app-field>
            <app-field label="Total Excess"><input class="num" name="excess" [ngModel]="e.excess" (ngModelChange)="e.excess = num($event)" /></app-field>
            <app-field label="Limit Exceeded"><input class="num" name="limitExceeded" [ngModel]="e.limitExceeded" (ngModelChange)="e.limitExceeded = num($event)" /></app-field>
          </fieldset>

          <fieldset aria-label="Medication & Injection">
            <legend>Medication &amp; Injection</legend>
            @for (md of e.medications; track $index) {
              <app-field [label]="$index + 1 + '. Medication'"><input [attr.name]="'medType' + $index" [(ngModel)]="md.type" /></app-field>
              <app-field label="Days"><input class="num" [attr.name]="'medDays' + $index" maxlength="10" [(ngModel)]="md.days" /></app-field>
              <app-field label="Qty"><input class="num" [attr.name]="'medQty' + $index" maxlength="10" [(ngModel)]="md.qty" /></app-field>
              <app-field label="Price"><input class="num" [attr.name]="'medRate' + $index" [(ngModel)]="md.rate" /></app-field>
            }
          </fieldset>

          <app-action-bar label="Invoice actions">
            <button appButton variant="primary" [busy]="busy()" (click)="add()">Add</button>
            <button appButton [disabled]="!line()" (click)="showWorksheet.set(true)">Compute</button>
            <button appButton [disabled]="!line()" [busy]="busy()" (click)="verify()">Verify</button>
            <button appButton variant="primary" [disabled]="!verified()" [busy]="busy()" (click)="save()">Save</button>
            <button appButton (click)="clearAll()">Clear All</button>
          </app-action-bar>

          @if (line(); as l) {
            <p class="totals" aria-label="Line">
              Bill Amt: <strong>{{ money(l.total) }}</strong> · Final Approved: <strong data-final>{{ money(l.worksheet.finalApproved) }}</strong> ·
              Payable to CLN: <strong>{{ money(l.worksheet.finalPayableToClinic) }}</strong>
              @if (caseNo()) { · Case No : <strong data-case>{{ caseNo() }}</strong> }
            </p>
          }
          @if (errors().length) {
            <div class="table-wrap">
              <table class="grid" aria-label="Verification errors">
                <thead><tr><th>Member No</th><th>Cov</th><th>Remarks</th></tr></thead>
                <tbody>
                  @for (er of errors(); track $index) {
                    <tr><td>{{ member()?.member?.mbmNumber }}</td><td>{{ member()?.member?.coverId }}</td><td>{{ er }}</td></tr>
                  }
                </tbody>
              </table>
            </div>
            <app-action-bar label="Error actions"><button appButton (click)="exportErrors()">Export errors</button></app-action-bar>
          }
        }
      </div>
    </app-page-shell>

    @if (searching()) {
      <app-member-search-modal (picked)="picked($event)" (closed)="searching.set(false)" />
    }
    @if (info() && member(); as mb) {
      <app-member-info-modal [mbmNumber]="mb.member.mbmNumber" [coverId]="mb.member.coverId" (closed)="info.set(false)" />
    }
    @if (icdFor() !== null) {
      <app-icd-search-modal (picked)="icdPicked($event)" (closed)="icdFor.set(null)" />
    }
    @if (showWorksheet() && line(); as l) {
      <app-worksheet-modal [worksheet]="l.worksheet" (closed)="showWorksheet.set(false)" />
    }
  `,
})
export class ClinicInvoicePage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = money
  protected readonly fmt = formatDate
  protected readonly gp = GP
  protected readonly sp = SP
  protected readonly groups: [string, Item[]][] = [['Procedures', PROCEDURES], ['Laboratory', LABORATORY], ['Immunization', IMMUNIZATION], ['Others', OTHERS]]
  protected readonly modes = signal<LookupOption[]>([])
  protected readonly months = signal<string[]>([])
  protected readonly clinic = signal<ClinicInfo | null>(null)
  protected readonly member = signal<InvoiceMember | null>(null)
  protected readonly line = signal<ClinicLine | null>(null)
  protected readonly errors = signal<string[]>([])
  protected readonly verified = signal(false)
  protected readonly caseNo = signal('')
  protected readonly otherDiagnosis = signal(false)
  protected readonly busy = signal(false)
  protected readonly searching = signal(false)
  protected readonly info = signal(false)
  protected readonly showWorksheet = signal(false)
  protected readonly icdFor = signal<number | null>(null)
  protected readonly version = signal(0)
  protected mbmText = ''
  protected e = this.emptyEntry()

  /** getTotal (display only; the API recomputes it). */
  protected readonly total = computed(() => {
    this.version()
    const b = (i: number) => this.e.b[i] ?? 0
    const a = (i: number) => this.e.a[i] ?? 0
    let sum = b(1) + b(2) + b(3) + b(4) + b(5) + b(6) + b(7) + b(8)
    for (let i = 7; i <= 33; i++) if (i !== 29) sum += a(i)
    return Math.round(sum * 100) / 100
  })

  async ngOnInit(): Promise<void> {
    try {
      const s = await firstValueFrom(this.http.get<{ receiptModes: LookupOption[]; processMonths: string[] }>(`${BASE}/setup`))
      this.modes.set(s.receiptModes)
      this.months.set(s.processMonths)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected num(v: unknown): number | null {
    const n = typeof v === 'number' ? v : Number(String(v ?? '').replace(/,/g, ''))
    this.version.update((x) => x + 1)
    this.unverify()
    return String(v ?? '').trim() === '' || Number.isNaN(n) ? null : n
  }

  /** TxtClinicCode_LostFocus. */
  protected async loadClinic(): Promise<void> {
    this.e.clinicCode = (this.e.clinicCode ?? '').trim().toUpperCase()
    this.clinic.set(null)
    if (!this.e.clinicCode) return
    try {
      const c = await firstValueFrom(this.http.get<ClinicInfo>(`${BASE}/clinic/${encodeURIComponent(this.e.clinicCode)}`))
      this.clinic.set(c)
      if (c.ffsStatus) this.e.ffs = c.ffsStatus
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CmdGo_Click. */
  protected async go(): Promise<void> {
    const text = this.mbmText.trim().toUpperCase()
    if (!text) return this.notify.notify('Please Key Member No Or IC Number ! ', 'error')
    this.busy.set(true)
    this.line.set(null)
    this.unverify()
    try {
      const m = await firstValueFrom(this.http.get<InvoiceMember>(`${BASE}/member/${encodeURIComponent(text)}`))
      this.member.set(m)
      this.e.mbmNumber = m.member.mbmNumber
      this.e.coverId = m.member.coverId
      this.e.a[29] = m.coPay
      for (const w of m.warnings) this.notify.notify(w, 'info')
    } catch (e) {
      this.member.set(null)
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected picked(r: MemberRow): void {
    this.searching.set(false)
    this.mbmText = r.coverId && r.coverId !== '00' ? `${r.mbmNumber}-${r.coverId}` : r.mbmNumber
    void this.go()
  }

  /** VerifyDiagnosis (+ the "others" indicator of diagnosis 5). */
  protected async diagnosis(n: number): Promise<void> {
    const code = (this.e.diagnoses[n] ?? '').trim()
    this.e.diagnosisDescriptions[n] = ''
    if (n === 4) this.otherDiagnosis.set(false)
    if (!code) return
    try {
      const d = await firstValueFrom(this.http.get<{ code: string; description: string; otherIndicator: string | null }>(`${BASE}/diagnosis/${encodeURIComponent(code)}`))
      this.e.diagnoses[n] = d.code
      this.e.diagnosisDescriptions[n] = d.description ?? ''
      if (n === 4) this.otherDiagnosis.set(d.otherIndicator?.trim() === 'Y')
      this.version.update((x) => x + 1)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected icdPicked(r: IcdRow): void {
    const n = this.icdFor()
    this.icdFor.set(null)
    if (n === null) return
    this.e.diagnoses[n] = r.reasonCode
    void this.diagnosis(n)
  }

  /** CmdAddRec_Click: the duplicate question is asked like the desktop's Yes / No. */
  protected async add(confirm = false): Promise<void> {
    this.busy.set(true)
    this.unverify()
    try {
      const r = await firstValueFrom(this.http.post<{ line: ClinicLine | null; confirm: string | null }>(`${BASE}/add`, this.request(confirm)))
      if (r.confirm) {
        this.busy.set(false)
        if (await this.notify.ask(r.confirm)) return this.add(true)
        return
      }
      this.line.set(r.line)
      this.confirmed = confirm
      await this.verify()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdVerify_Click. */
  protected async verify(): Promise<void> {
    try {
      const r = await firstValueFrom(this.http.post<{ errors: string[]; message: string | null }>(`${BASE}/verify`, this.request(this.confirmed)))
      this.errors.set(r.errors)
      this.verified.set(true)
      if (r.message) this.notify.notify(r.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdSave_Click. */
  protected async save(): Promise<void> {
    const m = this.member()
    const annual = Number(String(m?.annualLimit ?? '').replace(/,/g, ''))
    const question =
      m?.annualLimit !== 'FULL' && m?.asoInd === 'Y' && annual <= 0 ? 'ASO Limit already Exceed. Still you want to Proceed... ' : 'Are you sure ? '
    if (!(await this.notify.ask(question))) return
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ caseNo: string; message: string }>(BASE, this.request(this.confirmed)))
      this.caseNo.set(r.caseNo)
      this.verified.set(false)
      this.notify.notify(`${r.message} Case No : ${r.caseNo}`, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected exportErrors(): void {
    saveCsv([['Member No', 'Cov', 'Remarks'], ...this.errors().map((x) => [this.e.mbmNumber, this.e.coverId, x])], 'InvoiceErrors.csv')
  }

  protected clearAll(): void {
    this.e = this.emptyEntry()
    this.mbmText = ''
    this.clinic.set(null)
    this.member.set(null)
    this.line.set(null)
    this.caseNo.set('')
    this.otherDiagnosis.set(false)
    this.unverify()
    this.version.update((x) => x + 1)
  }

  private confirmed = false

  private unverify(): void {
    this.verified.set(false)
    this.errors.set([])
  }

  private request(confirmDuplicate: boolean) {
    return { entry: { ...this.e, medications: this.e.medications.filter((x) => x.type || x.rate) }, confirmDuplicate }
  }

  private emptyEntry() {
    return {
      invoiceNo: '', invoiceDate: '', receiptDate: '', receiptMode: '', clinicCode: '', totalInvoiceAmount: '', mbmNumber: '', coverId: '00',
      visitDate: '', visitTime: '', billNo: '', mcDays: null as number | null, diagnoses: ['', '', '', '', ''], diagnosisDescriptions: ['', '', '', '', ''],
      diagnosisOther: '', a: new Array<number | null>(34).fill(null), b: new Array<number | null>(9).fill(null), clinicBillAmount: '', ffs: '',
      processMonth: '', claimability: '', remark: '', adjustedAnnualLimit: null as number | null, excess: null as number | null,
      limitExceeded: null as number | null, pOthDesc: '', lOthDesc: '', iOthDesc: '',
      medications: Array.from({ length: 7 }, () => ({ type: '', days: '', qty: '', rate: '' })),
    }
  }
}
