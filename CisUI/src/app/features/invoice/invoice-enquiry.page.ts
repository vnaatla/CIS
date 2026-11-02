import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { NgTemplateOutlet } from '@angular/common'
import { FormsModule } from '@angular/forms'
import { Router } from '@angular/router'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob, saveCsv } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { CaseListModalComponent, CaseListRow } from './invoice-popups'
import { MemberInfoModalComponent } from './member-info.modal'
import { WorksheetRow } from './worksheet.modal'

const BASE = '/api/invoice/enquiry'
type Row = Record<string, unknown>

interface EnquiryCase {
  case: Row
  fees: Row | null
  member: { member: { mbmNumber: string; coverId: string; patientName: string; planCode: string; effDate: string | null; expDate: string | null };
    annualLimit: string; visitLimit: string; balanceLimit: string; accumAmount: number; specLimit: string; specBalance: string } | null
  worksheet: { rows: WorksheetRow[]; claim: Row | null }
  bordx: Row[]
  warnings: string[]
}

const ITEMS: [string, string][] = [
  ['CLMGPFeeBasicMed', 'GP Fee + LT Med.'], ['CLMGPFeeLTMed', 'GP Med.'], ['CLMGPFeeMedInj', 'GP Inj.'], ['CLMSPFeeBasicMed', 'SP Fee + LT Med.'],
  ['CLMSPFeeLTMed', 'SP Med.'], ['CLMSPFeeMedInj', 'SP Inj.'], ['CLMXRay', 'X-Ray'], ['CLMUltraSound', 'Utr Sound'], ['CLMECG', 'ECG'],
  ['CLMDressing', 'Dressing'], ['CLMNebulizer', 'Nebulizer'], ['CLMID', 'I&D'], ['CLMTSSTO', 'T&S/Sto'], ['CLMMinorSurgery', 'M. Surg.'],
  ['CLMPOthers', 'Proc. Oth'], ['CLMBloodDrawing', 'B. Drawing'], ['CLMBloodGlucose', 'B. Glucose'], ['CLMUrine', 'Urine'], ['CLMPapSmear', 'Pap Smear'],
  ['CLMLOthers', 'Lab. Oth'], ['CLMDTPPolio', 'DTP&Polio'], ['CLMDTPPolioHIB', 'DTP&Polio&HIB'], ['CLMHepatitisB', 'Hepatitis B'], ['CLMMMR', 'MMR'],
  ['CLMIOthers', 'Imm. Oth'], ['CLMDental', 'Dental'], ['CLMMaternity', 'Maternity'], ['CLMPRVCheckup', 'PRV Chk.'], ['CLMVision', 'Vision'],
  ['CLMTyphiadJabNoDed', 'Typhoid Jab'], ['CLMOthersNoDed', 'Others'],
]
const STATUS: Record<string, string> = { O: 'Open', C: 'Close', D: 'Delete' }
const m = (v: unknown) => (typeof v === 'number' ? formatMoney(v) : v == null || v === '' ? '' : String(v))
const d = (v: unknown) => formatDate(v as string | null)

/** Invoice > Clinic > Enquiry (FrmInvoiceEnq): one case, read-only, with its member, worksheet, bordereaux, payments, documents and medication. */
@Component({
  selector: 'app-invoice-enquiry-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, TabsComponent, TabDirective, NgTemplateOutlet, CaseListModalComponent, MemberInfoModalComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Invoice Enquiry" subtitle="Replaces FrmInvoiceEnq">
      <div class="screen-form">
        <fieldset aria-label="Case">
          <legend>Case</legend>
          <app-field label="Case No"><input name="caseNo" data-autofocus maxlength="20" [(ngModel)]="caseNo" (keydown.enter)="go()" /></app-field>
          <app-action-bar label="Case actions">
            <button appButton variant="primary" [busy]="busy()" (click)="go()">Go</button>
            <button appButton (click)="step(false)">Prev</button>
            <button appButton (click)="step(true)">Next</button>
            <button appButton (click)="searching.set(true)">Search</button>
            <button appButton [disabled]="!c()?.member" (click)="info.set(true)">MBM Info</button>
            <button appButton [disabled]="!c()" (click)="exclusions()">Exclusions</button>
          </app-action-bar>
        </fieldset>
        @if (c(); as x) {
          <fieldset aria-label="Invoice">
            <legend>{{ x.case['CASNumber'] }}</legend>
            <app-field label="Member"><input name="member" [value]="x.case['MBMnumber'] + '-' + x.case['CoverID']" disabled /></app-field>
            <app-field label="Patient"><input name="patient" [value]="x.member?.member?.patientName ?? ''" disabled /></app-field>
            <app-field label="Group Company"><input name="groupCompany" [value]="v(x.case['MBMGroupCompany'])" disabled /></app-field>
            <app-field label="Clinic"><input name="clinic" [value]="v(x.case['ClinicCode']) + ' - ' + (v(x.case['NPClinic']) || v(x.case['CLNName']))" disabled /></app-field>
            <app-field label="Inv. No"><input name="invoiceNo" [value]="v(x.case['InvoiceNo'])" disabled /></app-field>
            <app-field label="Inv Date"><input name="invoiceDate" [value]="d(x.case['InvoiceDate'])" disabled /></app-field>
            <app-field label="Vst Date"><input name="visitDate" [value]="d(x.case['DateVisit'])" disabled /></app-field>
            <app-field label="Bill No"><input name="billNo" [value]="v(x.case['BillNo'])" disabled /></app-field>
            <app-field label="Bill Amt"><input class="num" name="billAmt" [value]="m(x.case['INVAmount'])" disabled /></app-field>
            <app-field label="Pay To"><input name="payTo" [value]="v(x.case['INVPayTo']) === 'M' ? 'M-Reimbursement' : 'C-Cash Less'" disabled /></app-field>
            <app-field label="OP Batch"><input name="opBatch" [value]="v(x.case['OPBatchNo'])" disabled /></app-field>
            <app-field label="Status"><input name="status" [value]="status(x.case['CASStatus'])" disabled /></app-field>
            <app-field label="Claimability"><input name="claimability" [value]="v(x.case['Claimability'])" disabled /></app-field>
            <app-field label="Pending"><input name="pending" [value]="v(x.case['Pending'])" disabled /></app-field>
            <app-field label="Diagnosis" [wide]="true">
              <input name="diagnosis" [value]="[x.case['CASTreatReason1'], x.case['CASTreatReason2'], x.case['CASTreatReason3'], x.case['CASTreatReason4'], x.case['CASTreatReason5']].map(v).filter(f).join(', ')" disabled />
            </app-field>
            <app-field label="Remarks" [wide]="true"><input name="remark" [value]="v(x.case['CasRemark'])" disabled /></app-field>
          </fieldset>
          @if (x.member; as mb) {
            <fieldset aria-label="Limits">
              <legend>Limits</legend>
              <app-field label="Ann. Lmt"><input class="num" name="annualLimit" [value]="mb.annualLimit" disabled /></app-field>
              <app-field label="Vst Lmt"><input class="num" name="visitLimit" [value]="mb.visitLimit" disabled /></app-field>
              <app-field label="Acc Amt"><input class="num" name="accum" [value]="m(mb.accumAmount)" disabled /></app-field>
              <app-field label="Bal. Lmt"><input class="num" name="balance" [value]="mb.balanceLimit" disabled /></app-field>
              <app-field label="Specialist Lmt"><input class="num" name="specLimit" [value]="mb.specLimit" disabled /></app-field>
            </fieldset>
          }
          <app-tabs label="Case details" [(active)]="tab" (activeChange)="tabChanged($event)">
            <ng-template appTab="Items">
              <div class="table-wrap">
                <table class="grid" aria-label="Items">
                  <thead><tr><th>Item</th><th class="num">Actual</th></tr></thead>
                  <tbody>
                    @for (it of items; track it[0]) {
                      @if (x.case[it[0]]) { <tr><td>{{ it[1] }}</td><td class="num">{{ m(x.case[it[0]]) }}</td></tr> }
                    }
                  </tbody>
                </table>
              </div>
            </ng-template>
            <ng-template appTab="Worksheet">
              <div class="table-wrap">
                <table class="grid" aria-label="Worksheet">
                  <thead><tr><th>Benfits</th><th class="num">Actual</th><th class="num">Contracted Rate</th><th class="num">Payable to CLN</th>
                    <th class="num">Not Payable to CLN</th><th class="num">Benefit Lmt</th><th class="num">Approved Benefit</th><th class="num">Excess</th></tr></thead>
                  <tbody>
                    @for (r of x.worksheet.rows; track r.code) {
                      <tr><td>{{ r.benefit }}</td><td class="num">{{ m(r.actual) }}</td><td class="num">{{ m(r.contractedRate) }}</td>
                        <td class="num">{{ m(r.payableToClinic) }}</td><td class="num">{{ m(r.notPayableToClinic) }}</td><td class="num">{{ m(r.benefitLimit) }}</td>
                        <td class="num">{{ m(r.approved) }}</td><td class="num">{{ m(r.excess) }}</td></tr>
                    }
                    @if (x.worksheet.claim; as cl) {
                      <tr class="total"><td>Final Approved</td><td class="num">{{ m(cl['CLMTotalActual']) }}</td><td></td>
                        <td class="num">{{ m(cl['CLMPayableByMedix2C']) }}</td><td></td><td></td><td class="num" data-final>{{ m(cl['CLMFinalApprove']) }}</td>
                        <td class="num">{{ m(cl['CLMExcess']) }}</td></tr>
                    }
                  </tbody>
                </table>
              </div>
            </ng-template>
            <ng-template appTab="Bordx">
              <ng-container *ngTemplateOutlet="grid; context: { rows: x.bordx, label: 'Bordereaux' }" />
            </ng-template>
            <ng-template appTab="Payment">
              <ng-container *ngTemplateOutlet="grid; context: { rows: payments(), label: 'Payments' }" />
              <ng-container *ngTemplateOutlet="grid; context: { rows: receipts(), label: 'Receipts' }" />
            </ng-template>
            <ng-template appTab="Documents">
              <div class="table-wrap">
                <table class="grid" aria-label="Documents">
                  <thead><tr><th>Folder</th><th>File</th></tr></thead>
                  <tbody>
                    @for (doc of documents(); track doc.group + doc.fileName) {
                      <tr><td>{{ doc.group }}</td><td><button appButton (click)="open(doc.group, doc.fileName)">{{ doc.fileName }}</button></td></tr>
                    } @empty { <tr class="empty-row"><td colspan="2">No documents.</td></tr> }
                  </tbody>
                </table>
              </div>
            </ng-template>
            <ng-template appTab="Medication Request">
              <ng-container *ngTemplateOutlet="grid; context: { rows: requests(), label: 'Medication requests' }" />
              <ng-container *ngTemplateOutlet="grid; context: { rows: medicines(), label: 'Medicines' }" />
              <app-action-bar label="Medication actions"><button appButton (click)="printRequests()">Print</button></app-action-bar>
            </ng-template>
          </app-tabs>
        }
      </div>
    </app-page-shell>

    <ng-template #grid let-rows="rows" let-label="label">
      <div class="table-wrap">
        <table class="grid" [attr.aria-label]="label">
          <thead><tr>@for (k of keys(rows); track k) { <th>{{ k }}</th> }</tr></thead>
          <tbody>
            @for (r of rows; track $index) { <tr>@for (k of keys(rows); track k) { <td>{{ cell(r[k]) }}</td> }</tr> }
            @empty { <tr class="empty-row"><td>No records found.</td></tr> }
          </tbody>
        </table>
      </div>
    </ng-template>

    @if (searching()) {
      <app-case-list-modal (picked)="picked($event)" (closed)="searching.set(false)" />
    }
    @if (info() && c()?.member; as mb) {
      <app-member-info-modal [mbmNumber]="mb.member.mbmNumber" [coverId]="mb.member.coverId" (closed)="info.set(false)" />
    }
  `,
})
export class InvoiceEnquiryPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  private readonly router = inject(Router)

  protected readonly items = ITEMS
  protected readonly m = m
  protected readonly d = d
  protected readonly c = signal<EnquiryCase | null>(null)
  protected readonly payments = signal<Row[]>([])
  protected readonly receipts = signal<Row[]>([])
  protected readonly requests = signal<Row[]>([])
  protected readonly medicines = signal<Row[]>([])
  protected readonly documents = signal<{ group: string; fileName: string }[]>([])
  protected readonly busy = signal(false)
  protected readonly searching = signal(false)
  protected readonly info = signal(false)
  protected caseNo = ''
  protected tab = 0

  protected v(x: unknown): string {
    return x == null ? '' : String(x).trim()
  }

  protected f(x: string): boolean {
    return x.length > 0
  }

  protected status(x: unknown): string {
    return STATUS[this.v(x)] ?? this.v(x)
  }

  protected keys(rows: Row[]): string[] {
    return rows.length ? Object.keys(rows[0]) : []
  }

  protected cell(x: unknown): string {
    if (typeof x === 'number') return formatMoney(x)
    if (typeof x === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(x)) return formatDate(x)
    return x == null ? '' : String(x)
  }

  /** CmdGo_Click. */
  protected async go(): Promise<void> {
    this.caseNo = this.caseNo.trim().toUpperCase()
    if (!this.caseNo) return
    this.busy.set(true)
    this.c.set(null)
    this.tab = 0
    try {
      const x = await firstValueFrom(this.http.get<EnquiryCase>(`${BASE}/${encodeURIComponent(this.caseNo)}`))
      this.c.set(x)
      for (const w of x.warnings) this.notify.notify(w, 'info')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** CmdNext_Click / CmdPrev_Click. */
  protected async step(next: boolean): Promise<void> {
    try {
      const r = await firstValueFrom(this.http.get<{ casNumber: string }>(`${BASE}/${encodeURIComponent(this.caseNo.trim() || '0')}/step?next=${next}`))
      this.caseNo = r.casNumber
      await this.go()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected picked(r: CaseListRow): void {
    this.searching.set(false)
    this.caseNo = r.casNumber
    void this.go()
  }

  /** cmdExclusions_Click. */
  protected exclusions(): void {
    const grp = this.v(this.c()?.case['MBMGroupCompany'])
    if (!grp) return this.notify.notify('Please Check Group Company! ', 'info')
    void this.router.navigate(['/invoice/general-exclusions/enquiry'], { queryParams: { groupCompany: grp } })
  }

  /** SSTab1_Click: the payment, document and medication tabs load when opened. */
  protected async tabChanged(t: number): Promise<void> {
    const cas = encodeURIComponent(this.caseNo)
    try {
      if (t === 3) {
        const r = await firstValueFrom(this.http.get<{ payments: Row[]; receipts: Row[] }>(`${BASE}/${cas}/payments`))
        this.payments.set(r.payments)
        this.receipts.set(r.receipts)
      } else if (t === 4) {
        this.documents.set(await firstValueFrom(this.http.get<{ group: string; fileName: string }[]>(`${BASE}/${cas}/documents`)))
      } else if (t === 5) {
        const r = await firstValueFrom(this.http.get<{ requests: Row[]; medicines: Row[] }>(`${BASE}/${cas}/medication`))
        this.requests.set(r.requests)
        this.medicines.set(r.medicines)
      }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async open(group: string, name: string): Promise<void> {
    try {
      const url = `${BASE}/${encodeURIComponent(this.caseNo)}/documents/file?group=${encodeURIComponent(group)}&name=${encodeURIComponent(name)}`
      saveBlob(await firstValueFrom(this.http.get(url, { responseType: 'blob' })), name)
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  /** cmdPrint_Click: the medication requests. */
  protected printRequests(): void {
    const rows = this.requests()
    if (!rows.length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    const keys = Object.keys(rows[0])
    saveCsv([keys, ...rows.map((r) => keys.map((k) => this.cell(r[k])))], 'MedicationRequests.csv')
  }
}
