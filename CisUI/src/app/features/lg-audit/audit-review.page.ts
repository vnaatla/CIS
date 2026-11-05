import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { CaseListModalComponent, CaseListRow } from '../invoice/invoice-popups'

const BASE = '/api/lg-audit/audit-review'
type Row = Record<string, unknown>

interface AuditCase {
  case: Row
  diagnoses: { code: string; description: string | null }[]
}

const STATUS: Record<string, string> = { O: 'O - Open', C: 'C - Close', D: 'D - Delete' }
const CLAIMABILITY: Record<string, string> = { A: 'A - Accepted', R: 'R - Repudiated', C: 'C - Cancelled', S: 'S - Special Approve' }
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')

/** LG & Audit > Audit review (frmAuditInvoice): a case read-only, its audit details and its guarantee letter documents. */
@Component({
  selector: 'app-audit-review-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, TabsComponent, TabDirective, CaseListModalComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Audit Review" subtitle="Replaces frmAuditInvoice">
      <div class="screen-form">
        <fieldset aria-label="Case">
          <legend>Case</legend>
          <app-field label="Case No"><input name="caseNo" data-autofocus maxlength="20" [(ngModel)]="caseNo" (keydown.enter)="go()" /></app-field>
          <app-action-bar label="Case actions">
            <button appButton variant="primary" [busy]="busy()" (click)="go()">Go</button>
            <button appButton (click)="step(false)">Prev</button>
            <button appButton (click)="step(true)">Next</button>
            <button appButton (click)="searching.set(true)">Search</button>
            <button appButton (click)="clear()">Clear All</button>
          </app-action-bar>
        </fieldset>
        @if (c(); as x) {
          <app-tabs label="Audit review" [(active)]="tab" (activeChange)="tabChanged($event)">
            <ng-template appTab="Case">
              <fieldset aria-label="Case details" disabled>
                <legend>{{ s(x.case['CASNumber']) }}</legend>
                <app-field label="Status"><input name="status" [value]="status(x.case['CASStatus'])" /></app-field>
                <app-field label="Claimability"><input name="claimability" [value]="claimability(x.case['Claimability'])" /></app-field>
                <app-field label="Pending"><input name="pending" [value]="s(x.case['Pending']) === 'Y' ? 'Y - Yes' : 'N - No'" /></app-field>
                <app-field label="Pay To"><input name="payTo" [value]="s(x.case['INVPayTo']) === 'M' ? 'M - MEMBER' : 'C - CLINIC'" /></app-field>
                <app-field label="Clinic"><input name="clinic" [value]="s(x.case['ClinicCode']) + ' - ' + s(x.case['CLNName'])" /></app-field>
                <app-field label="Inv. No"><input name="invoiceNo" [value]="s(x.case['InvoiceNo'])" /></app-field>
                <app-field label="Inv Date"><input name="invoiceDate" [value]="d(x.case['InvoiceDate'])" /></app-field>
                <app-field label="Rec Date"><input name="receiptDate" [value]="d(x.case['InvoiceReceiptDate'])" /></app-field>
                <app-field label="Mth"><input name="processMonth" [value]="s(x.case['CLMProcMth']) ? s(x.case['CLMProcMth']) + '/' + s(x.case['CLMProcYR']) : ''" /></app-field>
                <app-field label="Member"><input name="member" [value]="s(x.case['MBMNumber']) + '-' + s(x.case['CoverID'])" /></app-field>
                <app-field label="Name"><input name="name" [value]="s(x.case['MBMName'])" /></app-field>
                <app-field label="IC No"><input name="ic" [value]="s(x.case['MBMIcBcPp'])" /></app-field>
                <app-field label="Dept"><input name="dept" [value]="s(x.case['MBMDepartment'])" /></app-field>
                <app-field label="Bill No"><input name="billNo" [value]="s(x.case['BillNo'])" /></app-field>
                <app-field label="Bill Amt"><input class="num" name="billAmt" [value]="money(x.case['INVBillAmt'])" /></app-field>
                <app-field label="Vst Date"><input name="visitDate" [value]="d(x.case['DateVisit'])" /></app-field>
                <app-field label="Vst Time"><input name="visitTime" [value]="s(x.case['TimeVisit']).slice(0, 5)" /></app-field>
                @for (dg of x.diagnoses; track $index) {
                  <app-field [label]="'Diagnosis ' + ($index + 1)"><input [attr.name]="'diag' + $index" [value]="dg.code + ' - ' + (dg.description ?? '')" /></app-field>
                }
                <app-field label="Prelim. Diagnosis" [wide]="true"><input name="priDiag" [value]="s(x.case['PrilimDiag'])" /></app-field>
                @for (i of [1, 2, 3, 4, 5, 6, 7]; track i) {
                  @if (s(x.case['MedInjType' + i])) {
                    <app-field [label]="i + '. Medication'">
                      <input [attr.name]="'med' + i" [value]="s(x.case['MedInjType' + i]) + ' · ' + s(x.case['MedInjDays' + i]) + ' days · ' + s(x.case['MedInjQty' + i]) + ' · ' + money(x.case['MEDCHARGES' + i])" />
                    </app-field>
                  }
                }
              </fieldset>
              <fieldset aria-label="Audit">
                <legend>Audit</legend>
                <app-field label="Query" [wide]="true" [required]="true"><input name="query" [(ngModel)]="audit.query" /></app-field>
                <app-field label="Answer" [wide]="true"><input name="answer" [(ngModel)]="audit.answer" /></app-field>
                <app-field label="Discount Received"><input name="discount" [(ngModel)]="audit.costAvoidance" /></app-field>
                <app-field label="Remarks" [wide]="true"><input name="remarks" [(ngModel)]="audit.remarks" /></app-field>
                <app-field label="Audit By" [required]="true"><input name="auditBy" [(ngModel)]="audit.auditBy" /></app-field>
                <app-field label="Audit Date" [required]="true"><app-date-input name="auditDate" ariaLabel="Audit Date" [(ngModel)]="audit.auditDate" /></app-field>
                <app-action-bar label="Audit actions"><button appButton variant="primary" [busy]="busy()" (click)="update()">Update</button></app-action-bar>
              </fieldset>
            </ng-template>
            <ng-template appTab="Documents">
              <div class="table-wrap">
                <table class="grid" aria-label="Documents">
                  <thead><tr><th>Guarantee Letter</th></tr></thead>
                  <tbody>
                    @for (f of documents(); track f) {
                      <tr><td><button appButton (click)="open(f)">{{ f }}</button></td></tr>
                    } @empty { <tr class="empty-row"><td>No documents.</td></tr> }
                  </tbody>
                </table>
              </div>
              <app-field label="File"><input type="file" name="file" (change)="picked = $any($event.target).files?.[0] ?? null" /></app-field>
              <app-action-bar label="Document actions"><button appButton [disabled]="!picked" (click)="copy()">Copy</button></app-action-bar>
            </ng-template>
          </app-tabs>
        }
      </div>
    </app-page-shell>
    @if (searching()) {
      <app-case-list-modal (picked)="casePicked($event)" (closed)="searching.set(false)" />
    }
  `,
})
export class AuditReviewPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly s = s
  protected readonly d = d
  protected readonly c = signal<AuditCase | null>(null)
  protected readonly documents = signal<string[]>([])
  protected readonly busy = signal(false)
  protected readonly searching = signal(false)
  protected caseNo = ''
  protected tab = 0
  protected picked: File | null = null
  protected audit = { query: '', answer: '', costAvoidance: '', remarks: '', auditBy: '', auditDate: '' }

  protected money(v: unknown): string {
    return v == null || s(v) === '' ? '' : formatMoney(Number(v))
  }

  protected status(v: unknown): string {
    return STATUS[s(v)] ?? '--Select--'
  }

  protected claimability(v: unknown): string {
    return CLAIMABILITY[s(v)] ?? '--Select--'
  }

  /** CmdGo_Click. */
  protected async go(): Promise<void> {
    this.caseNo = this.caseNo.trim().toUpperCase()
    if (!this.caseNo) return this.notify.notify('Not Record Found ! ', 'error')
    this.busy.set(true)
    this.c.set(null)
    this.tab = 0
    try {
      const x = await firstValueFrom(this.http.get<AuditCase>(`${BASE}/${encodeURIComponent(this.caseNo)}`))
      const r = x.case
      this.audit = {
        query: s(r['AuditQuery']), answer: s(r['AuditAnswer']), costAvoidance: s(r['AuditDiscountReceived']), remarks: s(r['AuditRemarks']),
        auditBy: s(r['AuditBy']), auditDate: d(r['AuditDate']),
      }
      this.c.set(x)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected async step(next: boolean): Promise<void> {
    try {
      const r = await firstValueFrom(this.http.get<{ casNumber: string }>(`${BASE}/${encodeURIComponent(this.caseNo.trim() || '0')}/step?next=${next}`))
      this.caseNo = r.casNumber
      await this.go()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected casePicked(r: CaseListRow): void {
    this.searching.set(false)
    this.caseNo = r.casNumber
    void this.go()
  }

  /** CmdUpdate_Click. */
  protected async update(): Promise<void> {
    if (!(await this.notify.ask('Are you sure want update Audit Details? '))) return
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/${encodeURIComponent(this.caseNo)}`, this.audit))
      this.notify.notify(r.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** SSTab1_Click: the Documents tab lists the case's guarantee letters. */
  protected async tabChanged(t: number): Promise<void> {
    if (t !== 1) return
    try {
      this.documents.set(await firstValueFrom(this.http.get<string[]>(`${BASE}/${encodeURIComponent(this.caseNo)}/documents`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async open(name: string): Promise<void> {
    try {
      const url = `${BASE}/${encodeURIComponent(this.caseNo)}/documents/file?name=${encodeURIComponent(name)}`
      saveBlob(await firstValueFrom(this.http.get(url, { responseType: 'blob' })), name)
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  /** CmdCopy_Click. */
  protected async copy(): Promise<void> {
    if (!this.picked) return
    const form = new FormData()
    form.append('file', this.picked)
    try {
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/${encodeURIComponent(this.caseNo)}/documents`, form))
      this.notify.notify(r.message, 'success')
      this.picked = null
      await this.tabChanged(1)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected clear(): void {
    this.caseNo = ''
    this.c.set(null)
    this.documents.set([])
    this.audit = { query: '', answer: '', costAvoidance: '', remarks: '', auditBy: '', auditDate: '' }
  }
}
