import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatDate } from '../../shared/lib/format'
import { parseDmy } from '../../shared/lib/masks'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/claim/pass-to-claim'

interface PassRow {
  passed: string
  casNumber: string
  mbmNumber: string
  patientName: string
  registerDate: string | null
  visitDate: string | null
  passDate: string | null
  noOfDay: string
  opBatchNo: string
  passBy: string
  clinicName: string
  claimStatus: string
}

/** lstCaseTrackList_ItemCheck: why a case cannot be checked. */
function checkMessage(r: PassRow): string | null {
  if (!r.registerDate) return 'Date of Register is Empty.'
  if (!r.visitDate) return 'Date of Admitted is Empty.'
  if (!r.clinicName) return 'Clinic Name is Empty.'
  if (r.claimStatus === '1') return 'File has been send out.'
  return null
}

/** Claim > Pass to Claim (frmPass2Claim): list cases (scan one case number or filter) and pass the checked ones to Claim. */
@Component({
  selector: 'app-pass-to-claim-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Pass to Claim" subtitle="Replaces frmPass2Claim">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <label class="check-field"><input type="checkbox" name="scan" [ngModel]="f.scan" (ngModelChange)="scanChanged($event)" /> Scan</label>
          <app-field [label]="f.scan ? 'Case No.' : 'Case No. From'">
            <input name="caseNo" data-autofocus [(ngModel)]="f.caseNo" (blur)="f.caseNo = f.caseNo.trim().toUpperCase()" (keydown.enter)="search()" />
          </app-field>
          @if (!f.scan) {
            <app-field label="Case No. To"><input name="caseNo2" [(ngModel)]="f.caseNo2" /></app-field>
          }
          <app-field label="Case Status">
            <select name="caseStatus" [(ngModel)]="f.caseStatus">
              <option value=""></option><option value="0 - Open">0 - Open</option><option value="1 - Delete">1 - Delete</option>
              <option value="2 - Open & Close">2 - Open &amp; Close</option>
            </select>
          </app-field>
          <app-field label="Claimability">
            <select name="claimability" [(ngModel)]="f.claimability">
              <option value=""></option><option value="A - Accepted">A - Accepted</option><option value="R - Rejected">R - Rejected</option>
            </select>
          </app-field>
          <app-field label="Pass By">
            <input name="passBy" list="p2c-users" [(ngModel)]="f.passBy" (focus)="loadUsers()" />
            <datalist id="p2c-users">
              @for (u of users(); track u) {
                <option [value]="u"></option>
              }
            </datalist>
          </app-field>
          <app-field label="Status">
            <select name="status" [(ngModel)]="f.status"><option value=""></option><option value="P - Passed">P - Passed</option><option value="N - Not">N - Not</option></select>
          </app-field>
          <app-field label="Date of Register From"><app-date-input name="regFrom" ariaLabel="Date of Register From" [(ngModel)]="f.regFrom" (focusout)="checkDate(f.regFrom)" /></app-field>
          <app-field label="Date of Register To"><app-date-input name="regTo" ariaLabel="Date of Register To" [(ngModel)]="f.regTo" (focusout)="checkDate(f.regTo)" /></app-field>
          <app-field label="Date of Admission From"><app-date-input name="visitFrom" ariaLabel="Date of Admission From" [(ngModel)]="f.visitFrom" (focusout)="checkDate(f.visitFrom)" /></app-field>
          <app-field label="Date of Admission To"><app-date-input name="visitTo" ariaLabel="Date of Admission To" [(ngModel)]="f.visitTo" (focusout)="checkDate(f.visitTo)" /></app-field>
          <app-field label="Date Passed From"><app-date-input name="passFrom" ariaLabel="Date Passed From" [(ngModel)]="f.passFrom" (focusout)="checkDate(f.passFrom)" /></app-field>
          <app-field label="Date Passed To"><app-date-input name="passTo" ariaLabel="Date Passed To" [(ngModel)]="f.passTo" (focusout)="checkDate(f.passTo)" /></app-field>
          <label class="check-field"><input type="checkbox" name="noRegister" [(ngModel)]="f.noRegister" /> Date of Register empty</label>
          <label class="check-field"><input type="checkbox" name="noVisit" [(ngModel)]="f.noVisit" /> Date of Admission empty</label>
          <label class="check-field"><input type="checkbox" name="noPass" [(ngModel)]="f.noPass" /> Date Passed empty</label>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Cases" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="onChecked($event)" />
        <p class="totals" aria-label="Totals">Records: <strong>{{ rows().length }}</strong> · Checked: <strong>{{ checked().size }}</strong></p>
        <app-action-bar label="Pass actions">
          <button appButton (click)="checkAll()">Check</button>
          <button appButton variant="primary" [busy]="busy()" [disabled]="rows().length === 0" (click)="pass()">Pass</button>
        </app-action-bar>
      </div>
    </app-page-shell>
  `,
})
export class PassToClaimPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly users = signal<string[]>([])
  protected readonly rows = signal<PassRow[]>([])
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly busy = signal(false)
  protected readonly key = (r: PassRow) => r.casNumber
  protected f = this.empty()
  private searched = this.empty()

  protected readonly columns: GridColumn<PassRow>[] = [
    { key: 'passed', label: 'Status' }, { key: 'casNumber', label: 'Case No' }, { key: 'mbmNumber', label: 'Member No' },
    { key: 'patientName', label: 'Patient Name' }, { key: 'registerDate', label: 'Date of Register', value: (r) => formatDate(r.registerDate) },
    { key: 'visitDate', label: 'Date of Admission', value: (r) => formatDate(r.visitDate) },
    { key: 'passDate', label: 'Date Passed', value: (r) => formatDate(r.passDate) }, { key: 'noOfDay', label: 'No of Day', align: 'right' },
    { key: 'opBatchNo', label: 'OP Batch No' }, { key: 'passBy', label: 'Pass By' }, { key: 'clinicName', label: 'Clinic' },
    { key: 'claimStatus', label: 'Claim Status' },
  ]

  protected checkDate(v: string): void {
    if (v && !parseDmy(v)) this.notify.notify('Invalid Date Format! ', 'error')
  }

  /** chkPass_Click: the Pass By list (loaded on first use). */
  protected async loadUsers(): Promise<void> {
    if (this.users().length) return
    try {
      this.users.set(await firstValueFrom(this.http.get<string[]>(`${BASE}/users`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** ChkScan_Click: the list and case numbers are cleared. */
  protected scanChanged(scan: boolean): void {
    this.f.scan = scan
    this.f.caseNo = ''
    this.f.caseNo2 = ''
    this.rows.set([])
    this.checked.set(new Set())
  }

  /** cmdSearch_Click: in scan mode a case already listed shows "Data has been uploaded!". */
  protected async search(): Promise<void> {
    this.f.caseNo = this.f.caseNo.trim().toUpperCase()
    if (this.f.scan && !this.f.caseNo) return this.notify.notify('Please Enter Case No', 'error')
    if (this.f.scan && this.rows().some((r) => r.casNumber.toUpperCase() === this.f.caseNo)) return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
    this.rows.set([])
    this.checked.set(new Set())
    try {
      this.searched = { ...this.f }
      this.rows.set(await firstValueFrom(this.http.post<PassRow[]>(`${BASE}/search`, this.f)))
      if (this.f.scan) this.f.caseNo = ''
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected async print(): Promise<void> {
    if (this.rows().length === 0) return this.notify.notify('Report cannot be printed without any record.', 'info')
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/export`, this.searched, { responseType: 'blob' })), 'PassToClaim.xlsx')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  protected onChecked(next: Set<string>): void {
    for (const k of next) {
      if (this.checked().has(k)) continue
      const row = this.rows().find((r) => r.casNumber === k)
      const message = row ? checkMessage(row) : null
      if (message) {
        next.delete(k)
        this.notify.notify(message, 'error')
      }
    }
    this.checked.set(new Set(next))
  }

  /** CmdCheck_Click (SetCheckAllItems): every case that may be checked. */
  protected checkAll(): void {
    this.checked.set(new Set(this.rows().filter((r) => !checkMessage(r)).map(this.key)))
  }

  /** CmdPass_Click. */
  protected async pass(): Promise<void> {
    if (this.rows().length === 0) return this.notify.notify('Please select a payor before proceeding!', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ message: string; rows: PassRow[] }>(BASE, { filters: this.searched, casNumbers: [...this.checked()] }))
      this.notify.notify(r.message, 'success')
      this.rows.set(r.rows)
      this.checked.set(new Set())
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
    this.checked.set(new Set())
  }

  private empty() {
    return {
      scan: true, caseNo: '', caseNo2: '', caseStatus: '2 - Open & Close', claimability: '', passBy: '', status: '', noVisit: false, noRegister: false,
      noPass: false, regFrom: '', regTo: '', visitFrom: '', visitTo: '', passFrom: '', passTo: '',
    }
  }
}
