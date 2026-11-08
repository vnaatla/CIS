import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'

const BASE = '/api/lg-audit/generate-lg'
type Row = Record<string, unknown>
type Kind = 'Ok' | 'Confirm' | 'Stop'

interface MemberRow {
  member: string
  cover: string
  name: string
  ic: string
  groupCompany: string
  plan: string
  planLgRemarks: string
  planSpecialRemark: string
  clinicalRemarks: string
  category: string
  subsidiary: string
  policyExpiry: string | null
}
interface Limits {
  annualLimit: string
  accumulated: string
  balance: string
  specAnnual: string
  specAccumulated: string
  specBalance: string
  topUp: string
  spTopUp: string
  noAnnLimitDed: string
  noSpecLimitDed: string
  planCode: string
  lgIssuedInd: string
  lgIssuedAfterLimitExInd: string
  usingLimit: string
  usingLimitDesc: string
  hrApproval: boolean
}
interface MemberInfo {
  member: MemberRow
  limits: Limits
  status: string
  validLg: boolean
  messages: string[]
  lgs: Row[]
}
interface LgDetail {
  lgNo: string
  messages: string[]
  member: string
  ic: string
  patient: string
  glType: string
  declineReason: string
  internalRemarks: string
  remarks: string
  limits: Limits
  pharmacy: boolean
  masterCase: boolean
  medRequestInProgress: boolean
  documents: string[]
}
interface Gate {
  kind: Kind
  message: string | null
  notices: string[]
}
interface Action {
  kind: Kind
  message: string | null
  yesNo: boolean
  messages: string[]
  letter: { content: string; contentType: string; fileName: string } | null
}

const s = (v: unknown) => (v == null ? '' : String(v).trim())
const money = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const DELIVERY_PRIVATE = 'Inpatient Delivery Private'
const DELIVERY_COMPLICATIONS = 'Inpatient Delivery Private-Complications'

/** LG & Audit > Generate LG (frmLGGeneration): member limits, the member's LGs, Edit / issue, decline, preview, documents. */
@Component({
  selector: 'app-generate-lg-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent, TabsComponent, TabDirective],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Generate Guarantee Letter" subtitle="Replaces frmLGGeneration">
      <div class="screen-form">
        <fieldset aria-label="Member">
          <legend>Member</legend>
          <app-field label="Membership No"><input name="mbmNo" data-autofocus maxlength="30" [(ngModel)]="mbmNo" (keydown.enter)="search()" /></app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <app-data-grid caption="Members" [columns]="memberColumns" [rows]="members()" [rowKey]="memberKey" [selectable]="true"
          [selectedKey]="info() ? info()!.member.cover : null" (rowSelected)="pickMember($any($event))" />

        @if (info(); as m) {
          <fieldset aria-label="Limits" disabled>
            <legend>{{ m.member.member }}-{{ m.member.cover }} · {{ m.member.groupCompany }} · Plan {{ m.limits.planCode }} · Status {{ m.status }}</legend>
            <app-field label="Annual Limit"><input name="annLmt" [value]="m.limits.annualLimit" /></app-field>
            <app-field label="Accumulated"><input name="accum" [value]="m.limits.accumulated" /></app-field>
            <app-field label="Balance"><input name="balLmt" [value]="m.limits.balance" /></app-field>
            <app-field label="Top Up"><input name="topUp" [value]="m.limits.topUp" /></app-field>
            <app-field label="Spec Annual"><input name="specAnn" [value]="m.limits.specAnnual" /></app-field>
            <app-field label="Spec Accumulated"><input name="specAcc" [value]="m.limits.specAccumulated" /></app-field>
            <app-field label="Spec Balance"><input name="specBal" [value]="m.limits.specBalance" /></app-field>
            <app-field label="SP Top Up"><input name="spTopUp" [value]="m.limits.spTopUp" /></app-field>
            <app-field label="LG Issued"><input name="lgIssued" [value]="m.limits.lgIssuedInd" /></app-field>
            <app-field label="After Limit Exceeded"><input name="afterEx" [value]="m.limits.lgIssuedAfterLimitExInd" /></app-field>
            <app-field label="Using Limit"><input name="usingLimit" [value]="m.limits.usingLimit + ' - ' + m.limits.usingLimitDesc" /></app-field>
            @if (m.member.planSpecialRemark) {
              <app-field label="Special Plan Remark" [wide]="true"><input name="splRemark" class="warn" [value]="m.member.planSpecialRemark" /></app-field>
            }
            @if (m.member.clinicalRemarks) {
              <app-field label="Clinical Remark" [wide]="true"><input name="clinicalRemark" [value]="m.member.clinicalRemarks" /></app-field>
            }
          </fieldset>
        }

        <app-tabs label="Generate LG" [(active)]="tab" (activeChange)="tabChanged($event)">
          <ng-template appTab="LG">
            <app-data-grid caption="LG list" [columns]="lgColumns" [rows]="lgs()" [rowKey]="lgKey" [selectable]="!editing()"
              [selectedKey]="lg()?.lgNo ?? null" (rowSelected)="pickLg($any($event))" />
            <fieldset aria-label="Guarantee letter" [disabled]="!entry()">
              <legend>{{ lg()?.lgNo ?? 'New LG' }}</legend>
              <app-field label="Patient"><input name="patient" [value]="f.patient" disabled /></app-field>
              <app-field label="IC No"><input name="ic" [value]="f.ic" disabled /></app-field>
              <app-field label="Group Clinics"><label class="check-field"><input type="checkbox" name="chkCLN" [(ngModel)]="groupClinics" (ngModelChange)="loadClinics()" /> List the group company's clinics</label></app-field>
              <app-field label="Clinic" [wide]="true" [required]="true">
                <input name="clinic" list="lg-clinics" [(ngModel)]="f.clinic" />
                <datalist id="lg-clinics">@for (c of clinics(); track c.code) { <option [value]="c.name + '     [ ' + c.code + ']'"></option> }</datalist>
              </app-field>
              <app-field label="Type of GL" [required]="true">
                <select name="glType" [(ngModel)]="f.glType"><option value=""></option>@for (t of glTypes(); track t) { <option [value]="t">{{ t }}</option> }</select>
              </app-field>
              <app-field label="Indicator"><select name="indicator" [(ngModel)]="f.indicator"><option value="N">N - New</option><option value="F">F - Follow up</option></select></app-field>
              <app-field label="Letter Head">
                <select name="letterHead" [(ngModel)]="f.letterHead">
                  <option value="MED">MED - MediExpress (Malaysia) Sdn Bhd</option><option value="UHC">UHC - United Health Connect Sdn Bhd</option>
                </select>
              </app-field>
              <app-field label="Doctor" [required]="true"><input name="doctor" maxlength="200" [(ngModel)]="f.doctor" /></app-field>
              <app-field label="Treatment Date" [required]="true"><app-date-input name="treatmentDate" ariaLabel="Treatment Date" [(ngModel)]="f.date" /></app-field>
              <app-field label="GL Amount"><input class="num" name="glAmount" [(ngModel)]="f.glAmount" /></app-field>
              @if (f.glType === deliveryPrivate) {
                <app-field label="Normal"><input class="num" name="maternityNormal" [(ngModel)]="f.normal" /></app-field>
                <app-field label="Caesarean"><input class="num" name="caesarean" [(ngModel)]="f.caesarean" /></app-field>
              }
              @if (f.glType === deliveryComplications) {
                <app-field label="Complications"><input class="num" name="complications" [(ngModel)]="f.complications" /></app-field>
              }
              <app-field label="Diagnosis" [wide]="true" [required]="true"><input name="diagnosis" [(ngModel)]="f.diagnosis" /></app-field>
              <app-field label="Remarks" [wide]="true"><textarea name="remarks" rows="4" [(ngModel)]="f.remarks"></textarea></app-field>
              <app-field label="Internal Remarks" [wide]="true"><input name="internalRemarks" [(ngModel)]="f.internalRemarks" /></app-field>
              <app-field label="Reason for Delay">
                <select name="delayReason" [(ngModel)]="f.delayReason"><option value=""></option>@for (d of delayReasons(); track d) { <option [value]="d">{{ d }}</option> }</select>
              </app-field>
            </fieldset>
            <fieldset aria-label="Decline">
              <legend>Decline</legend>
              <app-field label="Decline Reason">
                <select name="declineReason" [(ngModel)]="declineReason"><option value=""></option>@for (d of declineReasons(); track d) { <option [value]="d">{{ d }}</option> }</select>
              </app-field>
            </fieldset>
            <app-action-bar label="LG actions">
              <button appButton [disabled]="!info() || editing()" (click)="add()">Add</button>
              <button appButton [disabled]="!lg() || editing() || locked()" (click)="edit()">Edit</button>
              <button appButton variant="primary" [busy]="busy()" [disabled]="!entry() || locked()" (click)="issue()">{{ editing() ? 'Update' : 'LG' }}</button>
              <button appButton [disabled]="!lg()" (click)="preview()">Preview</button>
              <button appButton [disabled]="!lg() || locked()" (click)="decline()">Decline</button>
            </app-action-bar>
          </ng-template>
          <ng-template appTab="GP Cases">
            <app-data-grid caption="GP cases" [columns]="caseColumns" [rows]="gpCases()" [rowKey]="caseKey" />
          </ng-template>
          <ng-template appTab="Specialist Cases">
            <app-data-grid caption="Specialist cases" [columns]="caseColumns" [rows]="spCases()" [rowKey]="caseKey" />
          </ng-template>
          <ng-template appTab="Documents">
            <div class="table-wrap">
              <table class="grid" aria-label="Documents">
                <thead><tr><th>Document</th></tr></thead>
                <tbody>
                  @for (d of lg()?.documents ?? []; track d) { <tr><td><button appButton (click)="open(d)">{{ d }}</button></td></tr> }
                  @empty { <tr class="empty-row"><td>No documents.</td></tr> }
                </tbody>
              </table>
            </div>
            <app-field label="File"><input type="file" name="file" (change)="picked = $any($event.target).files?.[0] ?? null" /></app-field>
            <app-action-bar label="Document actions">
              <button appButton [disabled]="!picked || locked() || !!lg()?.masterCase" (click)="copy()">Copy File</button>
            </app-action-bar>
          </ng-template>
        </app-tabs>
      </div>
    </app-page-shell>
  `,
})
export class GenerateLgPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly deliveryPrivate = DELIVERY_PRIVATE
  protected readonly deliveryComplications = DELIVERY_COMPLICATIONS
  protected readonly members = signal<MemberRow[]>([])
  protected readonly info = signal<MemberInfo | null>(null)
  protected readonly lgs = signal<Row[]>([])
  protected readonly lg = signal<LgDetail | null>(null)
  protected readonly glTypes = signal<string[]>([])
  protected readonly declineReasons = signal<string[]>([])
  protected readonly delayReasons = signal<string[]>([])
  protected readonly clinics = signal<{ code: string; name: string }[]>([])
  protected readonly gpCases = signal<Row[]>([])
  protected readonly spCases = signal<Row[]>([])
  protected readonly busy = signal(false)
  protected readonly editing = signal(false)
  protected readonly adding = signal(false)
  protected readonly entry = computed(() => this.editing() || this.adding())
  protected readonly locked = computed(() => !!this.lg()?.medRequestInProgress)
  protected mbmNo = ''
  protected tab = 0
  protected groupClinics = false
  protected declineReason = ''
  protected picked: File | null = null
  protected f = this.emptyForm()

  protected readonly memberKey = (r: MemberRow) => r.cover
  protected readonly lgKey = (r: Row) => s(r['LGNo'])
  protected readonly caseKey = (r: Row) => s(r['CASNumber'])
  protected readonly memberColumns: GridColumn<MemberRow>[] = [
    { key: 'cover', label: 'Member No', value: (r) => `${r.member}-${r.cover}` }, { key: 'name', label: 'Name' }, { key: 'ic', label: 'IC No' },
    { key: 'plan', label: 'Plan' }, { key: 'planLgRemarks', label: 'LG Remarks' },
  ]
  protected readonly lgColumns: GridColumn<Row>[] = [
    { key: 'LGNo', label: 'LG No' }, { key: 'LGDate', label: 'LG Date', value: (r) => formatDate(s(r['LGDate'])) }, { key: 'LGIssuedBy', label: 'Issued By' },
    { key: 'DoctorName', label: 'Doctor' }, { key: 'CLNNameDesc', label: 'Clinic' }, { key: 'GLAmt', label: 'GL Amount', align: 'right', value: (r) => money(r['GLAmt']) },
    { key: 'Diagnosis', label: 'Diagnosis' }, { key: 'GLType', label: 'Type of GL' }, { key: 'CaseProgress', label: 'Progress' }, { key: 'LGTIME', label: 'Time' },
  ]
  protected readonly caseColumns: GridColumn<Row>[] = [
    { key: 'CASNumber', label: 'Case No' }, { key: 'CoverID', label: 'Cov' }, { key: 'DateVisit', label: 'Visit Date', value: (r) => formatDate(s(r['DateVisit'])) },
    { key: 'Claimability', label: 'Claimability' }, { key: 'ReasonDesc', label: 'Diagnosis' }, { key: 'CASStatus', label: 'Status' },
    { key: 'CLMTotalActual', label: 'Bill', align: 'right', value: (r) => money(r['CLMTotalActual'] ?? 0) },
    { key: 'CLMBordxAmt', label: 'Bordx', align: 'right', value: (r) => money(r['CLMBordxAmt'] ?? 0) },
    { key: 'CLMMCOProcFee', label: 'FFS', align: 'right', value: (r) => money(r['CLMMCOProcFee'] ?? 0) },
  ]

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ glTypes: string[]; declineReasons: string[]; delayReasons: string[] }>(`${BASE}/lookups`))
      this.glTypes.set(l.glTypes)
      this.declineReasons.set(l.declineReasons)
      this.delayReasons.set(l.delayReasons)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private emptyForm() {
    return {
      patient: '', ic: '', clinic: '', glType: '', indicator: 'N', letterHead: 'UHC', doctor: '', date: '', glAmount: '', normal: '', caesarean: '',
      complications: '', diagnosis: '', remarks: '', internalRemarks: '', delayReason: '',
    }
  }

  private async show(list: string[]): Promise<void> {
    for (const m of list) this.notify.notify(m, 'info')
  }

  /** cmdSearch_Click. */
  protected async search(): Promise<void> {
    this.busy.set(true)
    this.info.set(null)
    this.lgs.set([])
    this.lg.set(null)
    this.members.set([])
    this.editing.set(false)
    this.adding.set(false)
    try {
      const r = await firstValueFrom(this.http.get<{ rows: MemberRow[]; warranty: string | null }>(`${BASE}/members`, { params: { member: this.mbmNo } }))
      this.members.set(r.rows)
      if (r.warranty) this.notify.notify(r.warranty, 'error')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstMBM_Click. */
  protected async pickMember(r: MemberRow): Promise<void> {
    if (this.editing()) return
    this.busy.set(true)
    try {
      const m = await firstValueFrom(this.http.get<MemberInfo>(`${BASE}/members/${encodeURIComponent(r.member)}/${r.cover}`))
      if (m.limits.hrApproval && !(await this.notify.ask(' All MO/SP GL Request Must be Come From HR.'))) return
      this.info.set(m)
      this.lgs.set(m.lgs)
      this.lg.set(null)
      this.f = { ...this.emptyForm(), patient: r.name, ic: r.ic }
      this.gpCases.set([])
      this.spCases.set([])
      await this.show(m.messages)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstLG_Click. */
  protected async pickLg(r: Row): Promise<void> {
    const m = this.info()
    if (!m || this.editing()) return
    try {
      const params = { groupCompany: m.member.groupCompany, category: m.member.category }
      const d = await firstValueFrom(this.http.get<LgDetail>(`${BASE}/${encodeURIComponent(s(r['LGNo']))}`, { params }))
      this.lg.set(d)
      const amount = s(r['GLAmt']) ? Number(r['GLAmt']).toFixed(2) : ''
      this.f = {
        ...this.emptyForm(), patient: d.patient, ic: d.ic, clinic: `${s(r['CLNNameDesc'])}[${s(r['CLNName'])}]`, glType: s(r['GLType']), doctor: s(r['DoctorName']),
        date: s(r['LGDate']) ? formatDate(s(r['LGDate'])) : formatDate(new Date().toISOString()), glAmount: amount, normal: amount, complications: amount,
        diagnosis: s(r['Diagnosis']), remarks: d.remarks, internalRemarks: d.internalRemarks, delayReason: s(r['CasDelayReason']),
      }
      this.declineReason = d.declineReason
      this.adding.set(false)
      await this.show(d.messages)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** Runs a gate, asking its questions until it stops or passes. */
  private async gate(url: (confirmed: number) => string): Promise<boolean> {
    for (let confirmed = 0; ; confirmed++) {
      const g = await firstValueFrom(this.http.get<Gate>(url(confirmed)))
      await this.show(g.notices)
      if (g.kind === 'Ok') return true
      if (g.kind === 'Stop') {
        if (g.message) this.notify.notify(g.message, 'error')
        return false
      }
      if (!(await this.notify.ask(g.message ?? ''))) return false
    }
  }

  /** CmdAdd_Click. */
  protected async add(): Promise<void> {
    const m = this.info()
    if (!m) return this.notify.notify('Please select a member before you create the LG', 'error')
    try {
      if (await this.gate((c) => `${BASE}/members/${encodeURIComponent(m.member.member)}/${m.member.cover}/add-check?confirmed=${c}`)) {
        this.lg.set(null)
        this.adding.set(true)
        this.f = { ...this.emptyForm(), patient: m.member.name, ic: m.member.ic }
      }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CmdEdit_Click. */
  protected async edit(): Promise<void> {
    const d = this.lg()
    if (!d) return this.notify.notify('Please Select the LG !!', 'error')
    const q = `clinic=${encodeURIComponent(this.f.clinic)}&category=${encodeURIComponent(this.info()?.member.category ?? '')}`
    try {
      if (await this.gate((c) => `${BASE}/${encodeURIComponent(d.lgNo)}/edit-check?${q}&confirmed=${c}`)) {
        this.editing.set(true)
        this.adding.set(false)
      }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private request(answers: boolean[]) {
    const m = this.info()
    return {
      lgNo: this.lg()?.lgNo ?? '', member: m ? `${m.member.member}-${m.member.cover}` : '', patient: this.f.patient, ic: this.f.ic, clinicText: this.f.clinic,
      glType: this.f.glType, doctor: this.f.doctor, treatmentDate: this.f.date, remarks: this.f.remarks, internalRemarks: this.f.internalRemarks,
      diagnosis: this.f.diagnosis, glAmount: this.f.glAmount, maternityNormal: this.f.normal, caesarean: this.f.caesarean, complications: this.f.complications,
      delayReason: this.f.delayReason, groupCompany: m?.member.groupCompany ?? '', category: m?.member.category ?? '', subsidiary: m?.member.subsidiary ?? '',
      edited: this.editing(), answers,
    }
  }

  /** Posts an action, asking its questions until it is done or stops; true when done. */
  private async act(url: string, body: (answers: boolean[]) => object): Promise<boolean> {
    const answers: boolean[] = []
    for (;;) {
      const r = await firstValueFrom(this.http.post<Action>(url, body(answers)))
      if (r.kind === 'Confirm') {
        answers.push(await this.notify.ask(r.message ?? ''))
        continue
      }
      for (const m of r.messages) this.notify.notify(m, 'success')
      if (r.kind === 'Stop') {
        if (r.message) this.notify.notify(r.message, 'error')
        return false
      }
      if (r.letter) saveBlob(new Blob([Uint8Array.from(atob(r.letter.content), (ch) => ch.charCodeAt(0))], { type: r.letter.contentType }), r.letter.fileName)
      return true
    }
  }

  /** cmdLG_Click (LG / Update). */
  protected async issue(): Promise<void> {
    this.busy.set(true)
    try {
      if (await this.act(`${BASE}/issue`, (a) => this.request(a))) {
        this.editing.set(false)
        await this.refreshLgs()
      }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdDecline_Click. */
  protected async decline(): Promise<void> {
    const body = (answers: boolean[]) => ({
      lgNo: this.lg()?.lgNo ?? '', reason: this.declineReason, member: this.request([]).member, patient: this.f.patient, ic: this.f.ic,
      clinicText: this.f.clinic, doctor: this.f.doctor, diagnosis: this.f.diagnosis, treatmentDate: this.f.date, answers,
    })
    try {
      if (await this.act(`${BASE}/decline`, body)) await this.refreshLgs()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CmdPreview_Click. */
  protected async preview(): Promise<void> {
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/preview`, this.request([]), { responseType: 'blob' })), `Preview${this.lg()?.lgNo ?? ''}.pdf`)
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  private async refreshLgs(): Promise<void> {
    const m = this.info()
    if (!m) return
    const fresh = await firstValueFrom(this.http.get<MemberInfo>(`${BASE}/members/${encodeURIComponent(m.member.member)}/${m.member.cover}`))
    this.info.set(fresh)
    this.lgs.set(fresh.lgs)
    this.lg.set(null)
  }

  /** chkCLN_Click. */
  protected async loadClinics(): Promise<void> {
    if (!this.groupClinics) return this.clinics.set([])
    try {
      this.clinics.set(await firstValueFrom(this.http.get<{ code: string; name: string }[]>(`${BASE}/clinics`, { params: { company: this.info()?.member.groupCompany ?? '' } })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** SSTab_Click: the case tabs. */
  protected async tabChanged(t: number): Promise<void> {
    const m = this.info()
    if (!m || (t !== 1 && t !== 2)) return
    try {
      const params = { member: m.member.member, cover: m.member.cover, specialist: String(t === 2) }
      const rows = await firstValueFrom(this.http.get<Row[]>(`${BASE}/cases`, { params }))
      ;(t === 1 ? this.gpCases : this.spCases).set(rows)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async open(name: string): Promise<void> {
    try {
      const url = `${BASE}/${encodeURIComponent(this.lg()!.lgNo)}/documents/file?name=${encodeURIComponent(name)}`
      saveBlob(await firstValueFrom(this.http.get(url, { responseType: 'blob' })), name)
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  /** cmdCopyFile_Click. */
  protected async copy(): Promise<void> {
    const d = this.lg()
    if (!d) return this.notify.notify('Please Select Record!', 'error')
    if (!this.picked) return
    const form = new FormData()
    form.append('file', this.picked)
    try {
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/${encodeURIComponent(d.lgNo)}/documents`, form))
      this.notify.notify(r.message, 'success')
      this.picked = null
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected clear(): void {
    this.mbmNo = ''
    this.members.set([])
    this.info.set(null)
    this.lgs.set([])
    this.lg.set(null)
    this.editing.set(false)
    this.adding.set(false)
    this.f = this.emptyForm()
  }
}
