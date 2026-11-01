import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, input, output, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { parseDmy } from '../../shared/lib/masks'
import { ModalComponent } from '../../shared/modal.component'
import { ActionBarComponent } from '../../shared/screen.components'

const BASE = '/api/membership/registration'

interface Setup {
  races: LookupOption[]
  categories: { category: string; insType: string; planCode: string }[]
  groupCompanies: string[]
}
interface Found {
  mbmNumber: string | null
  active: boolean
  payor: string
  name: string
  ic: string
  race: string
  gender: string
  dob: string | null
  effDate: string | null
  expDate: string | null
  staffCategory: string | null
  department: string
  groupCompany: string
  employeeNo: string
}

type Key = 'ic' | 'name' | 'race' | 'gender' | 'staffCategory' | 'employeeNo' | 'effDate' | 'expDate' | 'dob' | 'groupCompany' | 'department'
const EMPTY = Object.fromEntries(['ic', 'name', 'race', 'gender', 'staffCategory', 'employeeNo', 'effDate', 'expDate', 'dob', 'groupCompany', 'department'].map((k) => [k, ''])) as Record<Key, string>

/** Clinic member registration (frmMemberRegistrationCLN), opened from the clinic invoice with the patient IC; returns the member number. */
@Component({
  selector: 'app-member-registration-modal',
  imports: [FormsModule, ModalComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-modal title="Clinic Member Registration" [wide]="true" (closed)="closed.emit()">
      <div class="screen-form">
        <fieldset aria-label="Member">
          <legend>Member</legend>
          <app-field label="IC No." [required]="true">
            <input name="regIc" data-autofocus maxlength="30" [(ngModel)]="f.ic" />
          </app-field>
          <app-action-bar label="Search member">
            <button appButton (click)="searchMember()">Search MBM</button>
          </app-action-bar>
          <app-field label="Member Name" [required]="true"><input name="regName" maxlength="200" [(ngModel)]="f.name" /></app-field>
          <app-field label="Race" [required]="true">
            <select name="regRace" [(ngModel)]="f.race">
              <option value=""></option>
              @for (r of setup().races; track r.code) {
                <option [value]="r.code">{{ r.code }}</option>
              }
            </select>
          </app-field>
          <app-field label="Gender" [required]="true">
            <select name="regGender" [(ngModel)]="f.gender"><option value=""></option><option value="M">M</option><option value="F">F</option></select>
          </app-field>
          <app-field label="Staff Category" [required]="true">
            <select name="regCategory" [(ngModel)]="f.staffCategory" (ngModelChange)="category()">
              <option value=""></option>
              @for (c of setup().categories; track c.category) {
                <option [value]="c.category">{{ c.category }}</option>
              }
            </select>
          </app-field>
          <app-field label="Employee No" [required]="true"><input name="regEmpNo" maxlength="50" [(ngModel)]="f.employeeNo" /></app-field>
          <app-field label="Birth Date" [required]="true">
            <app-date-input name="regDob" ariaLabel="Birth Date" [(ngModel)]="f.dob" (focusout)="age()" />
          </app-field>
        </fieldset>

        <fieldset aria-label="Policy">
          <legend>Policy</legend>
          <app-field label="Effective Date" [required]="true">
            <app-date-input name="regEff" ariaLabel="Effective Date" [(ngModel)]="f.effDate" (focusout)="effective()" />
          </app-field>
          <app-field label="Expiry Date" [required]="true">
            <app-date-input name="regExp" ariaLabel="Expiry Date" [(ngModel)]="f.expDate" (focusout)="checkDate(f.expDate)" />
          </app-field>
          <app-field label="Group Company" [required]="true">
            <input name="regGroup" list="reg-groups" maxlength="200" [(ngModel)]="f.groupCompany" />
            <datalist id="reg-groups">
              @for (g of setup().groupCompanies; track $index) {
                <option [value]="g"></option>
              }
            </datalist>
          </app-field>
          <app-field label="Department" [required]="true" hint="Type part of the name and press Search">
            <input name="regDept" list="reg-depts" maxlength="200" [(ngModel)]="f.department" />
            <datalist id="reg-depts">
              @for (d of departments(); track $index) {
                <option [value]="d"></option>
              }
            </datalist>
          </app-field>
          <app-action-bar label="Department search">
            <button appButton (click)="searchDepartments()">Search</button>
          </app-action-bar>
          <app-field label="Payor"><input name="regPayor" [value]="payor()" disabled /></app-field>
          <app-field label="Plan / Ins Type"><input name="regPlan" [value]="plan()" disabled /></app-field>
          <app-field label="Age Band"><input name="regAgeBand" [value]="ageBand()" disabled /></app-field>
          <app-field label="Annual Limit"><input class="num" name="regAnnual" [value]="annual()" disabled /></app-field>
          <app-field label="Visit Limit"><input class="num" name="regVisit" [value]="visit()" disabled /></app-field>
        </fieldset>

        @if (mbmNumber(); as no) {
          <p class="notice">Member No: <button type="button" class="link" (click)="registered.emit(no)">{{ no }}</button>
            — * Click Member No to proceed to Invoice Processing</p>
        }

        <app-action-bar label="Registration actions">
          <button appButton variant="primary" [disabled]="!canRegister()" [busy]="busy()" (click)="register()">Register</button>
        </app-action-bar>
      </div>
    </app-modal>
  `,
})
export class MemberRegistrationModalComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  /** The invoice's patient IC. */
  readonly initialIc = input('')
  /** lblMBMNo_DblClick: the member number goes back to the invoice. */
  readonly registered = output<string>()
  readonly closed = output<void>()

  protected readonly setup = signal<Setup>({ races: [], categories: [], groupCompanies: [] })
  protected readonly departments = signal<string[]>([])
  protected readonly payor = signal('')
  protected readonly plan = signal('')
  protected readonly ageBand = signal('')
  protected readonly annual = signal('')
  protected readonly visit = signal('')
  protected readonly mbmNumber = signal<string | null>(null)
  protected readonly canRegister = signal(true)
  protected readonly busy = signal(false)
  protected f: Record<Key, string> = { ...EMPTY }

  async ngOnInit(): Promise<void> {
    this.f.ic = this.initialIc()
    try {
      this.setup.set(await firstValueFrom(this.http.get<Setup>(`${BASE}/setup`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmbSearchMBM_Click → MBMDetails (SP_SearchCISMBM_IC). */
  protected async searchMember(): Promise<void> {
    const ic = this.f.ic
    this.f = { ...EMPTY, ic }
    this.mbmNumber.set(null)
    this.canRegister.set(true)
    try {
      const m = await firstValueFrom(this.http.get<Found>(`${BASE}/member`, { params: new HttpParams({ fromObject: { ic } }) }))
      this.payor.set(m.payor)
      this.f = {
        ...this.f, ic: m.ic || ic, name: m.name, race: m.race, gender: m.gender, dob: formatDate(m.dob), effDate: formatDate(m.effDate),
        expDate: formatDate(m.expDate), staffCategory: m.staffCategory ?? '', department: m.department, groupCompany: m.groupCompany, employeeNo: m.employeeNo,
      }
      if (m.active) {
        this.mbmNumber.set(m.mbmNumber)
        this.canRegister.set(false)
      }
      await this.category()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cboStaffCat_LostFocus: plan and insured type of the category, then the limits. */
  protected async category(): Promise<void> {
    const c = this.setup().categories.find((x) => x.category === this.f.staffCategory)
    this.plan.set(c ? `${c.planCode} / ${c.insType}` : '')
    await this.limits()
  }

  /** txtEffDate_LostFocus: limits, expiry = 31/12 of the year, age band. */
  protected async effective(): Promise<void> {
    if (!this.f.effDate) return
    const eff = parseDmy(this.f.effDate)
    if (!eff) return this.notify.notify('Invalid Date Format! ', 'error')
    this.f.expDate = `31/12/${eff.getFullYear()}`
    await this.limits()
    await this.age()
  }

  protected checkDate(v: string): void {
    if (v && !parseDmy(v)) this.notify.notify('Invalid Date Format! ', 'error')
  }

  /** CLNMemberCallAge. */
  protected async age(): Promise<void> {
    if (!this.f.dob || !this.f.effDate) return
    if (!parseDmy(this.f.dob)) return this.notify.notify('Invalid Date Format! ', 'error')
    try {
      const params = new HttpParams({ fromObject: { effDate: this.f.effDate, dob: this.f.dob } })
      const a = await firstValueFrom(this.http.get<{ age: string; ageBand: string; adultChild: string }>(`${BASE}/age`, { params }))
      this.ageBand.set(`${a.ageBand} (${a.adultChild})`)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** CmdSearch_Click: departments of G001 containing the text. */
  protected async searchDepartments(): Promise<void> {
    try {
      const params = new HttpParams({ fromObject: { text: this.f.department } })
      this.departments.set((await firstValueFrom(this.http.get<string[] | null>(`${BASE}/departments`, { params }))) ?? [])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdReg_Click. */
  protected async register(): Promise<void> {
    this.busy.set(true)
    try {
      const r = await firstValueFrom(this.http.post<{ message: string; mbmNumber: string }>(BASE, this.f))
      this.mbmNumber.set(r.mbmNumber)
      this.canRegister.set(false)
      this.notify.notify(r.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  private async limits(): Promise<void> {
    this.annual.set('')
    this.visit.set('')
    if (!this.f.staffCategory || !parseDmy(this.f.effDate)) return
    try {
      const params = new HttpParams({ fromObject: { staffCategory: this.f.staffCategory, effDate: this.f.effDate } })
      const l = await firstValueFrom(this.http.get<{ annualLimit: number; visitLimit: number; visitFull: boolean }>(`${BASE}/limits`, { params }))
      this.annual.set(formatMoney(l.annualLimit))
      this.visit.set(l.visitFull ? 'FULL' : formatMoney(l.visitLimit))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }
}
