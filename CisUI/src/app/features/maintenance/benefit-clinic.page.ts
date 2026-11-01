import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, input, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { openBlob } from '../../shared/lib/download'
import { parseDmy } from '../../shared/lib/masks'
import { ActionBarComponent, PageShellComponent, StatComponent } from '../../shared/screen.components'

const BASE = '/api/maintenance/benefit-clinic'

/**
 * Maintenance > Benefit > Clinic (frmMaintenanceBenefitsCln) and the never-opened frmMaintenanceBenefitsClnNew (route data legacyForm):
 * BNF Card Type + Eff Date, a grid of every benefit category with its RM Payable Fee, Search / Add / Save / Clear / Refresh / Print.
 */
@Component({
  selector: 'app-benefit-clinic-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Clinic Benefits Maintenance" [subtitle]="'Replaces ' + legacyForm()">
      <div class="screen-form">
        <fieldset aria-label="Clinic Benefits">
          <legend>Clinic Benefits</legend>
          <app-field label="BNF Card Type" [required]="true" hint="Searchable">
            <select name="cardType" [(ngModel)]="cardType">
              <option value=""></option>
              @for (t of cardTypes(); track t) {
                <option [value]="t">{{ t }}</option>
              }
            </select>
          </app-field>
          <app-field label="Eff Date" [required]="true" hint="Searchable">
            <app-date-input name="effDate" ariaLabel="Eff Date" [(ngModel)]="effDate" (focusout)="checkDate()" />
          </app-field>
          <p class="muted wide">Bold - Searchable · * - Mandatory Fields</p>
        </fieldset>

        <app-action-bar label="Clinic Benefits actions">
          <button appButton variant="primary" [disabled]="!searchEnabled()" [busy]="busy()" (click)="search()">Search</button>
          <button appButton (click)="add()">Add</button>
          <button appButton [disabled]="!saveEnabled()" [busy]="busy()" (click)="save()">Save</button>
          <button appButton variant="ghost" (click)="clear()">Clear</button>
          <button appButton variant="ghost" (click)="refresh()">Refresh</button>
          <button appButton variant="ghost" (click)="print()">Print</button>
          <span class="spacer"></span>
          <app-stat label="TOTAL RECORDS" [value]="total()" />
        </app-action-bar>

        <div class="table-wrap tall">
          <table class="grid" aria-label="Clinic benefit fees">
            <thead>
              <tr><th scope="col">Cat Code</th><th scope="col">Cat Name</th><th scope="col" class="num">RM Payable Fee</th></tr>
            </thead>
            <tbody>
              @for (c of categories(); track c.code) {
                <tr>
                  <td>{{ c.code }}</td>
                  <td>{{ c.name }}</td>
                  <td class="num">
                    <input class="cell-input num" inputmode="decimal" maxlength="50" [attr.aria-label]="'RM Payable Fee ' + c.code"
                      [ngModel]="fees()[c.code] ?? ''" (change)="setFee(c.code, $any($event.target))" />
                  </td>
                </tr>
              } @empty {
                <tr class="empty-row"><td colspan="3">No categories.</td></tr>
              }
            </tbody>
          </table>
        </div>
      </div>
    </app-page-shell>
  `,
})
export class BenefitClinicPage implements OnInit {
  readonly legacyForm = input('frmMaintenanceBenefitsCln')

  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly cardTypes = signal<string[]>([])
  protected readonly categories = signal<LookupOption[]>([])
  protected readonly fees = signal<Record<string, string>>({})
  protected readonly total = signal(0)
  protected readonly searchEnabled = signal(true)
  protected readonly saveEnabled = signal(false)
  protected readonly busy = signal(false)
  protected cardType = ''
  protected effDate = ''

  async ngOnInit(): Promise<void> {
    await this.refresh()
  }

  /** Form_Activate / CmdRefresh_Click → ClinicBNFcombolisting: card types and category rows. */
  protected async refresh(): Promise<void> {
    try {
      const setup = await firstValueFrom(this.http.get<{ cardTypes: string[]; categories: LookupOption[] }>(`${BASE}/setup`))
      this.cardTypes.set(setup.cardTypes)
      this.categories.set(setup.categories)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** txtTEffdate_LostFocus: warnings only, as in the desktop. */
  protected checkDate(): void {
    if (!this.effDate) return
    const d = parseDmy(this.effDate)
    if (!d) this.notify.notify('Invalid date format!', 'error')
    else if (d > new Date()) this.notify.notify('Date cannot greater than system date ! ', 'error')
  }

  /** FGrid_Click: only numeric entries are kept. */
  protected setFee(code: string, el: HTMLInputElement): void {
    const v = el.value.trim()
    const ok = v !== '' && !Number.isNaN(Number(v))
    if (!ok) el.value = ''
    this.fees.update((f) => ({ ...f, [code]: ok ? v : '' }))
  }

  /** SearchClinicBNF → ClinicBNFlisting: fills the RM Payable Fee column. */
  protected async search(): Promise<void> {
    this.fees.set({})
    this.total.set(0)
    this.busy.set(true)
    try {
      const params = new HttpParams({ fromObject: { cardType: this.cardType, effDate: this.effDate } })
      const rows = (await firstValueFrom(this.http.get<{ bnfCode: string; rmPerVisit: string | null }[] | null>(BASE, { params }))) ?? []
      this.fees.set(Object.fromEntries(rows.filter((r) => r.rmPerVisit).map((r) => [r.bnfCode, r.rmPerVisit as string])))
      this.total.set(rows.length)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'warn')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdAdd_Click: clear the form and the grid, enable Save. */
  protected async add(): Promise<void> {
    this.clear()
    this.saveEnabled.set(true)
    await this.refresh()
  }

  /** CmdClear_Click. */
  protected clear(): void {
    this.cardType = ''
    this.effDate = ''
    this.fees.set({})
    this.total.set(0)
    this.searchEnabled.set(true)
    this.saveEnabled.set(false)
  }

  /** CmdSave_Click → SaveClinicBNF (the API checks, refuses duplicates and saves one row per category). */
  protected async save(): Promise<void> {
    if (!this.cardType.trim()) return this.notify.notify('Please Enter Plan Code!', 'error')
    if (!parseDmy(this.effDate)) return this.notify.notify('Please Enter (Valid )effectice Date!', 'error')
    if (!(await this.notify.ask('Do you want to save this record?'))) return
    this.busy.set(true)
    try {
      const answer = await firstValueFrom(
        this.http.post<{ message: string }>(BASE, { cardType: this.cardType, effDate: this.effDate, fees: this.fees() }),
      )
      this.notify.notify(answer.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdprint_Click → RptClinicBNF ("CLINIC BENEFIT LISTING"). */
  protected async print(): Promise<void> {
    try {
      openBlob(await firstValueFrom(this.http.get('/api/maintenance/list-reports/clinic-benefit', { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }
}
