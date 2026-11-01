import { HttpClient, HttpResponse } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { fileNameFromDisposition, saveBlob } from '../../shared/lib/download'
import { formatDate } from '../../shared/lib/format'
import { parseDmy } from '../../shared/lib/masks'
import { ActionBarComponent, PageShellComponent, StatComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { CodeTablePanelComponent } from './code-table-panel.component'

const BASE = '/api/maintenance/benefit-member'

interface EntryRow {
  bnfCode: string
  calculation: string
  noOfDays: string
  noOfAnn: string
  rmPerVisit: string
  percentage: string
}
type EntryColumn = Exclude<keyof EntryRow, 'bnfCode'>

/** frmMaintenanceBenefits tab 1 "Benefit Entry": Health / Insured Type, Plan Code, Eff Date and the BNFNew grid (one row per category). */
@Component({
  selector: 'app-benefit-entry-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Benefit Entry">
        <legend>Benefit Entry</legend>
        <app-field label="Health Type" [required]="true" hint="Searchable">
          <select name="health" [(ngModel)]="health">
            <option value=""></option>
            @for (o of healthTypes(); track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Insured Type" [required]="true" hint="Searchable">
          <select name="insured" [(ngModel)]="insured">
            <option value=""></option>
            @for (o of insuredTypes(); track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Plan Code" [required]="true" hint="Searchable">
          <input name="plan" list="bnf-entry-plans" maxlength="10" [(ngModel)]="plan" />
          <datalist id="bnf-entry-plans">
            @for (o of plans(); track o.code) {
              <option [value]="o.code"></option>
            }
          </datalist>
        </app-field>
        <app-field label="Eff Date" [required]="true" hint="Searchable">
          <app-date-input name="effDate" ariaLabel="Eff Date" [(ngModel)]="effDate" />
        </app-field>
        <p class="muted wide">Calculation: 1 - Lower of Limit or % · 2 - Higher of Limit or % · 3 - Ascharge · 4 - BNF Limit</p>
      </fieldset>

      <app-action-bar label="Benefit Entry actions">
        <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="0" />
      </app-action-bar>

      <div class="table-wrap tall">
        <table class="grid" aria-label="Benefit entry">
          <thead>
            <tr>
              <th scope="col">Cat Code</th>
              <th scope="col">Cat Name</th>
              @for (c of columns; track c.key) {
                <th scope="col" class="num">{{ c.label }}</th>
              }
            </tr>
          </thead>
          <tbody>
            @for (cat of categories(); track cat.code) {
              <tr>
                <td>{{ cat.code }}</td>
                <td>{{ cat.name }}</td>
                @for (c of columns; track c.key) {
                  <td class="num">
                    <input class="cell-input num" [attr.inputmode]="c.key === 'rmPerVisit' ? null : 'decimal'" [attr.aria-label]="c.label + ' ' + cat.code"
                      [attr.maxlength]="c.key === 'rmPerVisit' || c.key === 'percentage' ? 15 : 10"
                      [ngModel]="cell(cat.code, c.key)" (change)="setCell(cat.code, c.key, $any($event.target))" />
                  </td>
                }
              </tr>
            } @empty {
              <tr class="empty-row"><td [attr.colspan]="2 + columns.length">No categories.</td></tr>
            }
          </tbody>
        </table>
      </div>
    </div>
  `,
})
export class BenefitEntryPanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly columns: { key: EntryColumn; label: string }[] = [
    { key: 'calculation', label: 'Calculation' },
    { key: 'noOfDays', label: 'No of Visit Per day' },
    { key: 'noOfAnn', label: 'No of Visit Per Annum' },
    { key: 'rmPerVisit', label: 'RM Per Visit' },
    { key: 'percentage', label: 'Percentage' },
  ]
  protected readonly healthTypes = signal<LookupOption[]>([])
  protected readonly insuredTypes = signal<LookupOption[]>([])
  protected readonly plans = signal<LookupOption[]>([])
  protected readonly categories = signal<LookupOption[]>([])
  protected readonly grid = signal<Record<string, Partial<EntryRow>>>({})
  protected readonly busy = signal(false)
  protected health = ''
  protected insured = ''
  protected plan = ''
  protected effDate = ''

  async ngOnInit(): Promise<void> {
    try {
      const s = await firstValueFrom(
        this.http.get<{ healthTypes: LookupOption[]; insuredTypes: LookupOption[]; plans: LookupOption[]; categories: LookupOption[] }>(`${BASE}/entry/setup`),
      )
      this.healthTypes.set(s.healthTypes)
      this.insuredTypes.set(s.insuredTypes)
      this.plans.set(s.plans)
      this.categories.set(s.categories)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected cell(code: string, key: EntryColumn): string {
    return this.grid()[code]?.[key] ?? ''
  }

  /** FGrid_Click: Calculation 1–4, numeric day / annum / percentage (else blank); RM Per Visit is free text. */
  protected setCell(code: string, key: EntryColumn, el: HTMLInputElement): void {
    let v = el.value.trim()
    const n = Number(v)
    if (v !== '' && key !== 'rmPerVisit' && (Number.isNaN(n) || (key === 'calculation' && (n < 1 || n > 4)))) v = ''
    el.value = v
    this.grid.update((g) => ({ ...g, [code]: { ...g[code], [key]: v } }))
  }

  private key() {
    return { healthType: this.health, insuredType: this.insured, planCode: this.plan, effDate: this.effDate }
  }

  /** searchtBNF → tBNFlisting (no message when nothing is saved: the grid stays empty). */
  protected async search(): Promise<void> {
    this.busy.set(true)
    this.grid.set({})
    try {
      const rows = (await firstValueFrom(this.http.get<EntryRow[] | null>(`${BASE}/entry`, { params: this.key() }))) ?? []
      this.grid.set(Object.fromEntries(rows.map((r) => [r.bnfCode, r])))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'warn')
    } finally {
      this.busy.set(false)
    }
  }

  protected add(): void {
    this.clear()
  }

  protected clear(): void {
    this.health = ''
    this.insured = ''
    this.plan = ''
    this.effDate = ''
    this.grid.set({})
  }

  /** cmdSave_Click → SaveBNF: the desktop's checks, then "Do you want to save this record?". */
  protected async save(): Promise<void> {
    const checks: [boolean, string][] = [
      [!this.health, 'Please Enter Health Type !'],
      [!this.insured, 'Please Enter Insured Type!'],
      [!this.plan.trim(), 'Please Enter Plan Code!'],
      [!parseDmy(this.effDate), 'Please Enter (Valid )effectice Date!'],
    ]
    const failed = checks.find(([bad]) => bad)
    if (failed) return this.notify.notify(failed[1], 'error')
    if (!(await this.notify.ask('Do you want to save this record?'))) return
    const rows = this.categories().map((c) => ({ bnfCode: c.code, ...this.grid()[c.code] }))
    this.busy.set(true)
    try {
      const answer = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/entry`, { ...this.key(), rows }))
      this.notify.notify(answer.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }
}

type ListingRow = Record<string, unknown> & { _id: string }
interface ListingSetup {
  healthTypes: LookupOption[]
  insuredTypes: LookupOption[]
  plans: LookupOption[]
  payors: LookupOption[]
  groupCompanies: LookupOption[]
  categories: LookupOption[]
}
const EMPTY_FILTERS = {
  bnfCode: '', healthType: '', insuredType: '', planCode: '', effDate: '', payorCode: '', groupCompany: '', calculation: '', noOfDays: '',
  rmPerVisit: '', noOfAnn: '', percentage: '',
}

/** frmMaintenanceBenefits tab 2 "Benefit Listing": filters, VIEW_CISBNF_PLN_SubCat list, Search / Refresh / Print (Excel) / Clear. */
@Component({
  selector: 'app-benefit-listing-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Benefit Listing">
        <legend>Benefit Listing</legend>
        <app-field label="BNF Code">
          <select name="bnfCode" [(ngModel)]="f.bnfCode">
            <option value=""></option>
            @for (o of setup().categories; track o.code) {
              <option [value]="o.code">{{ o.code }} -- {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Health Type">
          <select name="healthType" [(ngModel)]="f.healthType">
            <option value=""></option>
            @for (o of setup().healthTypes; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Insured Type">
          <select name="insuredType" [(ngModel)]="f.insuredType">
            <option value=""></option>
            @for (o of setup().insuredTypes; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Plan Code">
          <input name="planCode" list="bnf-list-plans" maxlength="10" [(ngModel)]="f.planCode" />
          <datalist id="bnf-list-plans">
            @for (o of setup().plans; track o.code) {
              <option [value]="o.code"></option>
            }
          </datalist>
        </app-field>
        <app-field label="Eff Date">
          <app-date-input name="effDate" ariaLabel="Eff Date" [(ngModel)]="f.effDate" (focusout)="checkDate()" />
        </app-field>
        <app-field label="Payor Code">
          <select name="payorCode" [(ngModel)]="f.payorCode">
            <option value=""></option>
            @for (o of setup().payors; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Group Company">
          <input name="groupCompany" list="bnf-list-groups" maxlength="50" [(ngModel)]="f.groupCompany" />
          <datalist id="bnf-list-groups">
            @for (o of setup().groupCompanies; track o.code) {
              <option [value]="o.code"></option>
            }
          </datalist>
        </app-field>
        <app-field label="BNF Calculation"><input name="calculation" maxlength="3" [(ngModel)]="f.calculation" /></app-field>
        <app-field label="No of Days"><input name="noOfDays" maxlength="3" [(ngModel)]="f.noOfDays" /></app-field>
        <app-field label="No of Ann."><input name="noOfAnn" maxlength="8" [(ngModel)]="f.noOfAnn" /></app-field>
        <app-field label="RM Per Visit"><input name="rmPerVisit" maxlength="8" [(ngModel)]="f.rmPerVisit" /></app-field>
        <app-field label="BNF Percentage"><input name="percentage" maxlength="3" [(ngModel)]="f.percentage" /></app-field>
      </fieldset>

      <app-action-bar label="Benefit Listing actions">
        <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
        <button appButton variant="ghost" (click)="refresh()">Refresh</button>
        <button appButton variant="ghost" [busy]="busy()" (click)="print()">Print</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="rows().length" />
      </app-action-bar>

      <app-data-grid caption="Benefit listing" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
        [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class BenefitListingPanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly setup = signal<ListingSetup>({ healthTypes: [], insuredTypes: [], plans: [], payors: [], groupCompanies: [], categories: [] })
  protected readonly rows = signal<ListingRow[]>([])
  protected readonly selected = signal<string | null>(null)
  protected readonly busy = signal(false)
  protected f = { ...EMPTY_FILTERS }

  protected readonly rowKey = (r: ListingRow) => r._id
  private static col(key: string, label: string, align?: 'right'): GridColumn<ListingRow> {
    return { key, label, align, value: (r) => (key === 'BNFEffDate' ? formatDate(r[key]) : (r[key] as string | number | null)) }
  }
  protected readonly columns: GridColumn<ListingRow>[] = [
    BenefitListingPanelComponent.col('BNFCode', 'BNF Code'),
    BenefitListingPanelComponent.col('BNFCATDescription', 'Description'),
    BenefitListingPanelComponent.col('HLTCode', 'HLT Code'),
    BenefitListingPanelComponent.col('INSCode', 'INS Code'),
    BenefitListingPanelComponent.col('PLNCode', 'PLN Code'),
    BenefitListingPanelComponent.col('BNFCalculation', 'BNF Calculation'),
    BenefitListingPanelComponent.col('BNFNoOfDays', 'No of Days', 'right'),
    BenefitListingPanelComponent.col('BNFNoOfAnn', 'No of Annum', 'right'),
    BenefitListingPanelComponent.col('BNFRMPerVisit', 'RM Per Visit', 'right'),
    BenefitListingPanelComponent.col('BNFPercentage', 'BNF Percentage', 'right'),
    { ...BenefitListingPanelComponent.col('BNFEffDate', 'Eff Date'), sortBy: (r) => String(r['BNFEffDate'] ?? '') },
    BenefitListingPanelComponent.col('PAYCode', 'Payor Code'),
    BenefitListingPanelComponent.col('GRPCompany', 'Group Company'),
  ]

  async ngOnInit(): Promise<void> {
    await this.refresh()
  }

  /** CmdRefresh_Click → BNFCombolisting2. */
  protected async refresh(): Promise<void> {
    try {
      this.setup.set(await firstValueFrom(this.http.get<ListingSetup>(`${BASE}/listing/setup`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** txtEffdate_LostFocus. */
  protected checkDate(): void {
    if (this.f.effDate && !parseDmy(this.f.effDate)) this.notify.notify('Invalid Date Format! ', 'error')
  }

  /** searchBNF → BNFListing. */
  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.selected.set(null)
    try {
      const rows = (await firstValueFrom(this.http.post<Record<string, unknown>[] | null>(`${BASE}/listing/search`, this.f))) ?? []
      this.rows.set(rows.map((r, i) => ({ ...r, _id: String(i) })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'warn')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstBNFlist_click: the row back into the filters. */
  protected select(r: ListingRow): void {
    this.selected.set(r._id)
    const s = (k: string) => (r[k] == null ? '' : String(r[k]).trim())
    this.f = {
      bnfCode: s('BNFCode'), healthType: s('HLTCode'), insuredType: s('INSCode'), planCode: s('PLNCode'), effDate: formatDate(r['BNFEffDate']),
      payorCode: s('PAYCode'), groupCompany: s('GRPCompany'), calculation: s('BNFCalculation'), noOfDays: s('BNFNoOfDays'),
      rmPerVisit: s('BNFRMPerVisit'), noOfAnn: s('BNFNoOfAnn'), percentage: s('BNFPercentage'),
    }
  }

  /** cmdprint_Click (tab 2) → ExportListViewtoExcel of the list. */
  protected async print(): Promise<void> {
    if (this.rows().length === 0) return this.notify.notify('Report cannot be printed without any record.', 'info')
    this.busy.set(true)
    try {
      const res: HttpResponse<Blob> = await firstValueFrom(
        this.http.post(`${BASE}/listing/export`, this.f, { responseType: 'blob', observe: 'response' }),
      )
      saveBlob(res.body!, fileNameFromDisposition(res.headers.get('content-disposition'), 'BenefitListing.xlsx'))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** CmdClear_Click. */
  protected clear(): void {
    this.f = { ...EMPTY_FILTERS }
    this.rows.set([])
    this.selected.set(null)
  }
}

/** Maintenance > Benefit > Member (frmMaintenanceBenefits): Category · Benefit Entry · Benefit Listing. */
@Component({
  selector: 'app-benefit-member-page',
  imports: [PageShellComponent, TabsComponent, TabDirective, CodeTablePanelComponent, BenefitEntryPanelComponent, BenefitListingPanelComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Benefits Maintenance" subtitle="Replaces frmMaintenanceBenefits">
      <app-tabs label="Benefits Maintenance">
        <ng-template appTab="Category"><app-code-table-panel tableKey="benefit-category" /></ng-template>
        <ng-template appTab="Benefit Entry"><app-benefit-entry-panel /></ng-template>
        <ng-template appTab="Benefit Listing"><app-benefit-listing-panel /></ng-template>
      </app-tabs>
    </app-page-shell>
  `,
})
export class BenefitMemberPage {}
