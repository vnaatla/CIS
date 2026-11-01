import { HttpClient, HttpParams } from '@angular/common/http'
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
import { openBlob } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

const BASE = '/api/maintenance/plan/mco'
type Row = Record<string, unknown> & { _id: string }
type Mode = 'search' | 'selected' | 'add'
const FEES = ['IA', 'IC', 'FA', 'FC', 'SFA', 'SFC', 'GA', 'GC', 'HA', 'HC', 'SHA', 'SHC'] as const
type Fee = (typeof FEES)[number]
const EMPTY_FEES = (): Record<Fee, string> => Object.fromEntries(FEES.map((f) => [f, ''])) as Record<Fee, string>
/** Insured types and their boxes (frmMaintenancePlan MCO frame; F / H also have supplementary fees, Frame12). */
const TYPES: { ins: string; label: string; adult: Fee; child: Fee; suppAdult?: Fee; suppChild?: Fee }[] = [
  { ins: 'I', label: 'Individual', adult: 'IA', child: 'IC' },
  { ins: 'F', label: 'Family', adult: 'FA', child: 'FC', suppAdult: 'SFA', suppChild: 'SFC' },
  { ins: 'G', label: 'Group Individual', adult: 'GA', child: 'GC' },
  { ins: 'H', label: 'Group Family', adult: 'HA', child: 'HC', suppAdult: 'SHA', suppChild: 'SHC' },
]

/** frmMaintenancePlan tab 6 "MCO Fee" (CISDB MCO, one row per insured type). */
@Component({
  selector: 'app-plan-mco-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="MCO Fee">
        <legend>MCO Fee</legend>
        <app-field label="Health Type" [required]="true" hint="Searchable">
          <select name="hltCode" [disabled]="locked()" [(ngModel)]="hltCode">
            <option value=""></option>
            @for (o of setup().healthTypes; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Plan Code" [required]="true" hint="Searchable">
          <input name="plnCode" list="mco-plans" maxlength="4" [disabled]="locked()" [(ngModel)]="plnCode" />
          <datalist id="mco-plans">
            @for (o of setup().plans; track o.code) {
              <option [value]="o.code"></option>
            }
          </datalist>
        </app-field>
        <app-field label="Supp MCO Status">
          <select name="suppStatus" [disabled]="locked()" [(ngModel)]="suppStatus">
            <option value=""></option>
            <option value="A">A - Share Limit for Family</option>
            <option value="B">B - Share Limit for Supplementary</option>
            <option value="C">C - Separate Limit for Supplementary</option>
          </select>
        </app-field>
        <app-field label="MCO Eff. Date" [required]="true" hint="Searchable">
          <app-date-input name="effDate" ariaLabel="MCO Eff. Date" [disabled]="locked()" [(ngModel)]="effDate" />
        </app-field>
        <div class="entry-table">
          <table class="grid" aria-label="MCO fees by insured type">
            <thead>
              <tr><th scope="col">Insured Type</th><th scope="col" class="num">Adult</th><th scope="col" class="num">Child</th>
                <th scope="col" class="num">Supp. Adult</th><th scope="col" class="num">Supp. Child</th></tr>
            </thead>
            <tbody>
              @for (t of types; track t.ins) {
                <tr>
                  <td>{{ t.ins }} - {{ t.label }}</td>
                  <td><input class="num" [attr.name]="t.adult" inputmode="decimal" [disabled]="locked()" [(ngModel)]="fees[t.adult]" /></td>
                  <td><input class="num" [attr.name]="t.child" inputmode="decimal" [disabled]="locked()" [(ngModel)]="fees[t.child]" /></td>
                  <td>
                    @if (t.suppAdult) {
                      <input class="num" [attr.name]="t.suppAdult" inputmode="decimal" [disabled]="locked() || !suppOn()" [(ngModel)]="fees[t.suppAdult]" />
                    }
                  </td>
                  <td>
                    @if (t.suppChild) {
                      <input class="num" [attr.name]="t.suppChild" inputmode="decimal" [disabled]="locked() || !suppOn()" [(ngModel)]="fees[t.suppChild]" />
                    }
                  </td>
                </tr>
              }
            </tbody>
          </table>
        </div>
      </fieldset>

      <app-action-bar label="MCO Fee actions">
        <button appButton variant="primary" [busy]="busy()" [disabled]="mode() === 'add'" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [disabled]="true" title="The desktop has no Edit for MCO fees">Edit</button>
        <button appButton variant="danger" [disabled]="mode() !== 'selected'" (click)="remove()">Delete</button>
        <button appButton [disabled]="mode() !== 'add'" [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <button appButton variant="ghost" (click)="refresh()">Refresh</button>
        <button appButton variant="ghost" (click)="print()">Print</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="rows().length" />
      </app-action-bar>

      <app-data-grid caption="MCO fees" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
        [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class PlanMcoPanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly types = TYPES
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'HLTCode', label: 'Health Type' },
    { key: 'INSCode', label: 'Insured Type' },
    { key: 'PLNCode', label: 'Plan Code' },
    { key: 'SuppMCOStatus', label: 'Supp. Status' },
    ...(['MCOAmtAdult', 'MCOAmtChild', 'MCOSuppAdultAmt', 'MCOSuppChildAmt'] as const).map((k, i) => ({
      key: k, label: ['Adult', 'Child', 'Supp. Adult', 'Supp. Child'][i], align: 'right' as const, value: (r: Row) => formatMoney(r[k]),
    })),
    { key: 'MCOEffDate', label: 'Eff. Date', value: (r) => formatDate(r['MCOEffDate']), sortBy: (r) => String(r['MCOEffDate'] ?? '') },
  ]
  protected readonly setup = signal<{ healthTypes: LookupOption[]; plans: LookupOption[] }>({ healthTypes: [], plans: [] })
  protected readonly rows = signal<Row[]>([])
  protected readonly selected = signal<string | null>(null)
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: Row) => r._id
  protected hltCode = ''
  protected plnCode = ''
  protected suppStatus = ''
  protected effDate = ''
  protected fees = EMPTY_FEES()
  private picked: Row | null = null
  private lastFilters: unknown = null

  async ngOnInit(): Promise<void> {
    await this.refresh()
  }

  protected locked(): boolean {
    return this.mode() === 'selected'
  }

  /** cboSuppMCOStatus_Click: supplementary fees only for B / C (Frame12). */
  protected suppOn(): boolean {
    return this.suppStatus === 'B' || this.suppStatus === 'C'
  }

  protected async refresh(): Promise<void> {
    this.reset()
    this.mode.set('search')
    try {
      this.setup.set(await firstValueFrom(this.http.get<{ healthTypes: LookupOption[]; plans: LookupOption[] }>(`${BASE}/setup`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private body(): unknown {
    return { mco: { hltCode: this.hltCode, plnCode: this.plnCode, suppStatus: this.suppStatus, effDate: this.effDate, fees: this.fees } }
  }

  /** SearchMCO → MCOListing. */
  protected async search(body: unknown = this.body()): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.selected.set(null)
    this.picked = null
    try {
      const rows = (await firstValueFrom(this.http.post<Record<string, unknown>[] | null>(`${BASE}/search`, body))) ?? []
      this.rows.set(rows.map((r, i) => ({ ...r, _id: String(i) })))
      this.lastFilters = body
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
      this.mode.set('search')
    }
  }

  /** lstMCOList_Click: the row's keys and its insured type's fees, read-only. */
  protected select(r: Row): void {
    this.selected.set(r._id)
    this.picked = r
    this.reset(false)
    const s = (k: string) => (r[k] == null ? '' : String(r[k]).trim())
    this.hltCode = s('HLTCode')
    this.plnCode = s('PLNCode')
    this.suppStatus = s('SuppMCOStatus')
    this.effDate = formatDate(r['MCOEffDate'])
    const t = TYPES.find((x) => x.ins === s('INSCode'))
    if (t) {
      this.fees[t.adult] = formatMoney(r['MCOAmtAdult'] ?? 0)
      this.fees[t.child] = formatMoney(r['MCOAmtChild'] ?? 0)
      if (t.suppAdult) this.fees[t.suppAdult] = formatMoney(r['MCOSuppAdultAmt'] ?? 0)
      if (t.suppChild) this.fees[t.suppChild] = formatMoney(r['MCOSuppChildAmt'] ?? 0)
    }
    this.mode.set('selected')
  }

  protected add(): void {
    this.reset()
    this.selected.set(null)
    this.mode.set('add')
  }

  protected clear(): void {
    this.reset()
    this.rows.set([])
    this.selected.set(null)
    this.mode.set('search')
  }

  /** SaveMCO: the per-type refusals, then "Record Saved ! ". */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to save the record?'))) return
    this.busy.set(true)
    try {
      const answer = await firstValueFrom(this.http.post<{ messages: string[] }>(BASE, this.body()))
      for (const m of answer.messages) this.notify.notify(m, m.startsWith('Record Not Save') ? 'error' : 'success')
      this.reset()
      if (this.lastFilters) await this.search(this.lastFilters)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** DeleteMCO (the selected row), then MCOListing. */
  protected async remove(): Promise<void> {
    if (!this.picked) return this.notify.notify('Please search the record', 'error')
    if (!(await this.notify.ask('Do you want to delete the record?'))) return
    try {
      const params = new HttpParams({ fromObject: { healthType: this.hltCode } })
      const answer = await firstValueFrom(this.http.delete<{ message: string }>(`${BASE}/${this.picked['MCOIndex']}`, { params }))
      this.notify.notify(answer.message, 'success')
      this.reset()
      await this.search({ mco: { fees: {} } })
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** PrintMCO → RptMCOList. */
  protected async print(): Promise<void> {
    try {
      openBlob(await firstValueFrom(this.http.get('/api/maintenance/list-reports/mco', { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  private reset(all = true): void {
    if (all) {
      this.hltCode = ''
      this.plnCode = ''
      this.suppStatus = ''
      this.effDate = ''
      this.picked = null
    }
    this.fees = EMPTY_FEES()
  }
}
