import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
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
import { formatDate } from '../../shared/lib/format'
import { parseDmy } from '../../shared/lib/masks'
import { ModalComponent } from '../../shared/modal.component'
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

const BASE = '/api/maintenance/entity'

interface Link {
  grpComCode: string
  grpCoName: string | null
  grpComClnCode: string
  clnName: string | null
  clnAddress: string | null
  clnPostCode: string | null
  clnState: string | null
  clnCity: string | null
  clnLocation: string | null
  cln24Hrs: string | null
  clnPhone: string | null
  clnFax: string | null
}
type Mode = 'search' | 'selected' | 'add' | 'edit'

/** frmMaintenanceEntity tab 3 "Group Company Clinic" (CISDB GRPCOMClinic / VIEW_GrpCom_Clinic). */
@Component({
  selector: 'app-entity-grpcom-panel',
  imports: [FormsModule, FieldComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Group Company Clinic">
        <legend>Group Company Clinic</legend>
        <app-field label="Group Company Name" [required]="true" hint="Searchable">
          <input name="groupCompany" list="grpcom-companies" [disabled]="mode() === 'selected'" [(ngModel)]="groupCompany" />
          <datalist id="grpcom-companies">
            @for (o of lookups().groupCompanies; track o.code) {
              <option [value]="o.code + ' - ' + o.name"></option>
            }
          </datalist>
        </app-field>
        <app-field label="Clinic Code" [required]="true" hint="Searchable">
          <input name="clinic" list="grpcom-clinics" [disabled]="mode() === 'selected'" [(ngModel)]="clinic" />
          <datalist id="grpcom-clinics">
            @for (o of lookups().clinics; track o.code) {
              <option [value]="o.code + ' - ' + o.name"></option>
            }
          </datalist>
        </app-field>
        <label class="form-field">
          <span>Load the group company and clinic lists</span>
          <input type="checkbox" name="loadLists" [ngModel]="listsLoaded()" (ngModelChange)="toggleLists($event)" />
        </label>
      </fieldset>

      <app-action-bar label="Group Company Clinic actions">
        <button appButton variant="primary" [busy]="busy()" [disabled]="mode() === 'add' || mode() === 'edit'" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [disabled]="mode() !== 'selected'" (click)="mode.set('edit')">Edit</button>
        <button appButton variant="danger" [disabled]="mode() !== 'selected'" (click)="remove()">Delete</button>
        <button appButton [disabled]="mode() !== 'add' && mode() !== 'edit'" [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <button appButton variant="ghost" (click)="refresh()">Refresh</button>
        <button appButton variant="ghost" (click)="print()">Print</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="total()" />
      </app-action-bar>

      <app-data-grid caption="Group company clinics" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true" [serverTotal]="total()" [serverPage]="pageNo()" (pageRequest)="loadPage($event)"
        [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class EntityGrpComPanelComponent {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly columns: GridColumn<Link>[] = [
    { key: 'grpComCode', label: 'Group Company Code' },
    { key: 'grpCoName', label: 'Group Company Name' },
    { key: 'grpComClnCode', label: 'Clinic Code' },
    { key: 'clnName', label: 'Clinic Name' },
    { key: 'clnAddress', label: 'Address' },
    { key: 'clnPostCode', label: 'Post Code' },
    { key: 'clnState', label: 'State' },
    { key: 'clnCity', label: 'City' },
    { key: 'clnLocation', label: 'Location' },
    { key: 'cln24Hrs', label: '24 Hours' },
    { key: 'clnPhone', label: 'Tel No' },
    { key: 'clnFax', label: 'Fax No' },
  ]
  protected readonly lookups = signal<{ groupCompanies: LookupOption[]; clinics: LookupOption[] }>({ groupCompanies: [], clinics: [] })
  protected readonly listsLoaded = signal(false)
  protected readonly rows = signal<Link[]>([])
  protected readonly total = signal(0)
  protected readonly selected = signal<string | null>(null)
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: Link) => `${r.grpComCode}|${r.grpComClnCode}`
  protected groupCompany = ''
  protected clinic = ''
  private picked: Link | null = null

  /** Check1_Click: load the combos, or empty them. */
  protected async toggleLists(on: boolean): Promise<void> {
    this.listsLoaded.set(on)
    if (!on) return this.lookups.set({ groupCompanies: [], clinics: [] })
    try {
      this.lookups.set(await firstValueFrom(this.http.get<{ groupCompanies: LookupOption[]; clinics: LookupOption[] }>(`${BASE}/group-company-clinics/lookups`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** SearchGrpComCLN → GRPComCLNListing. */
  protected readonly pageNo = signal(0)
  private query = { groupCompany: '', clinic: '', pageSize: 50, sortKey: '', sortDesc: false }

  /** Grid page / sort / size change (server paging). */
  protected async loadPage(r: { page: number; pageSize: number; sortKey: string | null; sortDir: 'asc' | 'desc' }): Promise<void> {
    this.query = { ...this.query, pageSize: r.pageSize, sortKey: r.sortKey ?? '', sortDesc: r.sortDir === 'desc' }
    await this.search(this.query.groupCompany, this.query.clinic, r.page)
  }

  protected async search(groupCompany = this.groupCompany, clinic = this.clinic, page = 0): Promise<void> {
    this.busy.set(true)
    if (page === 0) this.rows.set([])
    this.selected.set(null)
    this.picked = null
    this.query = { ...this.query, groupCompany, clinic }
    try {
      const q = this.query
      const params = new HttpParams({ fromObject: { groupCompany, clinic, page, pageSize: q.pageSize, sortKey: q.sortKey, sortDesc: q.sortDesc } })
      const result = await firstValueFrom(this.http.get<{ rows: Link[]; total: number } | null>(`${BASE}/group-company-clinics`, { params }))
      this.rows.set(result?.rows ?? [])
      this.total.set(result?.total ?? 0)
      this.pageNo.set(page)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
      this.mode.set('search')
    }
  }

  /** lstGrpComCLN_Click: the codes back into the boxes, read-only. */
  protected select(r: Link): void {
    this.selected.set(this.rowKey(r))
    this.picked = r
    this.groupCompany = r.grpComCode
    this.clinic = r.grpComClnCode
    this.mode.set('selected')
  }

  protected add(): void {
    this.groupCompany = ''
    this.clinic = ''
    this.rows.set([])
    this.total.set(0)
    this.selected.set(null)
    this.picked = null
    this.mode.set('add')
  }

  protected clear(): void {
    this.groupCompany = ''
    this.clinic = ''
    this.rows.set([])
    this.total.set(0)
    this.selected.set(null)
    this.picked = null
    this.mode.set('search')
  }

  /** CmdRefresh_Click: boxes and list cleared, combos reloaded when the lists are on. */
  protected async refresh(): Promise<void> {
    this.clear()
    if (this.listsLoaded()) await this.toggleLists(true)
  }

  /** SaveGrpComClinic, then the whole list again (capped). */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to save the record?'))) return
    const editing = this.mode() === 'edit'
    const body = { groupCompany: this.groupCompany, clinic: this.clinic, oldGroupCompany: this.picked?.grpComCode, oldClinic: this.picked?.grpComClnCode }
    this.busy.set(true)
    try {
      const url = `${BASE}/group-company-clinics`
      const answer = editing
        ? await firstValueFrom(this.http.put<{ message: string }>(url, body))
        : await firstValueFrom(this.http.post<{ message: string }>(url, body))
      this.notify.notify(answer.message, 'success')
      this.groupCompany = ''
      this.clinic = ''
      await this.search('', '')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** DeleteGrpComClinic, then the whole list again (capped). */
  protected async remove(): Promise<void> {
    if (!this.picked) return this.notify.notify('Please search the record', 'error')
    if (!(await this.notify.ask('Do you want to delete the record?'))) return
    try {
      const params = new HttpParams({ fromObject: { groupCompany: this.picked.grpComCode, clinic: this.picked.grpComClnCode } })
      const answer = await firstValueFrom(this.http.delete<{ message: string }>(`${BASE}/group-company-clinics`, { params }))
      this.notify.notify(answer.message, 'success')
      this.groupCompany = ''
      this.clinic = ''
      await this.search('', '')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** PrintGrpComCln → RptGrpComCLN. */
  protected async print(): Promise<void> {
    try {
      openBlob(await firstValueFrom(this.http.get('/api/maintenance/list-reports/group-company-clinic', { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }
}

interface Access {
  clinicCode: string
  clnName: string | null
  effDateFrom: string | null
  effDateTo: string | null
  createdBy: string | null
  createdDate: string | null
}
interface Found {
  code: string
  name: string | null
  location: string
  docName: string
}

/** frmMaintenanceEntity tab 4 "Clinic Access" (HISDB CLNAccessControl) with the frmSearchCLN look-up. */
@Component({
  selector: 'app-entity-access-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent, ModalComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Clinic Access">
        <legend>Clinic Access</legend>
        <app-field label="Clinic Code" [required]="true">
          <span style="display: flex; gap: 0.5rem">
            <input name="clinicCode" maxlength="10" [disabled]="!entering()" [(ngModel)]="code" />
            <button appButton variant="ghost" [disabled]="!entering()" (click)="lookup.set(true)">...</button>
          </span>
        </app-field>
        <app-field label="Clinic Name"><output name="clinicName">{{ clinicName() }}</output></app-field>
        <app-field label="Effective Date From" [required]="true">
          <app-date-input name="from" ariaLabel="Effective Date From" [disabled]="!entering()" [(ngModel)]="from" (focusout)="check(from, 'From')" />
        </app-field>
        <app-field label="Effective Date To" [required]="true">
          <app-date-input name="to" ariaLabel="Effective Date To" [disabled]="!entering()" [(ngModel)]="to" (focusout)="check(to, 'To')" />
        </app-field>
      </fieldset>

      <app-action-bar label="Clinic Access actions">
        <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [disabled]="!entering()" [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <button appButton variant="ghost" (click)="refresh()">Refresh</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="rows().length" />
      </app-action-bar>

      <app-data-grid caption="Clinic access (latest 100)" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [tall]="true" />
    </div>

    @if (lookup()) {
      <app-modal title="Search Clinic" [wide]="true" (closed)="lookup.set(false)">
        <div class="screen-form">
          <app-field label="Clinic Name" [required]="true">
            <input name="lookupName" data-autofocus [(ngModel)]="lookupName" (keydown.enter)="find()" />
          </app-field>
          <app-action-bar label="Search clinic actions">
            <button appButton variant="primary" [busy]="busy()" (click)="find()">Search</button>
          </app-action-bar>
          <app-data-grid caption="Clinics found" [columns]="foundColumns" [rows]="found()" [rowKey]="foundKey" [selectable]="true" (rowSelected)="pick($event)" />
          <p class="muted">Click a clinic to choose it.</p>
        </div>
      </app-modal>
    }
  `,
})
export class EntityAccessPanelComponent {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly columns: GridColumn<Access>[] = [
    { key: 'clinicCode', label: 'Clinic Code' },
    { key: 'clnName', label: 'Clinic Name' },
    { key: 'effDateFrom', label: 'Eff Date From', value: (r) => formatDate(r.effDateFrom), sortBy: (r) => r.effDateFrom ?? '' },
    { key: 'effDateTo', label: 'Eff Date To', value: (r) => formatDate(r.effDateTo), sortBy: (r) => r.effDateTo ?? '' },
    { key: 'createdBy', label: 'Created By' },
    { key: 'createdDate', label: 'Created Date', value: (r) => formatDate(r.createdDate), sortBy: (r) => r.createdDate ?? '' },
  ]
  protected readonly foundColumns: GridColumn<Found>[] = [
    { key: 'code', label: 'Clinic Code' },
    { key: 'name', label: 'Clinic Name' },
    { key: 'location', label: 'Location' },
    { key: 'docName', label: 'Doctor Name' },
  ]
  protected readonly rows = signal<Access[]>([])
  protected readonly found = signal<Found[]>([])
  protected readonly clinicName = signal('')
  protected readonly entering = signal(false)
  protected readonly lookup = signal(false)
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: Access) => `${r.clinicCode}|${r.createdDate}`
  protected readonly foundKey = (r: Found) => r.code
  protected code = ''
  protected from = ''
  protected to = ''
  protected lookupName = ''

  /** txtEffDateFrom_LostFocus / txtEffDateTo_LostFocus. */
  protected check(value: string, which: 'From' | 'To'): void {
    if (value && !parseDmy(value)) this.notify.notify(`Invalid Effective Date ${which} format!`, 'error')
  }

  /** SearchClinicAccess → ClinicAccessListing. */
  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    try {
      this.rows.set((await firstValueFrom(this.http.get<Access[] | null>(`${BASE}/clinic-access`))) ?? [])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  protected add(): void {
    this.reset()
    this.entering.set(true)
  }

  protected clear(): void {
    this.reset()
    this.rows.set([])
    this.entering.set(false)
  }

  /** CmdRefresh_Click: ClearForm + ClinicAccessListing. */
  protected async refresh(): Promise<void> {
    this.reset()
    this.entering.set(false)
    await this.search()
  }

  /** SaveClnAccess, then the list again. */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to save the record?'))) return
    this.busy.set(true)
    try {
      const answer = await firstValueFrom(
        this.http.post<{ message: string }>(`${BASE}/clinic-access`, { clinicCode: this.code, effDateFrom: this.from, effDateTo: this.to }),
      )
      await this.search()
      this.notify.notify(answer.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** frmSearchCLN.cmdSearch_Click. */
  protected async find(): Promise<void> {
    this.found.set([])
    this.busy.set(true)
    try {
      const params = new HttpParams({ fromObject: { name: this.lookupName } })
      this.found.set((await firstValueFrom(this.http.get<Found[] | null>(`${BASE}/clinic-search`, { params }))) ?? [])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstClinic_DblClick: code and name back into the tab. */
  protected pick(r: Found): void {
    this.code = r.code
    this.clinicName.set((r.name ?? '').trim())
    this.lookup.set(false)
  }

  private reset(): void {
    this.code = ''
    this.from = ''
    this.to = ''
    this.clinicName.set('')
  }
}
