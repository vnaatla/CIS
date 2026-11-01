import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { openBlob } from '../../shared/lib/download'
import { digitsOnly, parseDmy } from '../../shared/lib/masks'
import { ActionBarComponent, PageShellComponent, StatComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'

const BASE = '/api/maintenance/entity'

interface ListRow {
  code: string
  name: string | null
}
/** search = Search / Add usable · selected = row shown read-only · add / edit = entering. */
type Mode = 'search' | 'selected' | 'add' | 'edit'

/** Shared list / mode logic of the Clinic Group and Doctor tabs (Search, Add, Edit, Delete, Save, Clear, Refresh, Print). */
abstract class EntityListPanel {
  protected readonly http = inject(HttpClient)
  protected readonly notify = inject(NotifyService)
  protected readonly rows = signal<ListRow[]>([])
  protected readonly selected = signal<string | null>(null)
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: ListRow) => r.code
  protected abstract readonly path: string
  protected abstract readonly report: string

  protected abstract filters(): Record<string, string>
  protected abstract resetForm(): void
  protected abstract show(r: ListRow): Promise<void>
  protected abstract body(): unknown
  protected abstract code(): string

  /** cmdSearch_Click: list, then the boxes are cleared for the next search (as in the desktop). */
  protected async search(filters = this.filters()): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.selected.set(null)
    try {
      const params = new HttpParams({ fromObject: filters })
      this.rows.set((await firstValueFrom(this.http.get<ListRow[] | null>(`${BASE}/${this.path}`, { params }))) ?? [])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
      this.resetForm()
      this.mode.set('search')
    }
  }

  protected async select(r: ListRow): Promise<void> {
    this.selected.set(r.code)
    await this.show(r)
    this.mode.set('selected')
  }

  protected add(): void {
    this.resetForm()
    this.selected.set(null)
    this.mode.set('add')
  }

  protected edit(): void {
    if (this.selected()) this.mode.set('edit')
  }

  /** cmdClear_Click: boxes, list and TOTAL RECORDS. */
  protected clear(): void {
    this.resetForm()
    this.rows.set([])
    this.selected.set(null)
    this.mode.set('search')
  }

  /** cmdSave_Click: "Do you want to save the record?", then the full list again. */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to save the record?'))) return
    const editing = this.mode() === 'edit'
    this.busy.set(true)
    try {
      const url = `${BASE}/${this.path}`
      const answer = editing
        ? await firstValueFrom(this.http.put<{ message: string }>(url, this.body()))
        : await firstValueFrom(this.http.post<{ message: string }>(url, this.body()))
      this.notify.notify(answer.message, 'success')
      await this.search({})
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdDelete_Click. */
  protected async remove(): Promise<void> {
    if (!this.selected()) return this.notify.notify('Please search the record', 'error')
    if (!(await this.notify.ask('Do you want to delete the record?'))) return
    try {
      const answer = await firstValueFrom(this.http.delete<{ message: string }>(`${BASE}/${this.path}/${encodeURIComponent(this.code())}`))
      this.notify.notify(answer.message, 'success')
      await this.search({})
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async print(): Promise<void> {
    try {
      openBlob(await firstValueFrom(this.http.get(`/api/maintenance/list-reports/${this.report}`, { responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }
}

interface Group {
  code: string
  name: string
  initial: string
  add1: string
  add2: string
  add3: string
  contact: string
  admContact: string
  ceo: string
  phone: string
  fax: string
  email: string
  homepage: string
  effDate: string
  tax: string
  conversion: string
  prePost: string
}
const EMPTY_GROUP: Group = {
  code: '', name: '', initial: '', add1: '', add2: '', add3: '', contact: '', admContact: '', ceo: '', phone: '', fax: '', email: '',
  homepage: '', effDate: '', tax: '', conversion: '', prePost: '',
}

const ACTIONS = `
  <app-action-bar [label]="label + ' actions'">
    <button appButton variant="primary" [busy]="busy()" [disabled]="mode() === 'add' || mode() === 'edit'" (click)="search()">Search</button>
    <button appButton (click)="add()">Add</button>
    <button appButton [disabled]="mode() !== 'selected'" (click)="edit()">Edit</button>
    <button appButton variant="danger" [disabled]="mode() !== 'selected'" (click)="remove()">Delete</button>
    <button appButton [disabled]="mode() !== 'add' && mode() !== 'edit'" [busy]="busy()" (click)="save()">Save</button>
    <button appButton variant="ghost" (click)="clear()">Clear</button>
    <button appButton variant="ghost" (click)="clear()">Refresh</button>
    <button appButton variant="ghost" (click)="print()">Print</button>
    <span class="spacer"></span>
    <app-stat label="TOTAL RECORDS" [value]="rows().length" />
  </app-action-bar>
  <app-data-grid [caption]="label" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
    [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
`

/** frmMaintenanceEntity tab 0 "Clinic Group " (GRPClinic). */
@Component({
  selector: 'app-entity-group-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Clinic Group">
        <legend>Clinic Group</legend>
        <app-field label="Clinic Group Code" [required]="true" hint="Searchable">
          <input name="code" maxlength="6" [disabled]="mode() !== 'search'" [(ngModel)]="g.code" />
        </app-field>
        <app-field label="Initial Group Name" [required]="true">
          <input name="initial" maxlength="20" [disabled]="!entering() || mode() === 'edit'" [(ngModel)]="g.initial" />
        </app-field>
        <app-field label="Cln Grp. Name" [required]="true" hint="Searchable">
          <input name="name" maxlength="50" [disabled]="mode() === 'selected'" [(ngModel)]="g.name" />
        </app-field>
        <app-field label="Clinic Group Address" [required]="true" [wide]="true">
          <input name="add1" maxlength="50" [disabled]="!entering()" [(ngModel)]="g.add1" />
          <input name="add2" maxlength="50" [disabled]="!entering()" [(ngModel)]="g.add2" />
          <input name="add3" maxlength="50" [disabled]="!entering()" [(ngModel)]="g.add3" />
        </app-field>
        <app-field label="Contact Person"><input name="contact" maxlength="50" [disabled]="!entering()" [(ngModel)]="g.contact" /></app-field>
        <app-field label="Adm. Contact Person"><input name="admContact" maxlength="50" [disabled]="!entering()" [(ngModel)]="g.admContact" /></app-field>
        <app-field label="CEO Name"><input name="ceo" maxlength="50" [disabled]="!entering()" [(ngModel)]="g.ceo" /></app-field>
        <app-field label="Tel No"><input name="phone" maxlength="15" inputmode="numeric" [disabled]="!entering()" [ngModel]="g.phone" (ngModelChange)="g.phone = digits($event)" /></app-field>
        <app-field label="Fax No"><input name="fax" maxlength="15" inputmode="numeric" [disabled]="!entering()" [ngModel]="g.fax" (ngModelChange)="g.fax = digits($event)" /></app-field>
        <app-field label="Email"><input name="email" maxlength="50" [disabled]="!entering()" [(ngModel)]="g.email" /></app-field>
        <app-field label="Homepage"><input name="homepage" maxlength="50" [disabled]="!entering()" [(ngModel)]="g.homepage" /></app-field>
        <app-field label="Eff. Date" [required]="true">
          <app-date-input name="effDate" ariaLabel="Eff. Date" [disabled]="!entering()" [(ngModel)]="g.effDate" (focusout)="checkDate()" />
        </app-field>
        @for (yn of yesNo; track yn.key) {
          <app-field [label]="yn.label">
            <select [name]="yn.key" [disabled]="!entering()" [(ngModel)]="g[yn.key]">
              <option value=""></option>
              <option value="N">N - No</option>
              <option value="Y">Y - Yes</option>
            </select>
          </app-field>
        }
      </fieldset>
      ${ACTIONS}
    </div>
  `,
})
export class EntityGroupPanelComponent extends EntityListPanel {
  protected readonly label = 'Clinic groups'
  protected readonly path = 'groups'
  protected readonly report = 'group-clinic'
  protected readonly yesNo: { key: 'tax' | 'conversion' | 'prePost'; label: string }[] = [
    { key: 'conversion', label: 'Conversion' },
    { key: 'prePost', label: 'Pre&Post' },
    { key: 'tax', label: 'R&B Tax' },
  ]
  protected readonly columns: GridColumn<ListRow>[] = [
    { key: 'code', label: 'Clinic Group Code' },
    { key: 'name', label: 'Clinic Group Name' },
  ]
  protected g: Group = { ...EMPTY_GROUP }
  protected readonly digits = digitsOnly

  protected entering(): boolean {
    return this.mode() === 'add' || this.mode() === 'edit'
  }

  protected filters(): Record<string, string> {
    return Object.fromEntries(Object.entries({ code: this.g.code.trim(), name: this.g.name.trim() }).filter(([, v]) => v))
  }

  protected resetForm(): void {
    this.g = { ...EMPTY_GROUP }
  }

  protected async show(r: ListRow): Promise<void> {
    try {
      const g = await firstValueFrom(this.http.get<Partial<Record<keyof Group, string | null>>>(`${BASE}/groups/${encodeURIComponent(r.code)}`))
      this.g = Object.fromEntries(Object.keys(EMPTY_GROUP).map((k) => [k, g[k as keyof Group] ?? ''])) as unknown as Group
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected body(): unknown {
    return { group: this.g }
  }

  protected code(): string {
    return this.g.code
  }

  /** txtClnEffDate_LostFocus: warnings only. */
  protected checkDate(): void {
    if (!this.g.effDate) return
    const d = parseDmy(this.g.effDate)
    if (!d) this.notify.notify('Invalid date format!', 'error')
    else if (d > new Date()) this.notify.notify('Date cannot greater than system date ! ', 'error')
  }
}

/** frmMaintenanceEntity tab 1 "Doctor" (Doctor). */
@Component({
  selector: 'app-entity-doctor-panel',
  imports: [FormsModule, FieldComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="Doctor">
        <legend>Doctor</legend>
        <app-field label="Doctor Code" hint="Searchable">
          <input name="code" maxlength="10" [disabled]="mode() !== 'search'" [(ngModel)]="code_" />
        </app-field>
        <app-field label="Doctor Full Name" [required]="true" hint="Searchable" [wide]="true">
          <input name="name" maxlength="50" [disabled]="mode() === 'selected' || mode() === 'edit'" [(ngModel)]="name" />
        </app-field>
      </fieldset>
      ${ACTIONS}
    </div>
  `,
})
export class EntityDoctorPanelComponent extends EntityListPanel {
  protected readonly label = 'Doctors'
  protected readonly path = 'doctors'
  protected readonly report = 'doctor'
  protected readonly columns: GridColumn<ListRow>[] = [
    { key: 'code', label: 'Doctor Code' },
    { key: 'name', label: 'Doctor Name' },
  ]
  protected code_ = ''
  protected name = ''

  protected filters(): Record<string, string> {
    return Object.fromEntries(Object.entries({ code: this.code_.trim(), name: this.name.trim() }).filter(([, v]) => v))
  }

  protected resetForm(): void {
    this.code_ = ''
    this.name = ''
  }

  protected async show(r: ListRow): Promise<void> {
    this.code_ = r.code
    this.name = r.name ?? ''
  }

  protected body(): unknown {
    return { code: this.code_, name: this.name }
  }

  protected code(): string {
    return this.code_
  }
}

/** Maintenance > Entity (frmMaintenanceEntity): the desktop's five tabs. */
@Component({
  selector: 'app-entity-page',
  imports: [PageShellComponent, TabsComponent, TabDirective, EntityGroupPanelComponent, EntityDoctorPanelComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Entity Maintenance" subtitle="Replaces frmMaintenanceEntity">
      <app-tabs label="Entity Maintenance">
        <ng-template appTab="Clinic Group"><app-entity-group-panel /></ng-template>
        <ng-template appTab="Doctor"><app-entity-doctor-panel /></ng-template>
        <ng-template appTab="Clinic Info"><p class="muted">Clinic Info: not migrated yet.</p></ng-template>
        <ng-template appTab="Group Company Clinic"><p class="muted">Group Company Clinic: not migrated yet.</p></ng-template>
        <ng-template appTab="Clinic Access"><p class="muted">Clinic Access: not migrated yet.</p></ng-template>
      </app-tabs>
    </app-page-shell>
  `,
})
export class EntityPage {}
