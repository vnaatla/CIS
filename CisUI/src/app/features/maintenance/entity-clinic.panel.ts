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
import { digitsOnly, parseDmy } from '../../shared/lib/masks'
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

const BASE = '/api/maintenance/entity/clinics'

type Key =
  | 'code' | 'name' | 'initial' | 'groupCode' | 'type' | 'bnfType' | 'add1' | 'add2' | 'add3' | 'postCode' | 'city' | 'state' | 'receivedDate'
  | 'panel' | 'xRay' | 'formemaRef' | 'noOfDoc' | 'phone' | 'fax' | 'email' | 'homepage' | 'contact' | 'docName' | 'bankAcc' | 'bank' | 'branch'
  | 'accPerson' | 'hrs24' | 'remark' | 'ffs' | 'ffsPerc' | 'location' | 'operatingHr' | 'payName' | 'payAdd1' | 'payAdd2' | 'payAdd3'
  | 'online' | 'glFormat' | 'regNo'
type Clinic = Record<Key, string>
type Mode = 'search' | 'selected' | 'add' | 'edit'
type Row = Record<string, unknown> & { _id: string }

interface FieldDef {
  key: Key
  label: string
  /** text (default) · digits · date · select (options) · list (datalist of setup values) · area. */
  kind?: 'digits' | 'date' | 'select' | 'list' | 'area'
  max?: number
  required?: boolean
  searchable?: boolean
  options?: [string, string][]
  /** CLN column the row value comes from. */
  column: string
  wide?: boolean
}

const YN: [string, string][] = [['Y', 'Y - Yes'], ['N', 'N - No']]
/** Fields in the form's order: captions, lengths, mandatory marks (*) and searchable (bold) from frmMaintenanceEntity tab 2. */
const FIELDS: FieldDef[] = [
  { key: 'code', label: 'Clinic Code', max: 5, searchable: true, column: 'CLNCode' },
  { key: 'name', label: 'Clinic Name', max: 50, required: true, searchable: true, column: 'CLNName' },
  { key: 'initial', label: 'Clinic Initial', max: 50, required: true, searchable: true, column: 'CLNInitial' },
  { key: 'groupCode', label: 'Clinic Group', kind: 'select', searchable: true, column: 'ClinicGrpCode' },
  { key: 'type', label: 'Clinic Type', kind: 'select', required: true, searchable: true, options: [['G', 'G - General Practisional'], ['S', 'S - Specialised']], column: 'CLNType' },
  { key: 'bnfType', label: 'Benefit Type', kind: 'select', required: true, options: [['1', '1 - Standard'], ['2', '2 - Clinic'], ['3', '3 - Plan']], column: 'CLNBNFType' },
  { key: 'add1', label: 'Address', max: 200, required: true, column: 'CLNAdd1', wide: true },
  { key: 'add2', label: 'Address 2', max: 200, column: 'CLNAdd2', wide: true },
  { key: 'add3', label: 'Address 3', max: 200, column: 'CLNAdd3', wide: true },
  { key: 'postCode', label: 'Post Code', kind: 'digits', max: 5, required: true, searchable: true, column: 'CLNPostCode' },
  { key: 'city', label: 'City', kind: 'list', required: true, searchable: true, column: 'CLNCity' },
  { key: 'state', label: 'State', kind: 'list', required: true, searchable: true, column: 'CLNState' },
  { key: 'receivedDate', label: 'Received Date', kind: 'date', required: true, column: 'CLNReceivedDate' },
  { key: 'panel', label: 'Panel', kind: 'select', required: true, searchable: true, options: [...YN, ['D', 'D - Declined'], ['C', 'C - Confirmed']], column: 'CLNPanel' },
  { key: 'xRay', label: 'X-Ray', kind: 'select', options: YN, column: 'CLNXRay' },
  { key: 'formemaRef', label: 'Formema Ref', max: 15, column: 'CLNForRef' },
  { key: 'noOfDoc', label: 'No Of Doctor', kind: 'digits', max: 3, column: 'CLNNoofDoc' },
  { key: 'docName', label: 'Doctor Name', max: 50, column: 'CLNDocName' },
  { key: 'contact', label: 'Contact Person', max: 50, column: 'CLNRepresentative' },
  { key: 'phone', label: 'Tel. No', kind: 'digits', max: 12, column: 'CLNPhone' },
  { key: 'fax', label: 'Fax No', kind: 'digits', max: 12, column: 'CLNFax' },
  { key: 'email', label: 'Email', max: 1000, column: 'CLNEmail' },
  { key: 'homepage', label: 'Homepage', max: 50, column: 'CLNHomepage' },
  { key: 'hrs24', label: '24 Hours', kind: 'select', required: true, searchable: true, options: YN, column: 'CLN24Hrs' },
  { key: 'operatingHr', label: 'Operating Hour', max: 100, required: true, column: 'CLNOperatingHr' },
  { key: 'location', label: 'Location', max: 100, required: true, column: 'CLNLocation' },
  { key: 'ffs', label: 'FFS Status', kind: 'select', required: true, searchable: true, options: YN, column: 'CLNFFSStatus' },
  { key: 'ffsPerc', label: 'FFS Percentage', max: 50, column: 'CLNMCOProcfeePerc' },
  { key: 'online', label: 'Online Status', kind: 'select', searchable: true, options: YN, column: 'CLNOnlineStatus' },
  { key: 'glFormat', label: 'GL Format', kind: 'select', searchable: true, options: [['M', 'M - Medix'], ['H', 'H - Health Connect']], column: 'CLNGLFormat' },
  { key: 'regNo', label: 'Registration No', max: 50, column: 'CLNRegNo' },
  { key: 'bankAcc', label: 'Bank Account', max: 25, column: 'CLNBankAcc' },
  { key: 'bank', label: 'Bank', max: 50, column: 'CLNBankName' },
  { key: 'branch', label: 'Branch', max: 50, column: 'CLNBankBrn' },
  { key: 'accPerson', label: 'Account Person', max: 50, column: 'CLNAccPerson' },
  { key: 'payName', label: 'Pay. Clinic Name', max: 200, column: 'CLNPName', wide: true },
  { key: 'payAdd1', label: 'Pay. Add1', max: 50, column: 'CLNPAdd1', wide: true },
  { key: 'payAdd2', label: 'Pay. Add2', max: 50, column: 'CLNPAdd2', wide: true },
  { key: 'payAdd3', label: 'Pay. Add3', max: 50, column: 'CLNPAdd3', wide: true },
  { key: 'remark', label: 'Remarks', kind: 'area', max: 1500, column: 'CLNRemark', wide: true },
]
const EMPTY = Object.fromEntries(FIELDS.map((f) => [f.key, ''])) as Clinic
/** The ListView's 38 columns. */
const LIST: [string, string][] = [
  ['CLNCode', 'Clinic Code'], ['CLNName', 'Clinic Name'], ['ClinicGrpCode', 'Clinic Group Code'], ['CLNInitial', 'Clinic Initial'], ['CLNType', 'Clinic Type'],
  ['CLNBNFType', 'Benefit Type'], ['CLNAdd1', 'Address 1'], ['CLNAdd2', 'Address 2'], ['CLNAdd3', 'Address 3'], ['CLNPostCode', 'Post Code'], ['CLNCity', 'City'],
  ['CLNState', 'State'], ['CLNReceivedDate', 'Received Date'], ['CLNPanel', 'Panel'], ['CLNXRay', 'X-Ray'], ['CLNForRef', 'Formema Ref'],
  ['CLNNoofDoc', 'No Of Doctor'], ['CLNPhone', 'Tel. No'], ['CLNFax', 'Fax No'], ['CLNEmail', 'Email'], ['CLNHomepage', 'Homepage'], ['CLNBankAcc', 'Bank Account'],
  ['CLNBankName', 'Bank'], ['CLNBankBrn', 'Branch'], ['CLNAccPerson', 'Account Person'], ['CLN24Hrs', '24 Hours'], ['CLNRemark', 'Remarks'],
  ['CLNFFSStatus', 'FFS Status'], ['CLNOperatingHr', 'Operating Hour'], ['CLNPName', 'Pay. Clinic Name'], ['CLNPAdd1', 'Pay. Add1'], ['CLNPAdd2', 'Pay. Add2'],
  ['CLNPAdd3', 'Pay. Add3'], ['CLNPPostCode', 'Pay. Post Code'], ['CLNPCity', 'Pay. City'], ['CLNPState', 'Pay. State'], ['CLNOnlineStatus', 'Online Status'],
  ['CLNRegNo', 'Registration No'],
]

/** frmMaintenanceEntity tab 2 "Clinic Info" (HISDB CLN): entry / search boxes, the 38-column list, View / Back, Excel print. */
@Component({
  selector: 'app-entity-clinic-panel',
  imports: [FormsModule, FieldComponent, DateInputComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      @if (!viewing()) {
        <fieldset aria-label="Clinic Info">
          <legend>Clinic Info</legend>
          @for (f of fields; track f.key) {
            <app-field [label]="f.label" [required]="!!f.required" [hint]="f.searchable ? 'Searchable' : undefined" [wide]="!!f.wide">
              @switch (f.kind) {
                @case ('select') {
                  <select [name]="f.key" [attr.name]="f.key" [disabled]="!enabled(f)" [(ngModel)]="c[f.key]">
                    <option value=""></option>
                    @for (o of options(f); track o[0]) {
                      <option [value]="o[0]">{{ o[1] }}</option>
                    }
                  </select>
                }
                @case ('list') {
                  <input [name]="f.key" [attr.name]="f.key" [attr.list]="'cln-' + f.key" maxlength="30" [disabled]="!enabled(f)" [(ngModel)]="c[f.key]" />
                  <datalist [id]="'cln-' + f.key">
                    @for (o of (f.key === 'city' ? cities() : states()); track o.code) {
                      <option [value]="o.code"></option>
                    }
                  </datalist>
                }
                @case ('date') {
                  <app-date-input [name]="f.key" [attr.name]="f.key" [ariaLabel]="f.label" [disabled]="!enabled(f)" [(ngModel)]="c[f.key]" (focusout)="checkDate()" />
                }
                @case ('area') {
                  <textarea [name]="f.key" [attr.name]="f.key" rows="2" [attr.maxlength]="f.max" [disabled]="!enabled(f)" [(ngModel)]="c[f.key]"></textarea>
                }
                @default {
                  <input [name]="f.key" [attr.name]="f.key" [attr.maxlength]="f.max" [attr.inputmode]="f.kind === 'digits' ? 'numeric' : null" [disabled]="!enabled(f)"
                    [ngModel]="c[f.key]" (ngModelChange)="c[f.key] = f.kind === 'digits' ? digits($event) : $event" />
                }
              }
            </app-field>
          }
          <p class="muted wide">Pay. Post Code / City / State: {{ payPlace() || '—' }} (shown from the clinic; the desktop never saved them).</p>
        </fieldset>
      }

      <app-action-bar label="Clinic Info actions">
        <button appButton variant="primary" [busy]="busy()" [disabled]="mode() === 'add' || mode() === 'edit'" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [disabled]="mode() !== 'selected'" (click)="mode.set('edit')">Edit</button>
        <button appButton variant="danger" [disabled]="mode() !== 'selected'" (click)="remove()">Delete</button>
        <button appButton [disabled]="mode() !== 'add' && mode() !== 'edit'" [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        <button appButton variant="ghost" (click)="refresh()">Refresh</button>
        <button appButton variant="ghost" [busy]="busy()" (click)="print()">Print</button>
        <button appButton variant="ghost" (click)="viewing.set(!viewing())">{{ viewing() ? 'Back' : 'View' }}</button>
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="rows().length" />
      </app-action-bar>

      <app-data-grid caption="Clinics" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
        [selectedKey]="selected()" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class EntityClinicPanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly fields = FIELDS
  protected readonly columns: GridColumn<Row>[] = LIST.map(([key, label]) => ({
    key, label, align: key === 'CLNNoofDoc' ? 'right' : undefined,
    value: (r: Row) => (key === 'CLNReceivedDate' ? formatDate(r[key]) : (r[key] as string | number | null)),
  }))
  protected readonly groups = signal<LookupOption[]>([])
  protected readonly cities = signal<LookupOption[]>([])
  protected readonly states = signal<LookupOption[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly selected = signal<string | null>(null)
  protected readonly mode = signal<Mode>('search')
  protected readonly busy = signal(false)
  protected readonly viewing = signal(false)
  protected readonly payPlace = signal('')
  protected readonly rowKey = (r: Row) => r._id
  protected readonly digits = digitsOnly
  protected c: Clinic = { ...EMPTY, bnfType: '1' }
  private lastFilters: Clinic = { ...EMPTY }

  async ngOnInit(): Promise<void> {
    await this.refresh()
  }

  /** CmdRefresh_Click → ClearForm + CLNComboListing (Benefit Type defaults to "1 - Standard"). */
  protected async refresh(): Promise<void> {
    this.c = { ...EMPTY, bnfType: '1' }
    this.mode.set('search')
    try {
      const s = await firstValueFrom(this.http.get<{ groups: LookupOption[]; cities: LookupOption[]; states: LookupOption[] }>(`${BASE}/setup`))
      this.groups.set(s.groups)
      this.cities.set(s.cities)
      this.states.set(s.states)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected options(f: FieldDef): [string, string][] {
    return f.key === 'groupCode' ? this.groups().map((g) => [g.code, `${g.code} - ${g.name}`]) : (f.options ?? [])
  }

  /** Search: searchable boxes only · Add / Edit: all but the code (Edit also locks Clinic Name, as cmdEdit_Click does). */
  protected enabled(f: FieldDef): boolean {
    const m = this.mode()
    if (m === 'search') return true
    if (m === 'selected') return false
    if (f.key === 'code') return false
    return !(m === 'edit' && f.key === 'name')
  }

  /** txtInfoClnReceivedDate_LostFocus: warnings only. */
  protected checkDate(): void {
    if (!this.c.receivedDate) return
    const d = parseDmy(this.c.receivedDate)
    if (!d) this.notify.notify('Invalid date format!', 'error')
    else if (d > new Date()) this.notify.notify('Date cannot greater than system date ! ', 'error')
  }

  /** SearchCLN → CLNListing; the boxes are cleared afterwards (cmdSearch_Click). */
  protected async search(filters: Clinic = this.c): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.selected.set(null)
    this.payPlace.set('')
    try {
      const rows = (await firstValueFrom(this.http.post<Record<string, unknown>[] | null>(`${BASE}/search`, { clinic: filters }))) ?? []
      this.rows.set(rows.map((r, i) => ({ ...r, _id: String(i) })))
      this.lastFilters = { ...filters }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
      this.c = { ...EMPTY }
      this.mode.set('search')
    }
  }

  /** lstClinicInfoList_Click: every box from the row, read-only. */
  protected select(r: Row): void {
    this.selected.set(r._id)
    const v = (col: string) => (r[col] == null ? '' : String(r[col]).trim())
    this.c = Object.fromEntries(FIELDS.map((f) => [f.key, f.kind === 'date' ? formatDate(r[f.column]) : v(f.column)])) as Clinic
    this.payPlace.set([v('CLNPPostCode'), v('CLNPCity'), v('CLNPState')].filter(Boolean).join(' · '))
    this.mode.set('selected')
  }

  protected add(): void {
    this.c = { ...EMPTY }
    this.selected.set(null)
    this.payPlace.set('')
    this.mode.set('add')
  }

  protected clear(): void {
    this.c = { ...EMPTY }
    this.rows.set([])
    this.selected.set(null)
    this.payPlace.set('')
    this.mode.set('search')
  }

  /** cmdSave_Click → SaveClinic, then SearchCLN with the saved boxes (finds the clinic). */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask('Do you want to save the record?'))) return
    const editing = this.mode() === 'edit'
    this.busy.set(true)
    try {
      const answer = editing
        ? await firstValueFrom(this.http.put<{ message: string; code: string }>(BASE, { clinic: this.c }))
        : await firstValueFrom(this.http.post<{ message: string; code: string }>(BASE, { clinic: this.c }))
      await this.search({ ...this.c, code: answer.code })
      this.notify.notify(answer.message, 'success')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** cmdDelete_Click → DeleteClinic: then the boxes and the list are cleared. */
  protected async remove(): Promise<void> {
    if (!(await this.notify.ask('Do you want to delete the record?'))) return
    try {
      const url = `${BASE}/${encodeURIComponent(this.c.code)}?name=${encodeURIComponent(this.c.name)}`
      const answer = await firstValueFrom(this.http.delete<{ message: string }>(url))
      this.notify.notify(answer.message, 'success')
      this.clear()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdPrint_Click (tab 2) → ExportToExcelClinic of the listed clinics. */
  protected async print(): Promise<void> {
    if (this.rows().length === 0) return this.notify.notify('Report cannot be printed without any record.', 'info')
    this.busy.set(true)
    try {
      const res: HttpResponse<Blob> = await firstValueFrom(
        this.http.post(`${BASE}/export`, { clinic: this.lastFilters }, { responseType: 'blob', observe: 'response' }),
      )
      saveBlob(res.body!, fileNameFromDisposition(res.headers.get('content-disposition'), 'ClinicListing.xlsx'))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }
}
