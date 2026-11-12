import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/tracking/grp-company-clinic'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const first = (k: string, label: string): GridColumn<Row> => ({ key: k, label, value: (x) => s(x[k]) || 'Empty' })
const COLUMNS: Record<number, GridColumn<Row>[]> = {
  1: [first('GRPComCode', 'Group Code'), { key: 'GrpCoName', label: 'Group Name' }, { key: 'GrpComClnCode', label: 'Clinic Code' }, { key: 'CLNName', label: 'Clinic Name' }],
  2: [first('GRPComCode', 'Group Code'), { key: 'GrpCoName', label: 'Group Name' }, { key: 'TotalCLN', label: 'No of Clinic', align: 'right' }],
  3: [first('GrpComClnCode', 'Clinic Code'), { key: 'CLNName', label: 'Clinic Name' }, { key: 'GRPComCode', label: 'Group Code' }, { key: 'GrpCoName', label: 'Group Name' }],
  4: [first('GrpComClnCode', 'Clinic Code'), { key: 'CLNName', label: 'Clinic Name' }, { key: 'TotalCom', label: 'No of Group Company', align: 'right' }],
}

/** Tracking > Grp. Com. Clinic (FrmTracking2). */
@Component({
  selector: 'app-grp-company-clinic-tracking-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Group Company Clinic Tracking" subtitle="Replaces FrmTracking2">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="View" [wide]="true">
            <span class="radio-row">
              @for (v of views; track v.mode) {
                <label class="check-field"><input type="radio" name="mode" [value]="v.mode" [ngModel]="mode()" (ngModelChange)="setMode($event)" /> {{ v.label }}</label>
              }
            </span>
          </app-field>
          <app-field label="Group Code"><input name="groupCode" maxlength="30" [(ngModel)]="f.groupCode" /></app-field>
          <app-field label="Clinic Code"><input name="clinicCode" maxlength="30" [(ngModel)]="f.clinicCode" /></app-field>
          <app-field label="Group Company" [wide]="true">
            <input name="groupCompany" list="gcc-companies" [(ngModel)]="f.groupCompany" />
            <datalist id="gcc-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-field label="Clinic Name" [wide]="true">
            <input name="clinic" list="gcc-clinics" [(ngModel)]="f.clinic" />
            <datalist id="gcc-clinics">@for (c of clinics(); track $index) { <option [value]="c"></option> }</datalist>
          </app-field>
          <app-action-bar label="Tracking actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Group company clinics" [columns]="columns()" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">Records: <strong>{{ rows().length }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class GrpCompanyClinicTrackingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly views = [
    { mode: 1, label: 'By group company' }, { mode: 2, label: 'Clinics per group company' }, { mode: 3, label: 'By clinic' }, { mode: 4, label: 'Group companies per clinic' },
  ]
  protected readonly companies = signal<string[]>([])
  protected readonly clinics = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly mode = signal(1)
  protected readonly busy = signal(false)
  protected readonly columns = computed(() => COLUMNS[this.mode()])
  protected f = { groupCode: '', clinicCode: '', groupCompany: '', clinic: '' }
  protected readonly key = (x: Row) => `${this.rows().indexOf(x)}`

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ companies: string[]; clinics: string[] }>(`${BASE}/lookups`))
      this.companies.set(l.companies)
      this.clinics.set(l.clinics)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected setMode(m: number): void {
    this.mode.set(Number(m))
    this.rows.set([])
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    try {
      this.rows.set(await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, { ...this.f, mode: this.mode() })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    const cols = this.columns()
    saveCsv([cols.map((c) => c.label), ...this.rows().map((x) => cols.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'GrpCompanyClinicTracking.csv')
  }

  protected clear(): void {
    this.f = { groupCode: '', clinicCode: '', groupCompany: '', clinic: '' }
    this.rows.set([])
  }
}
