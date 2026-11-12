import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { formatDate } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/tracking/clinic'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const COLS: [string, string][] = [
  ['CLNCode', 'Code'], ['CLNName', 'Name'], ['ClinicGrpCode', 'Group'], ['CLNInitial', 'Initial'], ['CLNType', 'Type'], ['CLNBNFType', 'BNF Type'],
  ['CLNAdd1', 'Address 1'], ['CLNAdd2', 'Address 2'], ['CLNAdd3', 'Address 3'], ['CLNPostCode', 'Postcode'], ['CLNCity', 'City'], ['CLNState', 'State'],
  ['CLNReceivedDate', 'Received'], ['CLNPanel', 'Panel'], ['CLNXRay', 'X-Ray'], ['CLNForRef', 'For Ref'], ['CLNNoofDoc', 'No of Doc'], ['CLNPhone', 'Phone'],
  ['CLNFax', 'Fax'], ['CLNEmail', 'Email'], ['CLNHomepage', 'Homepage'], ['CLNBankAcc', 'Bank Acc'], ['CLNBankName', 'Bank'], ['CLNBankBrn', 'Branch'],
  ['CLNAccPerson', 'Acc Person'], ['CLN24Hrs', '24 Hrs'], ['CLNRemark', 'Remark'], ['CLNLocation', 'Location'], ['CLNFFSStatus', 'FFS'],
]

/** Tracking > Clinic (FrmTracking). */
@Component({
  selector: 'app-clinic-tracking-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Clinic Tracking" subtitle="Replaces FrmTracking">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Clinic Code"><input name="code" maxlength="10" [(ngModel)]="f.code" /></app-field>
          <app-field label="Clinic Name"><input name="name" maxlength="200" [(ngModel)]="f.name" /></app-field>
          <app-field label="Clinic Group" [wide]="true">
            <select name="group" [(ngModel)]="f.group"><option value=""></option>@for (g of groups(); track g) { <option [value]="g">{{ g }}</option> }</select>
          </app-field>
          <app-field label="Postcode"><input name="postcode" maxlength="10" [(ngModel)]="f.postcode" /></app-field>
          <app-field label="City">
            <select name="city" [(ngModel)]="f.city"><option value=""></option>@for (c of cities(); track c) { <option [value]="c">{{ c }}</option> }</select>
          </app-field>
          <app-field label="State">
            <select name="state" [(ngModel)]="f.state"><option value=""></option>@for (c of states(); track c) { <option [value]="c">{{ c }}</option> }</select>
          </app-field>
          <app-field label="Contact Person"><input name="contact" maxlength="200" [(ngModel)]="f.contact" /></app-field>
          <app-field label="Doctor Name"><input name="doctor" maxlength="200" [(ngModel)]="f.doctor" /></app-field>
          <app-field label="MIX Clinic"><input name="initial" maxlength="100" [(ngModel)]="f.initial" /></app-field>
          <app-field label="Location"><input name="location" maxlength="200" [(ngModel)]="f.location" /></app-field>
          <app-field label="For Ref"><input name="forRef" maxlength="50" [(ngModel)]="f.forRef" /></app-field>
          <app-field label="BNF Type">
            <select name="bnfType" [(ngModel)]="f.bnfType"><option value=""></option><option>1 - Standard</option><option>2 - Plan</option></select>
          </app-field>
          <app-field label="Clinic Type">
            <select name="type" [(ngModel)]="f.type"><option value=""></option><option>G - General Practisional</option><option>S - Specialised</option></select>
          </app-field>
          <app-field label="Panel">
            <select name="panel" [(ngModel)]="f.panel"><option value=""></option><option>Y - Yes</option><option>N - No</option><option>C - Confirm</option></select>
          </app-field>
          <app-field label="X-Ray">
            <select name="xray" [(ngModel)]="f.xray"><option value=""></option><option>Y - Yes</option><option>N - No</option></select>
          </app-field>
          <app-field label="FFS">
            <select name="ffs" [(ngModel)]="f.ffs"><option value=""></option><option>Y - Yes</option><option>N - No</option></select>
          </app-field>
          <app-field label="Received From"><app-date-input name="receivedFrom" ariaLabel="Received From" [(ngModel)]="f.receivedFrom" /></app-field>
          <app-field label="Received To"><app-date-input name="receivedTo" ariaLabel="Received To" [(ngModel)]="f.receivedTo" /></app-field>
          <app-action-bar label="Clinic tracking actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Clinics" [columns]="columns" [rows]="rows()" [rowKey]="key" />
        <p class="totals" aria-label="Totals">Records: <strong>{{ rows().length }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class ClinicTrackingPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly groups = signal<string[]>([])
  protected readonly cities = signal<string[]>([])
  protected readonly states = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly busy = signal(false)
  protected f = this.empty()
  protected readonly key = (x: Row) => s(x['CLNCode'])
  protected readonly columns: GridColumn<Row>[] = COLS.map(([key, label]) =>
    key === 'CLNReceivedDate' ? { key, label, value: (x: Row) => (x[key] ? formatDate(String(x[key])) : '') } : { key, label })

  private empty() {
    return {
      code: '', name: '', group: '', postcode: '', city: '', state: '', contact: '', doctor: '', initial: '', location: '', forRef: '', bnfType: '', type: '',
      panel: '', xray: '', ffs: '', receivedFrom: '', receivedTo: '',
    }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ groups: string[]; cities: string[]; states: string[] }>(`${BASE}/lookups`))
      this.groups.set(l.groups)
      this.cities.set(l.cities)
      this.states.set(l.states)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    try {
      this.rows.set(await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, this.f)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected print(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'ClinicTracking.csv')
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
  }
}
