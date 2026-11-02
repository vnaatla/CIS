import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, input, output, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { formatDate } from '../../shared/lib/format'
import { ModalComponent } from '../../shared/modal.component'
import { ActionBarComponent } from '../../shared/screen.components'

const BASE = '/api/invoice/popups'

/** A treatment reason picked in the ICD search (frmSearchICD). */
export interface IcdRow {
  reasonCode: string
  reasonDesc: string | null
}

/** A case picked in the case list (frmCasList). */
export interface CaseListRow {
  casNumber: string
  mbmNumber: string
  coverId: string
  name: string
  policyNo: string
  visitDate: string | null
  status: string
  clinic: string
  ic: string
}

/** frmSearchICD: treatment reasons by code (starts with), description (contains) and a second description joined with AND / OR. */
@Component({
  selector: 'app-icd-search-modal',
  imports: [FormsModule, ModalComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-modal title="Search Treatment Reason / ICD" [wide]="true" (closed)="closed.emit()">
      <div class="screen-form">
        <fieldset aria-label="ICD search">
          <legend>Search</legend>
          <app-field label="Code"><input name="icdCode" data-autofocus [(ngModel)]="code" (keydown.enter)="search()" /></app-field>
          <app-field label="Description"><input name="icdDesc" [(ngModel)]="desc" (keydown.enter)="search()" /></app-field>
          <app-field label="Criteria">
            <select name="icdCriteria" [(ngModel)]="criteria"><option value="AND">AND</option><option value="OR">OR</option></select>
          </app-field>
          <app-field label="Description 2"><input name="icdDesc2" [(ngModel)]="desc2" (keydown.enter)="search()" /></app-field>
          <app-action-bar label="ICD search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Treatment reasons" [columns]="columns" [rows]="rows()" [rowKey]="key" [selectable]="true" (rowSelected)="picked.emit($event)" />
      </div>
    </app-modal>
  `,
})
export class IcdSearchModalComponent {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  readonly picked = output<IcdRow>()
  readonly closed = output<void>()
  protected readonly rows = signal<IcdRow[]>([])
  protected readonly busy = signal(false)
  protected readonly key = (r: IcdRow) => r.reasonCode
  protected readonly columns: GridColumn<IcdRow>[] = [{ key: 'reasonCode', label: 'Code' }, { key: 'reasonDesc', label: 'Description' }]
  protected code = ''
  protected desc = ''
  protected desc2 = ''
  protected criteria = 'AND'

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    try {
      const params = new HttpParams({ fromObject: { code: this.code, desc: this.desc, desc2: this.desc2, criteria: this.criteria } })
      this.rows.set(await firstValueFrom(this.http.get<IcdRow[]>(`${BASE}/icd`, { params })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  protected clear(): void {
    this.code = this.desc = this.desc2 = ''
    this.criteria = 'AND'
    this.rows.set([])
  }
}

/** frmCasList: cases of active members by member no, policy no, name, IC or case no (at least 3 characters). */
@Component({
  selector: 'app-case-list-modal',
  imports: [FormsModule, ModalComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-modal title="Case List" [wide]="true" (closed)="closed.emit()">
      <div class="screen-form">
        <fieldset aria-label="Case search">
          <legend>Search</legend>
          <app-field label="Search"><input name="caseSearch" data-autofocus [(ngModel)]="text" (keydown.enter)="search()" /></app-field>
          <div class="radio-row" role="radiogroup" aria-label="Search by">
            @for (o of options; track o[0]) {
              <label><input type="radio" name="caseBy" [value]="o[0]" [(ngModel)]="by" /> {{ o[1] }}</label>
            }
          </div>
          <app-action-bar label="Case search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Cases" [columns]="columns" [rows]="rows()" [rowKey]="key" [selectable]="true" (rowSelected)="picked.emit($event)" />
        <p class="totals">Total: <strong>{{ rows().length }}</strong></p>
      </div>
    </app-modal>
  `,
})
export class CaseListModalComponent {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  /** Dental cases (CaseType D) instead of invoice cases (C). */
  readonly dental = input(false)
  readonly picked = output<CaseListRow>()
  readonly closed = output<void>()
  protected readonly rows = signal<CaseListRow[]>([])
  protected readonly busy = signal(false)
  protected readonly key = (r: CaseListRow) => r.casNumber
  protected readonly options: [number, string][] = [[1, 'Member No'], [2, 'Policy No'], [3, 'Name'], [4, 'IC/BC/PP'], [5, 'Case No']]
  protected readonly columns: GridColumn<CaseListRow>[] = [
    { key: 'casNumber', label: 'Case No' }, { key: 'mbmNumber', label: 'Member No' }, { key: 'coverId', label: 'Cov' }, { key: 'name', label: 'Name' },
    { key: 'policyNo', label: 'Policy No' }, { key: 'visitDate', label: 'Visit Date', value: (r) => formatDate(r.visitDate) },
    { key: 'status', label: 'Status' }, { key: 'clinic', label: 'Clinic' }, { key: 'ic', label: 'IC/BC/PP' },
  ]
  protected text = ''
  protected by = 3

  protected async search(): Promise<void> {
    if (this.text.length < 3) return this.notify.notify('Search String must be at least 3 characters!', 'info')
    this.busy.set(true)
    this.rows.set([])
    try {
      const params = new HttpParams({ fromObject: { dental: this.dental(), by: this.by, text: this.text } })
      this.rows.set(await firstValueFrom(this.http.get<CaseListRow[]>(`${BASE}/cases`, { params })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected clear(): void {
    this.text = ''
    this.by = 3
    this.rows.set([])
  }
}
