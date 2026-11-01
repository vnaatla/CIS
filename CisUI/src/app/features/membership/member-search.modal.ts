import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, input, output, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { formatDate } from '../../shared/lib/format'
import { ModalComponent } from '../../shared/modal.component'
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

/** One row of the member search pop-ups (api/membership/search). */
export interface MemberRow {
  mbmNumber: string
  coverId: string | null
  name: string | null
  ic: string | null
  policyNo: string | null
  policyStatus: string | null
  dataStatus: string | null
  effDate: string | null
  expDate: string | null
  groupCompany: string | null
  department: string | null
  supp: boolean
}

interface SearchResult {
  rows: MemberRow[]
  principals: number
  supps: number
  capped: boolean
}

/** search = frmSearchMBM (type and search) · ic = frmSearchMBMIC (patient IC, runs on open) · prefix = frmSearchMBM05 (member no., runs on open). */
export type MemberSearchMode = 'search' | 'ic' | 'prefix'

const OPTIONS: [string, string][] = [
  ['MemberNo', 'Mem Num'],
  ['PolicyNo', 'Policy Num'],
  ['Name', 'Mem Name'],
  ['IC', 'IC Num'],
  ['EmployeeNo', 'Employee Num'],
  ['PolicySubNo', 'Policy Sub Num'],
]

const date = (k: 'effDate' | 'expDate') => ({ value: (r: MemberRow) => formatDate(r[k]), sortBy: (r: MemberRow) => r[k] ?? '' })

/** The member search pop-ups; a click on a row returns it (the desktop's double-click). */
@Component({
  selector: 'app-member-search-modal',
  imports: [FormsModule, ModalComponent, FieldComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-modal [title]="mode() === 'search' ? 'Search For Member' : 'Search MBM'" [wide]="true" (closed)="closed.emit()">
      <div class="screen-form">
        @if (mode() === 'search') {
          <fieldset aria-label="Search">
            <legend>Search</legend>
            <app-field label="Search" [required]="true" hint="At least 3 characters">
              <input name="memberSearchText" data-autofocus [(ngModel)]="text" (keydown.enter)="search()" />
            </app-field>
            <div class="radio-row" role="radiogroup" aria-label="Search by">
              @for (o of options; track o[0]) {
                <label><input type="radio" name="memberSearchBy" [value]="o[0]" [(ngModel)]="by" /> {{ o[1] }}</label>
              }
            </div>
          </fieldset>
        }
        <app-action-bar label="Member search actions">
          @if (mode() === 'search') {
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton variant="ghost" (click)="clear()">Clear</button>
          } @else {
            <button appButton variant="ghost" [disabled]="rows().length === 0" (click)="print()">Print</button>
          }
          <span class="spacer"></span>
          <app-stat label="Principal" [value]="principals()" />
          <app-stat label="Supp" [value]="supps()" />
          <app-stat label="Record(s)" [value]="rows().length" />
        </app-action-bar>
        @if (capped()) {
          <p class="muted">Only the first 5,000 principals / supplementary members are listed; narrow the search.</p>
        }
        <app-data-grid caption="Members found" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true" (rowSelected)="picked.emit($event)" />
        <p class="muted">Click a member to choose it.</p>
      </div>
    </app-modal>
  `,
})
export class MemberSearchModalComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  readonly mode = input<MemberSearchMode>('search')
  /** Caller's box: the patient IC (ic) or member number (prefix). */
  readonly initialText = input('')
  readonly picked = output<MemberRow>()
  readonly closed = output<void>()

  protected readonly options = OPTIONS
  protected readonly rows = signal<MemberRow[]>([])
  protected readonly principals = signal(0)
  protected readonly supps = signal(0)
  protected readonly capped = signal(false)
  protected readonly busy = signal(false)
  protected readonly rowKey = (r: MemberRow) => `${r.supp ? 'S' : 'P'}|${r.mbmNumber}|${r.coverId ?? ''}|${r.name ?? ''}`
  protected text = ''
  /** Form_Load / cmdClear_Click: Mem Name is selected. */
  protected by = 'Name'
  protected columns: GridColumn<MemberRow>[] = []

  async ngOnInit(): Promise<void> {
    const search = this.mode() === 'search'
    this.columns = [
      { key: 'mbmNumber', label: 'Member No' },
      { key: 'coverId', label: 'ID' },
      { key: 'name', label: 'Member Name' },
      { key: 'ic', label: 'IC No.' },
      ...(search
        ? [{ key: 'policyStatus', label: 'PolSt' }, { key: 'dataStatus', label: 'DataSt' }]
        : [{ key: 'policyNo', label: 'Policy No' }, { key: 'policyStatus', label: 'Status' }]),
      { key: 'effDate', label: 'Eff Date', ...date('effDate') },
      { key: 'expDate', label: 'Exp Date', ...date('expDate') },
      { key: 'groupCompany', label: search ? 'Group Company' : 'Group Comp' },
      { key: 'department', label: 'Dept' },
    ]
    // frmSearchMBMIC / frmSearchMBM05 search as soon as they open.
    if (this.mode() === 'ic') await this.load(new HttpParams({ fromObject: { by: 'IcExact', text: this.initialText() } }), '/api/membership/search')
    if (this.mode() === 'prefix') await this.load(new HttpParams({ fromObject: { memberNo: this.initialText() } }), '/api/membership/search/prefix')
  }

  /** cmdSearch_Click. */
  protected async search(): Promise<void> {
    await this.load(new HttpParams({ fromObject: { by: this.by, text: this.text } }), '/api/membership/search')
  }

  /** cmdClear_Click: text, list and counters; Mem Name selected again. */
  protected clear(): void {
    this.text = ''
    this.by = 'Name'
    this.set(null)
  }

  /** CmdPrint_Click (frmSearchMBM05 / IC): the list to Excel. */
  protected print(): void {
    const header = this.columns.map((c) => c.label)
    const body = this.rows().map((r) => this.columns.map((c) => (c.value ? c.value(r) : (r[c.key as keyof MemberRow] as string | null))))
    saveCsv([header, ...body], 'MemberSearch.csv')
  }

  private async load(params: HttpParams, url: string): Promise<void> {
    this.busy.set(true)
    this.set(null)
    try {
      this.set(await firstValueFrom(this.http.get<SearchResult | null>(url, { params })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  private set(r: SearchResult | null): void {
    this.rows.set(r?.rows ?? [])
    this.principals.set(r?.principals ?? 0)
    this.supps.set(r?.supps ?? 0)
    this.capped.set(r?.capped ?? false)
  }
}
