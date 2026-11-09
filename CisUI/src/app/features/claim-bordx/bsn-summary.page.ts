import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())

/** Claim Bordx > Print Bordx > BSN-PrintSummary (frmBSNPrintSummary). */
@Component({
  selector: 'app-bsn-summary-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="BSN - Print Summary" subtitle="Replaces frmBSNPrintSummary">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          @for (b of bordx(); track $index) {
            <app-field [label]="$index === 0 ? 'Bordx Ref No' : 'Bordx Ref No ' + ($index + 1)">
              <input [attr.name]="'bordx' + $index" maxlength="20" [ngModel]="b" (ngModelChange)="set($index, $event)" (keydown.enter)="search()" />
            </app-field>
          }
          <app-action-bar label="BSN summary actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="add()">Add Bordx</button>
            <button appButton (click)="export()">Export</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="BSN bordereaux" [columns]="columns" [rows]="rows()" [rowKey]="key" />
      </div>
    </app-page-shell>
  `,
})
export class BsnSummaryPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly bordx = signal<string[]>([''])
  protected readonly rows = signal<Row[]>([])
  protected readonly busy = signal(false)
  protected readonly key = (x: Row) => s(x['BordxRefNo'])
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'BordxRefNo', label: 'Bordx Ref No' }, { key: 'NOOFCLAIMS', label: 'No of Claims', align: 'right' },
    { key: 'CLMBordxAmt', label: 'Bordx Amount', align: 'right', value: (x) => (x['CLMBordxAmt'] == null ? '' : formatMoney(Number(x['CLMBordxAmt']))) },
    { key: 'BordxDate', label: 'Bordx Date', value: (x) => (x['BordxDate'] ? formatDate(String(x['BordxDate'])) : '') },
  ]

  protected set(i: number, v: string): void {
    this.bordx.update((b) => b.map((x, j) => (j === i ? v : x)))
  }

  /** cmdincreasebutton_Click: one more bordx box. */
  protected add(): void {
    this.bordx.update((b) => [...b, ''])
  }

  /** cmdsearch_Click. */
  protected async search(): Promise<void> {
    const list = this.bordx().map((b) => b.trim()).filter((b) => b)
    if (!list.length) return this.notify.notify('Please Enter Bordx Refno..! ', 'error')
    this.busy.set(true)
    try {
      let params = new HttpParams()
      for (const b of list) params = params.append('bordxNo', b)
      this.rows.set(await firstValueFrom(this.http.get<Row[]>('/api/claim-bordx/bsn-summary', { params })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected export(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    saveCsv([this.columns.map((c) => c.label), ...this.rows().map((x) => this.columns.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'BSNPrintSummary.csv')
  }

  protected clear(): void {
    this.bordx.set([''])
    this.rows.set([])
  }
}
