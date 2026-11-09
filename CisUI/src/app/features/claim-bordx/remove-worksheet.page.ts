import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/claim-bordx/remove-worksheet'
type Row = Record<string, unknown>
interface Result {
  rows: Row[]
  printedDate: string | null
  bordxDate: string | null
  total: number
  totalApproved: number
}

const d = (v: unknown) => (v ? formatDate(String(v)) : '')

/** Claim Bordx > Worksheet > Remove Worksheet (FrmBordxRevise). */
@Component({
  selector: 'app-remove-worksheet-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Remove Worksheet from Bordereaux" subtitle="Replaces FrmBordxRevise">
      <div class="screen-form">
        <fieldset aria-label="Bordereaux">
          <legend>Bordereaux</legend>
          <app-field label="Bordx No" [required]="true"><input name="bordxNo" data-autofocus maxlength="20" [(ngModel)]="bordxNo" (keydown.enter)="go()" /></app-field>
          <app-field label="Bordx Date"><input name="bordxDate" disabled [value]="d(result()?.bordxDate)" /></app-field>
          <app-field label="Printed Date"><input name="printedDate" disabled [value]="d(result()?.printedDate)" /></app-field>
          <app-field label="Total"><input name="total" disabled [value]="result()?.total ?? ''" /></app-field>
          <app-field label="Total Amount"><input class="num" name="totalAmount" disabled [value]="result() ? money(result()!.totalApproved) : ''" /></app-field>
          <app-action-bar label="Remove worksheet actions">
            <button appButton variant="primary" [busy]="busy()" (click)="go()">Go</button>
            <button appButton variant="danger" [disabled]="!checked().size" (click)="reject()">Reject</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Bordereaux worksheets" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()"
          (checkedChange)="checked.set($event)" />
      </div>
    </app-page-shell>
  `,
})
export class RemoveWorksheetPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly d = d
  protected readonly money = (v: number) => (v < 0 ? `(${formatMoney(-v)})` : formatMoney(v))
  protected readonly rows = signal<Row[]>([])
  protected readonly result = signal<Result | null>(null)
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly busy = signal(false)
  protected bordxNo = ''
  private searched = ''
  protected readonly key = (r: Row) => String(r['CASNumber'] ?? '').trim()
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'CASNumber', label: 'Worksheet No' }, { key: 'ClaimOpendate', label: 'Open Date', value: (r) => d(r['ClaimOpendate']) },
    { key: 'MBMNumber', label: 'Member No' }, { key: 'PatientName', label: 'Patient' }, { key: 'DateVisit', label: 'Visit Date', value: (r) => d(r['DateVisit']) },
    { key: 'CLMTotalApproved', label: 'Approved', align: 'right', value: (r) => (r['CLMTotalApproved'] == null ? '' : formatMoney(Number(r['CLMTotalApproved']))) },
    { key: 'CASStatus', label: 'Status' },
  ]

  /** CmdGo_Click. */
  protected async go(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    this.result.set(null)
    this.checked.set(new Set())
    try {
      const r = await firstValueFrom(this.http.get<Result>(BASE, { params: { bordxNo: this.bordxNo } }))
      this.searched = this.bordxNo.trim()
      this.result.set(r)
      this.rows.set(r.rows)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** CmdReject_Click. */
  protected async reject(): Promise<void> {
    if (!(await this.notify.ask('You are going to remove a worksheet from current Bordereaux. \nAre you sure?'))) return
    try {
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/remove`, { bordxNo: this.searched, cases: [...this.checked()] }))
      this.notify.notify(r.message, 'success')
      this.clear()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected clear(): void {
    this.bordxNo = ''
    this.rows.set([])
    this.result.set(null)
    this.checked.set(new Set())
  }
}
