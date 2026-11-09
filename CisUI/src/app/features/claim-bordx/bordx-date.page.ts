import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { formatDate } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/claim-bordx/bordx-date'
interface Line {
  bordxRefNo: string
  bordxDate: string | null
  postingStatus: string
  noOfCases: number
}

/** Claim Bordx > ClaimBordxDate (FrmBordxDate). */
@Component({
  selector: 'app-bordx-date-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Change Bordx Date" subtitle="Replaces FrmBordxDate">
      <div class="screen-form">
        <fieldset aria-label="Bordereaux">
          <legend>Bordereaux</legend>
          <app-field label="Bordx No"><input name="bordxNo" data-autofocus maxlength="10" [(ngModel)]="bordxNo" (keydown.enter)="search()" /></app-field>
          <app-field label="New Bordx Date"><app-date-input name="bordxDate" ariaLabel="New Bordx Date" [(ngModel)]="bordxDate" /></app-field>
          <app-field label="Total Record"><input name="total" disabled [value]="rows().length || ''" /></app-field>
          <app-action-bar label="Bordx date actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton variant="primary" (click)="update()">Update</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Bordx lines" [columns]="columns" [rows]="rows()" [rowKey]="key" />
      </div>
    </app-page-shell>
  `,
})
export class BordxDatePage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly rows = signal<Line[]>([])
  protected readonly busy = signal(false)
  protected bordxNo = ''
  protected bordxDate = ''
  protected readonly key = (r: Line) => `${r.bordxRefNo}-${r.postingStatus}`
  protected readonly columns: GridColumn<Line>[] = [
    { key: 'bordxRefNo', label: 'Bordx No' }, { key: 'bordxDate', label: 'Bordx Date', value: (r) => (r.bordxDate ? formatDate(r.bordxDate) : '') },
    { key: 'postingStatus', label: 'Posting Status' }, { key: 'noOfCases', label: 'No of Cases', align: 'right' },
  ]

  /** CmdSearch_Click. */
  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    try {
      const r = await firstValueFrom(this.http.get<{ rows: Line[]; message: string | null }>(BASE, { params: { bordxNo: this.bordxNo } }))
      this.rows.set(r.rows)
      if (r.message) this.notify.notify(r.message, 'info')
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  /** CmdUpdate_Click. */
  protected async update(): Promise<void> {
    try {
      const r = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/update`, { bordxNo: this.bordxNo, bordxDate: this.bordxDate }))
      this.notify.notify(r.message, 'success')
      await this.search()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }
}
