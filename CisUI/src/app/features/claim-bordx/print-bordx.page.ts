import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { ActivatedRoute } from '@angular/router'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/claim-bordx/print'
type Row = Record<string, unknown>
interface Info {
  bordx: Row
  printed: boolean
  rows: Row[]
  totalAmount: number
}
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const m = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => (r[key] == null ? '' : formatMoney(Number(r[key]))) })

/** Claim Bordx > Print Bordx (frmPrintBordxExcel). */
@Component({
  selector: 'app-print-bordx-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="title" [subtitle]="'Replaces ' + form">
      <div class="screen-form">
        <fieldset aria-label="Bordereaux">
          <legend>Bordereaux</legend>
          <app-field label="Bordx No" [required]="true"><input name="bordxNo" data-autofocus maxlength="20" [(ngModel)]="f.bordxNo" (keydown.enter)="search()" /></app-field>
          <app-field label="Group Company"><input name="groupCompany" disabled [value]="s(info()?.bordx?.['groupCompany'] ?? info()?.bordx?.['GroupCompany'])" /></app-field>
          <app-field label="Month">
            <select name="month" [(ngModel)]="f.month" [disabled]="!!info()?.printed"><option value=""></option>@for (x of months; track x) { <option [value]="x">{{ x }}</option> }</select>
          </app-field>
          <app-field label="Year">
            <select name="year" [(ngModel)]="f.year" [disabled]="!!info()?.printed"><option value=""></option>@for (y of years; track y) { <option [value]="y">{{ y }}</option> }</select>
          </app-field>
          @if (variant === 'standard') {
            <app-field label="Company" [required]="true">
              <select name="company" [(ngModel)]="f.company"><option value=""></option>@for (c of companies; track c) { <option [value]="c">{{ c }}</option> }</select>
            </app-field>
            <app-field label="Bill Type" [required]="true">
              <select name="billType" [(ngModel)]="f.billType"><option [ngValue]="null"></option><option [ngValue]="0">Invoice</option><option [ngValue]="1">Disbursement Note</option></select>
            </app-field>
          }
          <app-field label="Report Type">
            <select name="reportType" [(ngModel)]="f.reportType">@for (t of reportTypes; track $index) { <option [ngValue]="$index">{{ t }}</option> }</select>
          </app-field>
          <app-action-bar label="Print bordx actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton [disabled]="!info()" (click)="print(false, false)">Detail Preview</button>
            <button appButton [disabled]="!info()" (click)="print(true, false)">Print Details</button>
            <button appButton [disabled]="!info()" (click)="print(false, true)">Summary Preview</button>
            <button appButton [disabled]="!info()" (click)="print(true, true)">Print Summary</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="Bordereaux worksheets" [columns]="columns" [rows]="info()?.rows ?? []" [rowKey]="key" />
        <p class="totals" aria-label="Totals">
          Total: <strong>{{ info()?.rows?.length ?? 0 }}</strong> · Total Amount: <strong>{{ money(info()?.totalAmount ?? 0) }}</strong> ·
          Printed: <strong>{{ info() ? (info()!.printed ? 'YES' : 'NO') : '-' }}</strong>
        </p>
      </div>
    </app-page-shell>
  `,
})
export class PrintBordxPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  private readonly data = inject(ActivatedRoute).snapshot.data
  protected readonly variant: string = this.data['variant'] ?? 'standard'
  protected readonly title: string = this.data['screenTitle'] ?? 'Print Bordereaux'
  protected readonly form: string = this.data['legacyForm'] ?? 'frmPrintBordxExcel'
  protected readonly s = s
  protected readonly money = (v: number) => formatMoney(v)
  protected readonly months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
  protected readonly years = Array.from({ length: new Date().getFullYear() + 2 - 2012 + 1 }, (_, i) => String(2012 + i))
  protected readonly companies = ['MediExpress', 'Health Connect', 'Sihat Johor']
  protected readonly reportTypes = ['By Group Company', 'By Department', 'By Relationship', 'By Cost Center', 'By Cost Center-Branch Code', 'By Subsidiary']
  protected readonly info = signal<Info | null>(null)
  protected readonly busy = signal(false)
  protected f = this.empty()
  protected readonly key = (r: Row) => s(r['CASNumber'])
  protected readonly columns: GridColumn<Row>[] = [
    { key: 'PLNCode', label: 'Plan' }, { key: 'CASNumber', label: 'Case No' }, { key: 'Name', label: 'Name' }, { key: 'MBMnumber', label: 'Member No' },
    { key: 'ClinicCode', label: 'Clinic' }, { key: 'CLNName', label: 'Clinic Name' }, m('CLMTotalActual', 'Bill Amt'), m('CLMCoPayCLNColl', 'Co-Pay'),
    m('CLMFinalApprove', 'Approved'), m('CLMMCOProcFee', 'MCO Fee'), m('CLMBordxAmt', 'Bordx Amt'), { key: 'BordxRefNo', label: 'Bordx No' },
    { key: 'DateVisit', label: 'Visit Date', value: (r) => (r['DateVisit'] ? formatDate(String(r['DateVisit'])) : '') },
  ]

  private empty() {
    return { bordxNo: '', month: '', year: '', company: '', billType: null as number | null, reportType: 0 }
  }

  /** cmdSearch_Click / GetBordxInfo. */
  protected async search(): Promise<void> {
    this.busy.set(true)
    this.info.set(null)
    try {
      const r = await firstValueFrom(this.http.get<Info>(BASE, { params: { bordxNo: this.f.bordxNo } }))
      this.info.set(r)
      this.f.month = s(r.bordx['BORDXPrintMonth'] ?? r.bordx['bordxPrintMonth'])
      this.f.year = s(r.bordx['BORDXPrintYear'] ?? r.bordx['bordxPrintYear'])
      this.f.reportType = s(r.bordx['RptTypeByDept'] ?? r.bordx['rptTypeByDept']) === '1' ? 1 : 0
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  protected async print(print: boolean, summary: boolean): Promise<void> {
    try {
      const blob = await firstValueFrom(this.http.post(BASE, { ...this.f, print, summary, variant: this.variant }, { responseType: 'blob' }))
      saveBlob(blob, `PrintBordx_${this.f.bordxNo.trim()}.xlsx`)
      if (print && !this.info()?.printed) await this.search()
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }

  protected clear(): void {
    this.f = this.empty()
    this.info.set(null)
  }
}
