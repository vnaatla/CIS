import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, computed, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { DateInputComponent } from '../../shared/date-input.component'
import { FieldComponent } from '../../shared/field.component'
import { saveCsv } from '../../shared/lib/download'
import { formatDate, formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/receipt/payor'
type Row = Record<string, unknown>
const s = (v: unknown) => (v == null ? '' : String(v).trim())
const d = (v: unknown) => (v ? formatDate(String(v)) : '')
const m = (v: unknown) => (v == null || s(v) === '' ? '' : formatMoney(Number(v)))
const money = (key: string, label: string): GridColumn<Row> => ({ key, label, align: 'right', value: (r) => m(r[key]) })

/** Receipt > Fr Grp Company (FrmReceiptPayor). */
@Component({
  selector: 'app-receipt-payor-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, DateInputComponent, ButtonDirective, ActionBarComponent, DataGridComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Receipt from Group Company" subtitle="Replaces FrmReceiptPayor">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Receipt By">
            <span class="radio-row">
              <label class="check-field"><input type="radio" name="mode" [value]="false" [(ngModel)]="f.worksheet" (ngModelChange)="mode()" /> Bordx</label>
              <label class="check-field"><input type="radio" name="mode" [value]="true" [(ngModel)]="f.worksheet" (ngModelChange)="mode()" /> Worksheet</label>
            </span>
          </app-field>
          <app-field label="Scan"><label class="check-field"><input type="checkbox" name="scan" [(ngModel)]="f.scan" /> Scan one {{ f.worksheet ? 'worksheet' : 'bordx' }} no.</label></app-field>
          <app-field [label]="!f.worksheet && f.scan ? 'Bordx No' : 'Bordx No From'"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" /></app-field>
          @if (f.worksheet || !f.scan) {
            <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
          }
          <app-field [label]="f.worksheet && f.scan ? 'Worksheet No' : 'Worksheet No From'"><input name="caseFrom" maxlength="30" [(ngModel)]="f.caseFrom" /></app-field>
          @if (!(f.worksheet && f.scan)) {
            <app-field label="Worksheet No To"><input name="caseTo" maxlength="30" [(ngModel)]="f.caseTo" /></app-field>
          }
          <app-field label="Group Company" [wide]="true">
            <input name="company" list="rp-companies" [(ngModel)]="f.company" (change)="loadSubsidiaries()" />
            <datalist id="rp-companies">@for (c of companies(); track c) { <option [value]="c"></option> }</datalist>
          </app-field>
          @if (f.worksheet) {
            <app-field label="Subsidiary">
              <select name="subsidiary" [(ngModel)]="f.subsidiary"><option value=""></option>@for (x of subsidiaries(); track x) { <option [value]="x">{{ x }}</option> }</select>
            </app-field>
          }
          <app-field label="Payor">
            <select name="payor" [(ngModel)]="f.payor" (ngModelChange)="loadClients()"><option value=""></option>@for (p of payors(); track p.code) { <option [value]="p.code + ' - ' + p.name">{{ p.code }} - {{ p.name }}</option> }</select>
          </app-field>
          <app-field label="Insured Client">
            <select name="insuredClient" [(ngModel)]="f.insuredClient"><option value=""></option>@for (x of clients(); track x) { <option [value]="x">{{ x }}</option> }</select>
          </app-field>
          <app-field label="Insured">
            <select name="insured" [(ngModel)]="f.insured"><option value=""></option><option value="YES">YES</option><option value="NO">NO</option></select>
          </app-field>
          <app-field label="Bordx Status">
            <select name="bordxStatus" [(ngModel)]="f.bordxStatus">@for (p of statuses; track $index) { <option [ngValue]="$index">{{ $index }} - {{ p }}</option> }</select>
          </app-field>
          <app-field label="WS Status">
            <select name="wsStatus" [(ngModel)]="f.wsStatus"><option [ngValue]="0">0 - All</option><option [ngValue]="1">1 - Paid</option><option [ngValue]="2">2 - Unpaid</option></select>
          </app-field>
          <app-field label="Receipt No From"><input name="receiptFrom" maxlength="30" [(ngModel)]="f.receiptFrom" /></app-field>
          <app-field label="Receipt No To"><input name="receiptTo" maxlength="30" [(ngModel)]="f.receiptTo" /></app-field>
          <app-field label="Receipt Date From"><app-date-input name="receiptDateFrom" ariaLabel="Receipt Date From" [(ngModel)]="f.receiptDateFrom" /></app-field>
          <app-field label="Receipt Date To"><app-date-input name="receiptDateTo" ariaLabel="Receipt Date To" [(ngModel)]="f.receiptDateTo" /></app-field>
          <app-field label="Client PV No From"><input name="clientPvFrom" maxlength="30" [(ngModel)]="f.clientPvFrom" /></app-field>
          <app-field label="Client PV No To"><input name="clientPvTo" maxlength="30" [(ngModel)]="f.clientPvTo" /></app-field>
          <app-field label="Medix Receipt From"><input name="medixReceiptFrom" maxlength="30" [(ngModel)]="f.medixReceiptFrom" /></app-field>
          <app-field label="Medix Receipt To"><input name="medixReceiptTo" maxlength="30" [(ngModel)]="f.medixReceiptTo" /></app-field>
          <app-field label="A/C Month From"><input name="monthFrom" maxlength="2" [(ngModel)]="f.monthFrom" /></app-field>
          <app-field label="A/C Month To"><input name="monthTo" maxlength="2" [(ngModel)]="f.monthTo" /></app-field>
          <app-field label="A/C Year From"><input name="yearFrom" maxlength="4" [(ngModel)]="f.yearFrom" /></app-field>
          <app-field label="A/C Year To"><input name="yearTo" maxlength="4" [(ngModel)]="f.yearTo" /></app-field>
          <app-action-bar label="Receipt search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="export()">Export</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>

        <fieldset aria-label="Receipt">
          <legend>Receipt</legend>
          <app-field label="Receipt No" [required]="true"><input name="receiptNo" maxlength="30" [(ngModel)]="r.receiptNo" /></app-field>
          <app-field label="Receipt Date" [required]="true"><app-date-input name="receiptDate" ariaLabel="Receipt Date" [(ngModel)]="r.receiptDate" /></app-field>
          <app-field label="Cheque No" [required]="true"><input name="chequeNo" maxlength="30" [(ngModel)]="r.chequeNo" /></app-field>
          <app-field label="Receipt Amount" [required]="true"><input class="num" name="receiptAmt" maxlength="20" [(ngModel)]="r.receiptAmt" /></app-field>
          <app-field label="A/C Period Month"><input name="month" maxlength="2" [(ngModel)]="r.month" /></app-field>
          <app-field label="A/C Period Year"><input name="year" maxlength="4" [(ngModel)]="r.year" /></app-field>
          <app-field label="Remark" [wide]="true"><input name="remark" maxlength="500" [(ngModel)]="r.remark" /></app-field>
          @if (!f.worksheet) {
            <app-field label="OP Claim Amt"><input class="num" name="opClaimAmt" maxlength="20" [(ngModel)]="r.opClaimAmt" /></app-field>
            <app-field label="OP Claim Desc"><input name="opClaimDesc" maxlength="500" [(ngModel)]="r.opClaimDesc" /></app-field>
            <app-field label="OP Fee Amt"><input class="num" name="opFeeAmt" maxlength="20" [(ngModel)]="r.opFeeAmt" /></app-field>
            <app-field label="OP Fee Desc"><input name="opFeeDesc" maxlength="500" [(ngModel)]="r.opFeeDesc" /></app-field>
          }
          <app-action-bar label="Receipt actions">
            @if (editing()) {
              <button appButton variant="primary" (click)="update()">Update</button>
            } @else {
              <button appButton variant="primary" [busy]="saving()" (click)="save()">Save</button>
              <button appButton (click)="editing.set(true)">Edit</button>
            }
          </app-action-bar>
        </fieldset>

        <app-data-grid [caption]="f.worksheet ? 'Worksheet receipts' : 'Bordx receipts'" [columns]="f.worksheet ? wksColumns : bordxColumns" [rows]="rows()"
          [rowKey]="key" [checkable]="true" [checked]="checked()" (checkedChange)="check($event)" [selectable]="true" [selectedKey]="selectedKey()"
          (rowSelected)="select($event)" />
        <p class="totals" aria-label="Totals">Total: <strong>{{ rows().length }}</strong> · Selected amount: <strong>{{ money(selected()) }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class ReceiptPayorPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly statuses = ['All', 'Paid', 'Unpaid', 'Partial Paid', 'Paid & Partial Paid', 'Unpaid & Partial Paid']
  protected readonly companies = signal<string[]>([])
  protected readonly payors = signal<{ code: string; name: string }[]>([])
  protected readonly clients = signal<string[]>([])
  protected readonly subsidiaries = signal<string[]>([])
  protected readonly rows = signal<Row[]>([])
  protected readonly checked = signal<ReadonlySet<string>>(new Set())
  protected readonly selectedKey = signal<string | null>(null)
  protected readonly busy = signal(false)
  protected readonly saving = signal(false)
  protected readonly editing = signal(false)
  protected readonly selected = computed(() =>
    this.rows().filter((x) => this.checked().has(this.key(x))).reduce((t, x) => t + Number((s(x['BordxPAYFRStatus']) === 'P' ? x['Balance'] : x['TickAmount']) ?? 0), 0),
  )
  protected f = this.empty()
  protected r = this.emptyReceipt()
  private selectedRow: Row | null = null
  protected readonly key = (x: Row) => `${s(x['BordxRefNo'])}|${s(x['CASNumber'])}|${s(x['ReceiptNo'])}`
  protected readonly bordxColumns: GridColumn<Row>[] = [
    { key: 'BordxPAYCode', label: 'Payor' }, { key: 'BordxRefNo', label: 'Bordx No' }, money('BordxClaimAmt', 'Claim Amt'), money('CheckAmt', 'Received'),
    money('Balance', 'Balance'), { key: 'Status', label: 'Status' }, { key: 'BordxDate', label: 'Bordx Date', value: (x) => d(x['BordxDate']) },
    { key: 'Subsidiary', label: 'Subsidiary' }, { key: 'GroupCompany', label: 'Group Company' }, { key: 'ReceiptNo', label: 'Receipt No' },
    { key: 'ClientPVNo', label: 'Client PV No' }, { key: 'MedixReceiptNo', label: 'Medix Receipt No' }, { key: 'isInsured', label: 'Insured' },
  ]
  protected readonly wksColumns: GridColumn<Row>[] = [
    { key: 'BordxPAYCode', label: 'Payor' }, { key: 'BordxRefNo', label: 'Bordx No' }, { key: 'CASNumber', label: 'Case No' }, money('CLMBordxAmt', 'Bordx Amt'),
    money('CheckAmt', 'Received'), money('Balance', 'Balance'), { key: 'ReceiptNo', label: 'Receipt No' },
    { key: 'ReceiptDate', label: 'Receipt Date', value: (x) => d(x['ReceiptDate']) }, { key: 'Status', label: 'Status' },
    { key: 'BordxDate', label: 'Bordx Date', value: (x) => d(x['BordxDate']) }, { key: 'Subsidiary', label: 'Subsidiary' }, { key: 'GroupCompany', label: 'Group Company' },
    { key: 'ClientPVNo', label: 'Client PV No' }, { key: 'MedixReceiptNo', label: 'Medix Receipt No' }, { key: 'isInsured', label: 'Insured' },
  ]

  private empty() {
    return {
      worksheet: false, scan: true, bordxFrom: '', bordxTo: '', caseFrom: '', caseTo: '', company: '', subsidiary: '', payor: '', insuredClient: '', insured: '',
      bordxStatus: 0, wsStatus: 0, receiptFrom: '', receiptTo: '', receiptDateFrom: '', receiptDateTo: '', clientPvFrom: '', clientPvTo: '', medixReceiptFrom: '',
      medixReceiptTo: '', monthFrom: '', monthTo: '', yearFrom: '', yearTo: '',
    }
  }

  private emptyReceipt() {
    return { receiptNo: '', receiptDate: '', chequeNo: '', receiptAmt: '', month: '', year: '', remark: '', opClaimAmt: '', opClaimDesc: '', opFeeAmt: '', opFeeDesc: '' }
  }

  async ngOnInit(): Promise<void> {
    try {
      const l = await firstValueFrom(this.http.get<{ companies: string[]; payors: { code: string; name: string }[] }>(`${BASE}/lookups`))
      this.companies.set(l.companies)
      this.payors.set(l.payors)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** OptBordx_Click / OptWks_Click. */
  protected mode(): void {
    this.f.scan = true
    this.r = this.emptyReceipt()
    this.editing.set(false)
    this.rows.set([])
    this.checked.set(new Set())
  }

  protected async loadSubsidiaries(): Promise<void> {
    this.subsidiaries.set(this.f.company.trim() ? await firstValueFrom(this.http.get<string[]>(`${BASE}/subsidiaries`, { params: { company: this.f.company.trim() } })) : [])
  }

  protected async loadClients(): Promise<void> {
    this.f.insuredClient = ''
    this.clients.set(this.f.payor ? await firstValueFrom(this.http.get<string[]>(`${BASE}/clients`, { params: { payor: this.f.payor } })) : [])
  }

  /** cmdSearch_Click: scan mode appends. */
  protected async search(): Promise<void> {
    const scanKey = this.f.worksheet ? this.f.caseFrom.trim() : this.f.bordxFrom.trim()
    if (this.f.scan && scanKey && this.rows().some((x) => s(x[this.f.worksheet ? 'CASNumber' : 'BordxRefNo']) === scanKey))
      return this.notify.notify('Data has been uploaded!', 'error')
    this.busy.set(true)
    try {
      const rows = await firstValueFrom(this.http.post<Row[]>(`${BASE}/search`, this.f))
      this.rows.set(this.f.scan ? [...this.rows(), ...rows] : rows)
      if (!this.f.scan) this.checked.set(new Set())
      if (this.f.scan && this.f.worksheet) this.f.caseFrom = ''
      else if (this.f.scan) this.f.bordxFrom = ''
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstBordx_ItemCheck / ProcessCheckedItem. */
  protected check(keys: Set<string>): void {
    const before = this.checked()
    for (const x of this.rows()) {
      const k = this.key(x)
      if (!keys.has(k) || before.has(k)) continue
      const status = s(this.f.worksheet ? x['PVRecPayorStatus'] : x['BordxPAYFRStatus'])
      if (status === 'F') {
        keys.delete(k)
        this.notify.notify('Record has been paid !', 'error')
      } else if (!this.f.worksheet && status === 'P') this.notify.notify('Record has been Partial Paid !', 'info')
    }
    this.checked.set(keys)
  }

  /** lstBordx_Click / lstWorksheet_Click: in edit mode the row's receipt fills the frame. */
  protected async select(x: Row): Promise<void> {
    this.selectedKey.set(this.key(x))
    this.selectedRow = x
    if (!this.editing()) return
    try {
      const params: Record<string, string> = { bordxNo: s(x['BordxRefNo']) }
      if (this.f.worksheet) params['casNumber'] = s(x['CASNumber'])
      const rc = await firstValueFrom(this.http.get<Row>(`${BASE}/receipt`, { params }))
      this.r = {
        ...this.emptyReceipt(), receiptNo: s(rc['ReceiptNo']), receiptDate: d(rc['ReceiptDate']), receiptAmt: s(rc['ReceiptAmt']), chequeNo: s(rc['ChequeNo']),
        remark: s(rc['Remark']), month: s(rc['RECAccPeriodMth']), year: s(rc['RECAccPeriodYR']),
      }
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  private ticked() {
    return this.rows().filter((x) => this.checked().has(this.key(x))).map((x) => ({ bordxNo: s(x['BordxRefNo']), casNumber: this.f.worksheet ? s(x['CASNumber']) : null }))
  }

  /** CmdSave_Click (SaveBordx / SaveWorksheet). */
  protected async save(): Promise<void> {
    if (!(await this.notify.ask(this.f.worksheet ? 'Do you want to update the record?' : 'Do you want to update the record ?'))) return
    this.saving.set(true)
    try {
      const receiptNo = this.r.receiptNo.trim()
      const res = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/save`, { ...this.r, worksheet: this.f.worksheet, company: this.f.company, ticked: this.ticked() }))
      this.notify.notify(res.message, 'success')
      this.r = this.emptyReceipt()
      this.checked.set(new Set())
      this.rows.set(await firstValueFrom(this.http.post<Row[]>(`${BASE}/relist`, { worksheet: this.f.worksheet }, { params: { receiptNo } })))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.saving.set(false)
    }
  }

  /** CmdUpdate_Click (UpdateBordx / UpdateWorksheet). */
  protected async update(): Promise<void> {
    if (!this.rows().length || !this.selectedRow) return this.notify.notify('Please select the record', 'error')
    if (!(await this.notify.ask('Do you want to update the record?'))) return
    try {
      const selected = { bordxNo: s(this.selectedRow['BordxRefNo']), casNumber: this.f.worksheet ? s(this.selectedRow['CASNumber']) : null }
      const res = await firstValueFrom(this.http.post<{ message: string }>(`${BASE}/update`, { ...this.r, worksheet: this.f.worksheet, selected, ticked: [] }))
      this.notify.notify(res.message, 'success')
      this.editing.set(false)
      this.r = this.emptyReceipt()
      if (!this.f.scan) await this.search()
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected export(): void {
    if (!this.rows().length) return this.notify.notify('Report cannot be printed without any record.', 'info')
    const cols = this.f.worksheet ? this.wksColumns : this.bordxColumns
    saveCsv([cols.map((c) => c.label), ...this.rows().map((x) => cols.map((c) => (c.value ? c.value(x) : s(x[c.key]))))], 'ReceiptPayor.csv')
  }

  protected clear(): void {
    const worksheet = this.f.worksheet
    this.f = { ...this.empty(), worksheet }
    this.r = this.emptyReceipt()
    this.editing.set(false)
    this.rows.set([])
    this.checked.set(new Set())
  }
}
