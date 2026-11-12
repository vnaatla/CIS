import { HttpClient, HttpResponse } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { ActivatedRoute } from '@angular/router'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { formatMoney } from '../../shared/lib/format'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

type Format = 'bintulu' | 'toshiba' | 'singer' | 'tims' | 'cimb' | 'omni' | 'utusan'
const SCREENS: Record<Format, { title: string; form: string; range: boolean; clinic: boolean; confirm: boolean }> = {
  bintulu: { title: 'Claim ASCII (Bintulu)', form: 'frmClaimDetailsASCII', range: true, clinic: true, confirm: true },
  toshiba: { title: 'Claim ASCII (Toshiba)', form: 'frmClaimASCIIToshiba', range: true, clinic: true, confirm: true },
  singer: { title: 'Claim ASCII (Singer)', form: 'frmClaimASCIISinger', range: true, clinic: false, confirm: true },
  tims: { title: 'Claim ASCII (TIMS)', form: 'frmClaimASCIITIMS', range: false, clinic: false, confirm: false },
  cimb: { title: 'Claim ASCII (CIMB)', form: 'frmClaimASCIICIMB', range: false, clinic: false, confirm: false },
  omni: { title: 'Claim ASCII (OMNI)', form: 'frmClaimDataOMNI', range: false, clinic: false, confirm: false },
  utusan: { title: 'Claim ASCII (Utusan Melayu)', form: 'frmClaimASCIIUtusan', range: true, clinic: false, confirm: true },
}

/** Claim ASCII exports (route data `format`). */
@Component({
  selector: 'app-claim-ascii-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="screen.title" [subtitle]="'Replaces ' + screen.form">
      <div class="screen-form">
        <fieldset aria-label="Generate">
          <legend>Generate</legend>
          <app-field [label]="screen.range ? 'Bordx No From' : 'Bordx No'"><input name="bordxFrom" maxlength="20" [(ngModel)]="f.bordxFrom" /></app-field>
          @if (screen.range) {
            <app-field label="Bordx No To"><input name="bordxTo" maxlength="20" [(ngModel)]="f.bordxTo" /></app-field>
          }
          @if (screen.clinic) {
            <app-field label="Clinic" [wide]="true">
              <input name="clinic" list="ascii-clinics" placeholder="code - name" [(ngModel)]="f.clinic" />
              <datalist id="ascii-clinics">@for (c of clinics(); track $index) { <option [value]="c"></option> }</datalist>
            </app-field>
          }
          @if (format === 'cimb') {
            <app-field label="Cash Type">
              <select name="cashType" [(ngModel)]="f.cashType"><option>C - Cashless</option><option>R - Reimbursement</option></select>
            </app-field>
          }
          <app-action-bar label="Export actions">
            <button appButton variant="primary" [busy]="busy()" (click)="generate()">Generate</button>
          </app-action-bar>
        </fieldset>
        <p class="totals" aria-label="Totals">Total Visits: <strong>{{ visits() }}</strong> · Total Amount: <strong>{{ money(total()) }}</strong></p>
      </div>
    </app-page-shell>
  `,
})
export class ClaimAsciiPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  protected readonly format = inject(ActivatedRoute).snapshot.data['format'] as Format
  protected readonly screen = SCREENS[this.format]
  private readonly base = `/api/reports/claim-ascii/${this.format}`

  protected readonly money = (v: number) => formatMoney(v)
  protected readonly clinics = signal<string[]>([])
  protected readonly visits = signal(0)
  protected readonly total = signal(0)
  protected readonly busy = signal(false)
  protected f = { bordxFrom: '', bordxTo: '', clinic: '', cashType: 'C - Cashless' }

  async ngOnInit(): Promise<void> {
    if (!this.screen.clinic) return
    try {
      this.clinics.set(await firstValueFrom(this.http.get<string[]>(`${this.base}/clinics`)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  protected async generate(): Promise<void> {
    if (this.screen.confirm && !(await this.notify.ask('Are you sure?'))) return
    this.busy.set(true)
    this.visits.set(0)
    this.total.set(0)
    try {
      const res: HttpResponse<Blob> = await firstValueFrom(this.http.post(this.base, this.f, { responseType: 'blob', observe: 'response' }))
      const name = /filename\*?=(?:UTF-8'')?"?([^";]+)"?/i.exec(res.headers.get('Content-Disposition') ?? '')?.[1]
      saveBlob(res.body!, decodeURIComponent(name ?? `${this.format}.txt`))
      this.visits.set(Number(res.headers.get('X-Total-Visits') ?? 0))
      this.total.set(Number(res.headers.get('X-Total-Amount') ?? 0))
      this.notify.notify('File Generated Sucessfully...', 'success')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }
}
