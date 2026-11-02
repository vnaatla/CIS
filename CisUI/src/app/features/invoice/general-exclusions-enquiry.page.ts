import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { ActivatedRoute } from '@angular/router'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { FieldComponent } from '../../shared/field.component'
import { openBlob } from '../../shared/lib/download'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'

const BASE = '/api/invoice/general-exclusions'

/** The 24 check boxes in form order with the desktop's captions. */
const BOXES: [string, string][] = [
  ['Dermatitis', 'Dermatitis/Skin Disorder'], ['Gynaecological', 'Gynaecological Disorder'], ['Dental', 'Dental'], ['Optical', 'Optical'],
  ['MedCheckup', 'Medical Check-up'], ['PreEmpCheckup', 'Pre-Employment Check-up'], ['Hypertension', 'Hypertension'], ['Diabetes', 'Diabetes'],
  ['IHDCAD', 'IHD/CAD'], ['Thyroid', 'Tyroid Disorder'], ['Immunisation', 'Immunisation'], ['HRT', 'HRT'], ['PregnacyTest', 'Prenancy Related/Test'],
  ['Antenatal', 'Antenatal/Postnatal'], ['Abortion', 'Abortion/Miscarriage'], ['Infertility', 'Infertility'], ['Contraception', 'Contraception'],
  ['Genitourinary', 'Genitourinary Disorder'], ['Vitamins', 'Vitamin purchase'], ['BloodTest', 'Blood Test'], ['XRAY', 'X-Ray/MRI'], ['PapSmear', 'Pap-Smear'],
  ['Typhoid', 'Typhoid Vaccine'], ['LTMedication', 'Long-Term Medication'],
]

/** Invoice > General Exclusions (Enquiry) (frmGeneralBenefitExcEnq): read-only exclusions and the policy wording files of a group company. */
@Component({
  selector: 'app-general-exclusions-enquiry-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, TabsComponent, TabDirective],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="General Exclusions (Enquiry)" subtitle="Replaces frmGeneralBenefitExcEnq">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Group Company" [required]="true">
            <input name="groupCompany" list="ge-groups" data-autofocus [(ngModel)]="groupCompany" (keydown.enter)="search()" />
            <datalist id="ge-groups">
              @for (g of groups(); track $index) {
                <option [value]="g"></option>
              }
            </datalist>
          </app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
          </app-action-bar>
        </fieldset>

        <app-tabs label="General exclusions" [(active)]="tab">
          <ng-template appTab="Remarks">
            <fieldset aria-label="Exclusions">
              <legend>Exclusions</legend>
              @for (b of boxes; track b[0]) {
                <label class="check-field"><input type="checkbox" [attr.name]="b[0]" [checked]="flags()[b[0]]" disabled /> {{ b[1] }}</label>
              }
              <app-field label="Remarks" [wide]="true">
                <textarea name="remarks" rows="6" [value]="remarks()" disabled></textarea>
              </app-field>
            </fieldset>
          </ng-template>
          <ng-template appTab="Benifit Schedule">
            @if (!searched()) {
              <p class="muted">Search a group company first.</p>
            } @else if (!configured()) {
              <p class="muted">The policy wording folder (CLNPolicyWording) is not configured on the API server (Files:Folders:PolicyWording).</p>
            } @else if (files().length === 0) {
              <p class="muted">No policy wording files for {{ searched() }}.</p>
            } @else {
              <ul class="file-list" aria-label="Policy wording files">
                @for (f of files(); track f) {
                  <li><button type="button" class="link" (click)="open(f)">{{ f }}</button></li>
                }
              </ul>
            }
          </ng-template>
        </app-tabs>
      </div>
    </app-page-shell>
  `,
})
export class GeneralExclusionsEnquiryPage implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  private readonly route = inject(ActivatedRoute)

  protected readonly boxes = BOXES
  protected readonly groups = signal<string[]>([])
  protected readonly flags = signal<Record<string, boolean>>({})
  protected readonly remarks = signal('')
  protected readonly files = signal<string[]>([])
  protected readonly configured = signal(true)
  protected readonly searched = signal('')
  protected readonly busy = signal(false)
  protected groupCompany = ''
  protected tab = 0

  /** MBM Enquiry's Exclusions button opens the page with ?groupCompany= and searches it (cmdExclusions_Click). */
  async ngOnInit(): Promise<void> {
    const preset = this.route.snapshot.queryParamMap.get('groupCompany')
    if (preset) {
      this.groupCompany = preset
      void this.search()
    }
    try {
      this.groups.set((await firstValueFrom(this.http.get<string[] | null>(`${BASE}/group-companies`))) ?? [])
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    }
  }

  /** cmdSearch_Click: boxes cleared, Remarks tab, CallExclusions; the Benifit Schedule tab lists the company's files. */
  protected async search(): Promise<void> {
    this.flags.set({})
    this.remarks.set('')
    this.files.set([])
    this.searched.set('')
    this.tab = 0
    this.busy.set(true)
    try {
      const params = new HttpParams({ fromObject: { groupCompany: this.groupCompany } })
      const v = await firstValueFrom(this.http.get<{ groupCompany: string; flags: Record<string, boolean>; remarks: string }>(BASE, { params }))
      this.flags.set(v.flags)
      this.remarks.set(v.remarks)
      this.searched.set(v.groupCompany)
      const pw = await firstValueFrom(this.http.get<{ configured: boolean; files: string[] }>(`${BASE}/policy-wording`, { params }))
      this.configured.set(pw.configured)
      this.files.set(pw.files)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstPolicyWording_Click: opens the file. */
  protected async open(fileName: string): Promise<void> {
    try {
      const params = new HttpParams({ fromObject: { groupCompany: this.searched(), fileName } })
      openBlob(await firstValueFrom(this.http.get(`${BASE}/policy-wording/file`, { params, responseType: 'blob' })))
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'error')
    }
  }
}
