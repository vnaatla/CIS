import { HttpClient } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { blobErrorMessage, errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { saveBlob } from '../../shared/lib/download'
import { ModalComponent } from '../../shared/modal.component'
import { ActionBarComponent, PageShellComponent } from '../../shared/screen.components'

const BASE = '/api/invoice/mma-cpt'

interface MmaRow {
  mainBp: string
  opcs: string
  narrative: string
  surgCat: string
  anaesCat: string
  surgeonFee: string
  anaesFee: string
  partAFee: string
  page: string
  part: string
  cptCode: string
  cptDesc: string
}

/** Invoice > Other > MMA/CPT (frmMMASchedule): search the MMA schedule; a row opens its details (frmShowInvestigation). */
@Component({
  selector: 'app-mma-cpt-page',
  imports: [FormsModule, PageShellComponent, FieldComponent, ButtonDirective, ActionBarComponent, DataGridComponent, ModalComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="MMA Schedule / CPT" subtitle="Replaces frmMMASchedule, frmShowInvestigation">
      <div class="screen-form">
        <fieldset aria-label="Search">
          <legend>Search</legend>
          <app-field label="Main Body Part"><input name="mainBp" data-autofocus [(ngModel)]="f.mainBp" (keydown.enter)="search()" /></app-field>
          <app-field label="OPCS Code"><input name="opcs" [(ngModel)]="f.opcs" (keydown.enter)="search()" /></app-field>
          <app-field label="Procedures"><input name="procedure" [(ngModel)]="f.procedure" (keydown.enter)="search()" /></app-field>
          <app-field label="Surgical Category"><input name="surgCat" [(ngModel)]="f.surgCat" /></app-field>
          <app-field label="Anaesthetic Category"><input name="anaesCat" [(ngModel)]="f.anaesCat" /></app-field>
          <app-field label="MMA Part"><input name="part" maxlength="10" [(ngModel)]="f.part" /></app-field>
          <app-field label="Page From"><input name="pageFrom" [(ngModel)]="f.pageFrom" /></app-field>
          <app-field label="Page To"><input name="pageTo" [(ngModel)]="f.pageTo" /></app-field>
          <app-field label="CPT Code"><input name="cptCode" [(ngModel)]="f.cptCode" /></app-field>
          <app-field label="CPT Description"><input name="cptDesc" [(ngModel)]="f.cptDesc" /></app-field>
          <app-action-bar label="Search actions">
            <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
            <button appButton (click)="print()">Print</button>
            <button appButton (click)="clear()">Clear</button>
          </app-action-bar>
        </fieldset>
        <app-data-grid caption="MMA schedule" [columns]="columns" [rows]="rows()" [rowKey]="key" [selectable]="true" (rowSelected)="detail.set($event)" />
      </div>
    </app-page-shell>

    @if (detail(); as d) {
      <app-modal title="MMA Details" (closed)="detail.set(null)">
        <div class="screen-form">
          <fieldset aria-label="MMA details">
            <legend>{{ d.opcs }}</legend>
            <app-field label="Body Part"><input name="dBodyPart" [value]="d.mainBp" disabled /></app-field>
            <app-field label="OPCS Code"><input name="dOpcs" [value]="d.opcs" disabled /></app-field>
            <app-field label="Surgical Category"><input name="dSurgCat" [value]="d.surgCat" disabled /></app-field>
            <app-field label="Surgical Fee"><input class="num" name="dSurgFee" [value]="d.surgeonFee" disabled /></app-field>
            <app-field label="Anaesthetic Category"><input name="dAnaesCat" [value]="d.anaesCat" disabled /></app-field>
            <app-field label="Anaesthetic Fee"><input class="num" name="dAnaesFee" [value]="d.anaesFee" disabled /></app-field>
            <app-field label="Part A Fee"><input class="num" name="dPartA" [value]="d.partAFee" disabled /></app-field>
            <app-field label="Page"><input name="dPage" [value]="d.page" disabled /></app-field>
            <app-field label="MMA Part"><input name="dPart" [value]="d.part" disabled /></app-field>
            <app-field label="Description" [wide]="true"><textarea name="dDesc" rows="3" [value]="d.narrative" disabled></textarea></app-field>
          </fieldset>
        </div>
      </app-modal>
    }
  `,
})
export class MmaCptPage {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)

  protected readonly rows = signal<MmaRow[]>([])
  protected readonly detail = signal<MmaRow | null>(null)
  protected readonly busy = signal(false)
  protected readonly key = (r: MmaRow) => `${r.opcs}|${r.cptCode}|${r.page}|${r.part}|${r.narrative}`
  protected f = this.empty()

  protected readonly columns: GridColumn<MmaRow>[] = [
    { key: 'mainBp', label: 'Main Body Part' }, { key: 'opcs', label: 'OPCS Code' }, { key: 'narrative', label: 'Procedures' },
    { key: 'surgCat', label: 'Surgical Cat' }, { key: 'anaesCat', label: 'Anaesthetic Cat' }, { key: 'surgeonFee', label: 'Surgeon Fee', align: 'right' },
    { key: 'anaesFee', label: 'Anaesthetist Fee', align: 'right' }, { key: 'partAFee', label: 'Part A Fee', align: 'right' }, { key: 'page', label: 'Page' },
    { key: 'part', label: 'Part' }, { key: 'cptCode', label: 'CPT Code' }, { key: 'cptDesc', label: 'CPT Description' },
  ]

  protected async search(): Promise<void> {
    this.busy.set(true)
    this.rows.set([])
    try {
      this.rows.set(await firstValueFrom(this.http.post<MmaRow[]>(`${BASE}/search`, this.f)))
    } catch (e) {
      this.notify.notify(errorMessage(e), 'info')
    } finally {
      this.busy.set(false)
    }
  }

  protected async print(): Promise<void> {
    try {
      saveBlob(await firstValueFrom(this.http.post(`${BASE}/export`, this.f, { responseType: 'blob' })), 'MMASchedule.xlsx')
    } catch (e) {
      this.notify.notify(await blobErrorMessage(e), 'info')
    }
  }

  protected clear(): void {
    this.f = this.empty()
    this.rows.set([])
  }

  private empty() {
    return { mainBp: '', opcs: '', procedure: '', surgCat: '', anaesCat: '', part: '', pageFrom: '', pageTo: '', cptCode: '', cptDesc: '' }
  }
}
