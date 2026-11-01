import { HttpClient, HttpParams } from '@angular/common/http'
import { ChangeDetectionStrategy, Component, OnInit, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { firstValueFrom } from 'rxjs'
import { errorMessage } from '../../core/http/errors'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { DataGridComponent, GridColumn } from '../../shared/data-grid.component'
import { FieldComponent } from '../../shared/field.component'
import { ActionBarComponent, StatComponent } from '../../shared/screen.components'

interface Binding {
  id: number
  groupCompany: string | null
  payCode: string | null
  planCode: string | null
  glType: string | null
  glTemplate: string | null
  lgFormat: string | null
  vipInd: string | null
  depInd: string | null
  depTemplate: string | null
  tmpRemarks: string | null
}
interface Lookups {
  groupCompanies: LookupOption[]
  glTypes: LookupOption[]
  templates: LookupOption[]
  lgFormats: LookupOption[]
}
interface Form {
  groupCode: string
  planCode: string
  payCode: string
  glType: string
  template: string
  dependentTemplate: string
  lgFormat: string
  vip: string
  remarks: string
}
const EMPTY: Form = { groupCode: '', planCode: '', payCode: '', glType: '', template: '', dependentTemplate: '', lgFormat: '', vip: '', remarks: '' }

/** frmLGMaintenance tab 0 "GL Template Binding" (Tbl_OPGLMaster): group company → plans / pay codes, GL type, templates, LG format, VIP. */
@Component({
  selector: 'app-lg-binding-panel',
  imports: [FormsModule, FieldComponent, ButtonDirective, DataGridComponent, ActionBarComponent, StatComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="screen-form">
      <fieldset aria-label="GL Template Binding">
        <legend>GL Template Binding</legend>
        <app-field label="Group Company" [required]="true" hint="Searchable">
          <select name="groupCode" [ngModel]="form().groupCode" (ngModelChange)="pickGroup($event)">
            <option value=""></option>
            @for (o of lookups().groupCompanies; track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="Pay Code" [required]="true" hint="Searchable">
          <select name="payCode" [ngModel]="form().payCode" (ngModelChange)="patch({ payCode: $event })">
            <option value=""></option>
            @for (o of payCodes(); track o.code) {
              <option [value]="o.code">{{ o.code }}</option>
            }
          </select>
        </app-field>
        <app-field label="Plan Code" [required]="true" hint="Searchable">
          <select name="planCode" [ngModel]="form().planCode" (ngModelChange)="patch({ planCode: $event })">
            <option value=""></option>
            @for (o of plans(); track o.code) {
              <option [value]="o.code">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="GLType" [required]="true">
          <select name="glType" [ngModel]="form().glType" (ngModelChange)="patch({ glType: $event })">
            <option value=""></option>
            @for (o of lookups().glTypes; track o.code) {
              <option [value]="o.name">{{ o.code }} - {{ o.name }}</option>
            }
          </select>
        </app-field>
        <app-field label="LG Template Name">
          <select name="template" [ngModel]="form().template" (ngModelChange)="pickTemplate($event)">
            <option value=""></option>
            @for (o of lookups().templates; track o.code) {
              <option [value]="o.code">{{ o.code }}</option>
            }
          </select>
        </app-field>
        <app-field label="Dependent Template">
          <select name="dependentTemplate" [ngModel]="form().dependentTemplate" (ngModelChange)="patch({ dependentTemplate: $event })">
            <option value=""></option>
            @for (o of lookups().templates; track o.code) {
              <option [value]="o.code">{{ o.code }}</option>
            }
          </select>
        </app-field>
        <app-field label="LG Format" [required]="true">
          <select name="lgFormat" [ngModel]="form().lgFormat" (ngModelChange)="patch({ lgFormat: $event })">
            <option value=""></option>
            @for (o of lookups().lgFormats; track o.code) {
              <option [value]="o.code">{{ o.code }}</option>
            }
          </select>
        </app-field>
        <app-field label="VIP Indicator">
          <select name="vip" [ngModel]="form().vip" (ngModelChange)="patch({ vip: $event })">
            <option value=""></option>
            <option value="N">N-No</option>
            <option value="Y">Y-Yes</option>
          </select>
        </app-field>
        <app-field label="Template Remarks" [wide]="true">
          <textarea name="remarks" rows="3" maxlength="2000" [ngModel]="form().remarks" (ngModelChange)="patch({ remarks: $event })"></textarea>
        </app-field>
      </fieldset>

      <app-action-bar label="GL Template Binding actions">
        <button appButton variant="primary" [busy]="busy()" (click)="search()">Search</button>
        <button appButton (click)="add()">Add</button>
        <button appButton [disabled]="!selected()" (click)="editing.set(true); saving.set(true)">Edit</button>
        <button appButton [disabled]="!saving()" [busy]="busy()" (click)="save()">Save</button>
        <button appButton variant="ghost" (click)="clear()">Clear</button>
        @if (lookupsFailed()) {
          <button appButton variant="ghost" (click)="loadLookups()">Reload lists</button>
        }
        <span class="spacer"></span>
        <app-stat label="TOTAL RECORDS" [value]="rows().length" />
      </app-action-bar>

      <app-data-grid caption="GL template bindings" [columns]="columns" [rows]="rows()" [rowKey]="rowKey" [selectable]="true"
        [selectedKey]="selected() ? '' + selected()!.id : null" (rowSelected)="select($event)" [tall]="true" />
    </div>
  `,
})
export class LgBindingPanelComponent implements OnInit {
  private readonly http = inject(HttpClient)
  private readonly notify = inject(NotifyService)
  private readonly base = '/api/maintenance/lg-binding'

  protected readonly lookups = signal<Lookups>({ groupCompanies: [], glTypes: [], templates: [], lgFormats: [] })
  protected readonly plans = signal<LookupOption[]>([])
  protected readonly payCodes = signal<LookupOption[]>([])
  protected readonly form = signal<Form>(EMPTY)
  protected readonly rows = signal<Binding[]>([])
  protected readonly selected = signal<Binding | null>(null)
  protected readonly editing = signal(false)
  protected readonly saving = signal(false)
  protected readonly busy = signal(false)
  private lastGroupName = ''

  protected readonly rowKey = (r: Binding) => String(r.id)
  protected readonly columns: GridColumn<Binding>[] = [
    { key: 'id', label: 'ID', align: 'right' },
    { key: 'groupCompany', label: 'Group Company' },
    { key: 'payCode', label: 'Pay Code' },
    { key: 'planCode', label: 'Plan Code' },
    { key: 'glType', label: 'GL Type' },
    { key: 'glTemplate', label: 'GL Template' },
    { key: 'lgFormat', label: 'LG Format' },
    { key: 'vipInd', label: 'VIP' },
  ]

  async ngOnInit(): Promise<void> {
    await this.loadLookups()
  }

  /** Form_Load binds. Retried once, then again from the Reload lists button, so a slow database does not leave empty drop-downs. */
  protected async loadLookups(): Promise<void> {
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        this.lookups.set(await firstValueFrom(this.http.get<Lookups>(`${this.base}/lookups`)))
        this.lookupsFailed.set(false)
        return
      } catch (e) {
        if (attempt === 2) {
          this.lookupsFailed.set(true)
          this.notify.notify(`The drop-down lists could not be loaded: ${errorMessage(e)}`, 'error')
        }
      }
    }
  }

  protected readonly lookupsFailed = signal(false)

  protected patch(p: Partial<Form>): void {
    this.form.update((f) => ({ ...f, ...p }))
  }

  /** cmbGrpCompany_LostFocus → BindPlanCode / BindPayCode. */
  protected async pickGroup(code: string): Promise<void> {
    this.patch({ groupCode: code, planCode: '', payCode: '' })
    this.plans.set([])
    this.payCodes.set([])
    if (!code) return
    const p = await firstValueFrom(this.http.get<{ plans: LookupOption[]; payCodes: LookupOption[] }>(`${this.base}/groups/${encodeURIComponent(code)}/plans`))
    this.plans.set(p.plans)
    this.payCodes.set(p.payCodes)
  }

  /** CmbTemplateName_Click: remarks of the template. */
  protected pickTemplate(name: string): void {
    this.patch({ template: name, remarks: this.lookups().templates.find((t) => t.code === name)?.name ?? '' })
  }

  /** What the desktop stores / searches as GRPCompnay: Split("CODE - NAME", "-")(1), so a name with "-" is cut at it. */
  private static legacyName(g: LookupOption): string {
    return `${g.code} - ${g.name}`.split('-')[1]?.trim() ?? ''
  }

  private groupName(): string {
    const code = this.form().groupCode
    const group = this.lookups().groupCompanies.find((g) => g.code === code)
    return group ? LgBindingPanelComponent.legacyName(group) : ''
  }

  protected async search(): Promise<void> {
    await this.load(this.groupName() || this.lastGroupName)
  }

  private async load(groupName: string): Promise<void> {
    this.busy.set(true)
    try {
      const params = new HttpParams({ fromObject: { groupName, payCode: this.form().payCode, planCode: this.form().planCode } })
      this.rows.set(await firstValueFrom(this.http.get<Binding[]>(this.base, { params })))
      this.lastGroupName = groupName
      this.selected.set(null)
      this.editing.set(false)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'warn')
    } finally {
      this.busy.set(false)
    }
  }

  /** lstLGTemplate_Click: the row's values back into the combos. */
  protected async select(r: Binding): Promise<void> {
    this.selected.set(r)
    const group = this.lookups().groupCompanies.find((g) => LgBindingPanelComponent.legacyName(g) === (r.groupCompany ?? '').trim())
    if (group) await this.pickGroup(group.code)
    this.form.set({
      groupCode: group?.code ?? '', planCode: r.planCode ?? '', payCode: r.payCode ?? '', glType: r.glType ?? '', template: r.glTemplate ?? '',
      dependentTemplate: r.depTemplate ?? '', lgFormat: r.lgFormat ?? '', vip: r.vipInd ?? '', remarks: r.tmpRemarks ?? '',
    })
  }

  protected add(): void {
    this.clear()
    this.saving.set(true)
  }

  /** Clear: fields and the grid (the desktop's LG Clear kept the list; emptied here like the other Maintenance screens). */
  protected clear(): void {
    this.form.set(EMPTY)
    this.plans.set([])
    this.payCodes.set([])
    this.rows.set([])
    this.selected.set(null)
    this.editing.set(false)
    this.saving.set(false)
  }

  /** cmdSave_Click → SaveGLTemplate, then Clear and Search again on the same group company (SearchGrpComp). */
  protected async save(): Promise<void> {
    const f = this.form()
    const body = {
      groupName: this.groupName(), planCode: f.planCode, glType: f.glType, template: f.template, payCode: f.payCode, lgFormat: f.lgFormat,
      vip: f.vip, dependentTemplate: f.dependentTemplate, remarks: f.remarks,
    }
    this.busy.set(true)
    try {
      const answer = this.editing()
        ? await firstValueFrom(this.http.put<{ message: string }>(this.base, body))
        : await firstValueFrom(this.http.post<{ message: string }>(this.base, body))
      this.notify.notify(answer.message, 'success')
      const group = body.groupName
      this.clear()
      await this.load(group)
    } catch (e) {
      this.notify.notify(errorMessage(e), 'error')
    } finally {
      this.busy.set(false)
    }
  }
}
