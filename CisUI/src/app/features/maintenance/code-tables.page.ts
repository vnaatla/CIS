import { ChangeDetectionStrategy, Component, input } from '@angular/core'
import { PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { CodeTablePanelComponent } from './code-table-panel.component'

export interface CodeTableTab {
  label: string
  key: string
}

/**
 * A maintenance screen made of code-table tabs (route data: title, legacyForm, tabs). One tab = no tab strip (frmCheckingItems);
 * several = the desktop's SSTab (frmMaintenanceCal, frmMaintenanceOthers).
 */
@Component({
  selector: 'app-code-tables-page',
  imports: [PageShellComponent, TabsComponent, TabDirective, CodeTablePanelComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell [title]="title()" [subtitle]="'Replaces ' + legacyForm()">
      @if (tabs().length === 1) {
        <app-code-table-panel [tableKey]="tabs()[0].key" />
      } @else {
        <app-tabs [label]="title()">
          @for (t of tabs(); track t.key) {
            <ng-template [appTab]="t.label"><app-code-table-panel [tableKey]="t.key" /></ng-template>
          }
        </app-tabs>
      }
    </app-page-shell>
  `,
})
export class CodeTablesPage {
  readonly title = input.required<string>()
  readonly legacyForm = input.required<string>()
  readonly tabs = input.required<CodeTableTab[]>()
}
