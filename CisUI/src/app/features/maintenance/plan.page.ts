import { ChangeDetectionStrategy, Component } from '@angular/core'
import { PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { CodeTablePanelComponent } from './code-table-panel.component'
import { PlanCodePanelComponent } from './plan-code.panel'
import { PlanPremiumPanelComponent } from './plan-premium.panel'

/** Maintenance > Plan (frmMaintenancePlan): the desktop's eight SSTab tabs in order. */
@Component({
  selector: 'app-plan-page',
  imports: [PageShellComponent, TabsComponent, TabDirective, CodeTablePanelComponent, PlanCodePanelComponent, PlanPremiumPanelComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Plan Maintenance" subtitle="Replaces frmMaintenancePlan">
      <app-tabs label="Plan Maintenance">
        <ng-template appTab="HealthType"><app-code-table-panel tableKey="health-type" /></ng-template>
        <ng-template appTab="Insured Type"><app-code-table-panel tableKey="insured-type" /></ng-template>
        <ng-template appTab="Plan Code"><app-plan-code-panel /></ng-template>
        <ng-template appTab="Premium"><app-plan-premium-panel /></ng-template>
        @for (t of pending; track t) {
          <ng-template [appTab]="t"><p class="muted">{{ t }}: not migrated yet.</p></ng-template>
        }
      </app-tabs>
    </app-page-shell>
  `,
})
export class PlanPage {
  protected readonly pending = ['Annual Limit', 'Visit Limit', 'MCO Fee', 'Co-Payment (MBM)']
}
