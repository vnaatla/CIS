import { ChangeDetectionStrategy, Component } from '@angular/core'
import { PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { CodeTablePanelComponent } from './code-table-panel.component'
import { LgBindingPanelComponent } from './lg-binding-panel.component'

/** Maintenance > LGMaintenance (frmLGMaintenance): the desktop's four SSTab tabs in their order. */
@Component({
  selector: 'app-lg-maintenance-page',
  imports: [PageShellComponent, TabsComponent, TabDirective, CodeTablePanelComponent, LgBindingPanelComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="LG Maintenance" subtitle="Replaces frmLGMaintenance">
      <app-tabs label="LG Maintenance">
        <ng-template appTab="GL Template Binding"><app-lg-binding-panel /></ng-template>
        <ng-template appTab="GL Type Binding"><app-code-table-panel tableKey="gl-type" /></ng-template>
        <ng-template appTab="Template Name Binding"><app-code-table-panel tableKey="gl-template-name" /></ng-template>
        <ng-template appTab="Decline Reason"><app-code-table-panel tableKey="decline-reason" /></ng-template>
      </app-tabs>
    </app-page-shell>
  `,
})
export class LgMaintenancePage {}
