import { ChangeDetectionStrategy, Component } from '@angular/core'
import { PageShellComponent } from '../../shared/screen.components'
import { TabDirective, TabsComponent } from '../../shared/tabs.component'
import { BenefitEntryPanelComponent } from './benefit-member.page'
import { CodeTablePanelComponent } from './code-table-panel.component'

/** Maintenance > Benefit > Dental (frmMaintenanceBNFDental): Tooth Category · Tooth Reference · Benefit Category · Benefit Entry. */
@Component({
  selector: 'app-benefit-dental-page',
  imports: [PageShellComponent, TabsComponent, TabDirective, CodeTablePanelComponent, BenefitEntryPanelComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-page-shell title="Dental Benefits Maintenance" subtitle="Replaces frmMaintenanceBNFDental">
      <app-tabs label="Dental Benefits Maintenance">
        <ng-template appTab="Tooth Category"><app-code-table-panel tableKey="tooth-category" /></ng-template>
        <ng-template appTab="Tooth Reference"><app-code-table-panel tableKey="tooth-reference" /></ng-template>
        <ng-template appTab="Benefit Category"><app-code-table-panel tableKey="dental-benefit-category" /></ng-template>
        <ng-template appTab="Benefit Entry">
          <app-benefit-entry-panel api="/api/maintenance/benefit-dental" printReport="dental-benefit-entry" [showRefresh]="true" />
        </ng-template>
      </app-tabs>
    </app-page-shell>
  `,
})
export class BenefitDentalPage {}
