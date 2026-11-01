import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'
import type { CodeTableTab } from './code-tables.page'

const codeTables = (path: string, title: string, legacyForm: string, access: number, tabs: CodeTableTab[]) => ({
  path,
  title: `CIS - ${title}`,
  canActivate: [accessGuard],
  data: { access, legacyForm, title, tabs },
  loadComponent: () => import('./code-tables.page').then((m) => m.CodeTablesPage),
})

/** Maintenance screens (Phase 1). Paths match layout/menu.ts. */
export const maintenanceRoutes: Routes = [
  codeTables('maintenance/check-item', 'Check Item', 'frmCheckingItems', Access.MaintenanceCheckItem, [{ label: 'Checking Items', key: 'check-items' }]),
  codeTables('maintenance/calculation', 'Calculation', 'frmMaintenanceCal', Access.MaintenanceCalculation, [
    { label: 'Co-Payment Cal.', key: 'co-payment' },
    { label: 'Fee Payable 2 Clinic', key: 'fee' },
  ]),
  {
    path: 'maintenance/lg',
    title: 'CIS - LG Maintenance',
    canActivate: [accessGuard],
    data: { access: Access.MaintenanceLg, legacyForm: 'frmLGMaintenance' },
    loadComponent: () => import('./lg-maintenance.page').then((m) => m.LgMaintenancePage),
  },
  codeTables('maintenance/case', 'Case', 'frmMaintenanceOthers', Access.MaintenanceCase, [
    { label: 'Treatment Reason', key: 'treatment-reason' },
    { label: 'Procedures', key: 'procedures' },
    { label: 'Drugs', key: 'drugs' },
    { label: 'Diagnosis Not Payable', key: 'diagnosis-not-payable' },
    { label: 'Repudiation', key: 'repudiation' },
  ]),
]
