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
  {
    path: 'maintenance/plan',
    title: 'CIS - Plan Maintenance',
    canActivate: [accessGuard],
    data: { access: Access.MaintenancePlan, legacyForm: 'frmMaintenancePlan' },
    loadComponent: () => import('./plan.page').then((m) => m.PlanPage),
  },
  {
    path: 'maintenance/entity',
    title: 'CIS - Entity Maintenance',
    canActivate: [accessGuard],
    data: { access: Access.MaintenanceEntity, legacyForm: 'frmMaintenanceEntity' },
    loadComponent: () => import('./entity.page').then((m) => m.EntityPage),
  },
  {
    path: 'maintenance/benefit/member',
    title: 'CIS - Benefits Maintenance',
    canActivate: [accessGuard],
    data: { access: Access.MaintenanceBenefitMember, legacyForm: 'frmMaintenanceBenefits' },
    loadComponent: () => import('./benefit-member.page').then((m) => m.BenefitMemberPage),
  },
  {
    path: 'maintenance/benefit/dental',
    title: 'CIS - Dental Benefits Maintenance',
    canActivate: [accessGuard],
    data: { access: Access.MaintenanceBenefitMember, legacyForm: 'frmMaintenanceBNFDental' },
    loadComponent: () => import('./benefit-dental.page').then((m) => m.BenefitDentalPage),
  },
  {
    path: 'maintenance/benefit/clinic',
    title: 'CIS - Clinic Benefits Maintenance',
    canActivate: [accessGuard],
    data: { access: Access.MaintenanceBenefitClinic, legacyForm: 'frmMaintenanceBenefitsCln' },
    loadComponent: () => import('./benefit-clinic.page').then((m) => m.BenefitClinicPage),
  },
  {
    path: 'maintenance/benefit/clinic-new',
    title: 'CIS - Clinic Benefits Maintenance',
    canActivate: [accessGuard],
    data: { access: Access.MaintenanceBenefitClinic, legacyForm: 'frmMaintenanceBenefitsClnNew' },
    loadComponent: () => import('./benefit-clinic.page').then((m) => m.BenefitClinicPage),
  },
  codeTables('maintenance/case', 'Case', 'frmMaintenanceOthers', Access.MaintenanceCase, [
    { label: 'Treatment Reason', key: 'treatment-reason' },
    { label: 'Procedures', key: 'procedures' },
    { label: 'Drugs', key: 'drugs' },
    { label: 'Diagnosis Not Payable', key: 'diagnosis-not-payable' },
    { label: 'Repudiation', key: 'repudiation' },
  ]),
]
