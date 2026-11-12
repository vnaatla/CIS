import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Tracking and report screens (Phase 10). Paths match layout/menu.ts. */
export const reportsRoutes: Routes = [
  {
    path: 'tracking/clinic',
    title: 'CIS - Clinic Tracking',
    canActivate: [accessGuard],
    data: { access: Access.ClinicTracking, legacyForm: 'FrmTracking' },
    loadComponent: () => import('./clinic-tracking.page').then((m) => m.ClinicTrackingPage),
  },
  {
    path: 'tracking/grp-company-clinic',
    title: 'CIS - Group Company Clinic Tracking',
    canActivate: [accessGuard],
    data: { access: Access.ClinicTracking, legacyForm: 'FrmTracking2' },
    loadComponent: () => import('./grp-company-clinic-tracking.page').then((m) => m.GrpCompanyClinicTrackingPage),
  },
]
