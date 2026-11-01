import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Membership screens (Phase 2). Paths match layout/menu.ts. */
export const membershipRoutes: Routes = [
  {
    path: 'membership/topup',
    title: 'CIS - MBM Topup Adj',
    canActivate: [accessGuard],
    data: { access: Access.MbmTopup, legacyForm: 'FrmMBMTopupLmt' },
    loadComponent: () => import('./mbm-topup.page').then((m) => m.MbmTopupPage),
  },
  {
    path: 'membership/adjustment',
    title: 'CIS - MBM Adjustment',
    canActivate: [accessGuard],
    data: { access: Access.MbmAdjustment, legacyForm: 'FrmMBMCLNAdj' },
    loadComponent: () => import('./mbm-adjustment.page').then((m) => m.MbmAdjustmentPage),
  },
]
