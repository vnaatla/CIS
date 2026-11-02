import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Claim screens. Paths match layout/menu.ts. */
export const claimRoutes: Routes = [
  {
    path: 'claim/pass-to-claim',
    title: 'CIS - Pass to Claim',
    canActivate: [accessGuard],
    data: { access: Access.PassToClaim, legacyForm: 'frmPass2Claim' },
    loadComponent: () => import('./pass-to-claim.page').then((m) => m.PassToClaimPage),
  },
]
