import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Claim Bordx screens. Paths match layout/menu.ts. */
export const claimBordxRoutes: Routes = [
  {
    path: 'claim-bordx/insert-worksheet',
    title: 'CIS - Insert Worksheet',
    canActivate: [accessGuard],
    data: { access: Access.BordxInsert, legacyForm: 'frmWSBordx' },
    loadComponent: () => import('./insert-worksheet.page').then((m) => m.InsertWorksheetPage),
  },
  {
    path: 'claim-bordx/remove-worksheet',
    title: 'CIS - Remove Worksheet',
    canActivate: [accessGuard],
    data: { access: Access.BordxRemove, legacyForm: 'FrmBordxRevise' },
    loadComponent: () => import('./remove-worksheet.page').then((m) => m.RemoveWorksheetPage),
  },
]
