import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Receipt screens. Paths match layout/menu.ts. */
export const receiptRoutes: Routes = [
  {
    path: 'receipt/member',
    title: 'CIS - Receipt from Member',
    canActivate: [accessGuard],
    data: { access: Access.ReceiptMember, legacyForm: 'FrmReceiptMember' },
    loadComponent: () => import('./receipt-member.page').then((m) => m.ReceiptMemberPage),
  },
  {
    path: 'receipt/payor',
    title: 'CIS - Receipt from Group Company',
    canActivate: [accessGuard],
    data: { access: Access.ReceiptPayor, legacyForm: 'FrmReceiptPayor' },
    loadComponent: () => import('./receipt-payor.page').then((m) => m.ReceiptPayorPage),
  },
]
