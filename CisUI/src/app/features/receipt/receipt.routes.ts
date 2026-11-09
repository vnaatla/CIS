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
]
