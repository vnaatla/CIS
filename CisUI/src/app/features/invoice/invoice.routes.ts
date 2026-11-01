import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Invoice screens. Paths match layout/menu.ts. */
export const invoiceRoutes: Routes = [
  {
    path: 'invoice/case-transfer',
    title: 'CIS - Case Transfer',
    canActivate: [accessGuard],
    data: { access: Access.CaseTransfer, legacyForm: 'frmMembershipTransfer' },
    loadComponent: () => import('./case-transfer.page').then((m) => m.CaseTransferPage),
  },
  {
    path: 'invoice/general-exclusions/enquiry',
    title: 'CIS - General Exclusions (Enquiry)',
    canActivate: [accessGuard],
    data: { access: Access.GeneralExclusionsEnquiry, legacyForm: 'frmGeneralBenefitExcEnq' },
    loadComponent: () => import('./general-exclusions-enquiry.page').then((m) => m.GeneralExclusionsEnquiryPage),
  },
]
