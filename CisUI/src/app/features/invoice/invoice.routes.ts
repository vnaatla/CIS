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
  {
    path: 'invoice/batch-transfer',
    title: 'CIS - Batch Transfer',
    canActivate: [accessGuard],
    data: { access: Access.InvoiceAdjustment, legacyForm: 'FrmInvoiceBatchNoUpd' },
    loadComponent: () => import('./batch-transfer.page').then((m) => m.BatchTransferPage),
  },
  {
    path: 'invoice/update',
    title: 'CIS - Invoice Update',
    canActivate: [accessGuard],
    data: { access: Access.InvoiceAdjustment, legacyForm: 'frmUpdateCLNInvoice' },
    loadComponent: () => import('./invoice-update.page').then((m) => m.InvoiceUpdatePage),
  },
  {
    path: 'invoice/op-batch',
    title: 'CIS - OP Batch/Invoice',
    canActivate: [accessGuard],
    data: { access: Access.InvoiceAdjustment, legacyForm: 'frmUpdateInvoiceBatch' },
    loadComponent: () => import('./op-batch.page').then((m) => m.OpBatchPage),
  },
  {
    path: 'invoice/mma-cpt',
    title: 'CIS - MMA/CPT',
    canActivate: [accessGuard],
    data: { access: Access.MmaCpt, legacyForm: 'frmMMASchedule' },
    loadComponent: () => import('./mma-cpt.page').then((m) => m.MmaCptPage),
  },
]
