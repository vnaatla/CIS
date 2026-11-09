import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** A/C Export screens. Paths match layout/menu.ts. */
export const acExportRoutes: Routes = [
  {
    path: 'ac-export/invoice/clinic',
    title: 'CIS - Invoice Posting - Clinic',
    canActivate: [accessGuard],
    data: { access: Access.AcExport, legacyForm: 'frmPostingInvoiceCln' },
    loadComponent: () => import('./posting-invoice-clinic.page').then((m) => m.PostingInvoiceClinicPage),
  },
  {
    path: 'ac-export/invoice/clinic-gst',
    title: 'CIS - Invoice Posting - Clinic (GST)',
    canActivate: [accessGuard],
    data: { access: Access.AcExport, legacyForm: 'frmPostingInvoiceClnGST', variant: 'gst' },
    loadComponent: () => import('./posting-invoice-clinic.page').then((m) => m.PostingInvoiceClinicPage),
  },
  {
    path: 'ac-export/invoice/clinic-gst-2017',
    title: 'CIS - Invoice Posting - Clinic (GST-2017)',
    canActivate: [accessGuard],
    data: { access: Access.AcExport, legacyForm: 'frmPostingInvClnGST2017', variant: 'gst2017' },
    loadComponent: () => import('./posting-invoice-clinic.page').then((m) => m.PostingInvoiceClinicPage),
  },
  {
    path: 'ac-export/invoice/clinic-gst-2017-zeroised',
    title: 'CIS - Invoice Posting - Clinic (GST-2017 Zeroised)',
    canActivate: [accessGuard],
    data: { access: Access.AcExport, legacyForm: 'frmPostingInvClnGST2017Zeroised', variant: 'zeroised' },
    loadComponent: () => import('./posting-invoice-clinic.page').then((m) => m.PostingInvoiceClinicPage),
  },
  {
    path: 'ac-export/invoice/member',
    title: 'CIS - Invoice Posting - Member',
    canActivate: [accessGuard],
    data: { access: Access.AcExport, legacyForm: 'frmPostingInvoiceMember' },
    loadComponent: () => import('./posting-invoice-member.page').then((m) => m.PostingInvoiceMemberPage),
  },
]
