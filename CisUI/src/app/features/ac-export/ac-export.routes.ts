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
]
