import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** MCO Fees screens. Paths match layout/menu.ts. */
export const mcoFeesRoutes: Routes = [
  {
    path: 'mco-fees/invoice',
    title: 'CIS - MCO Invoice',
    canActivate: [accessGuard],
    data: { access: Access.McoInvoice, legacyForm: 'frmMCOInvoiceNew' },
    loadComponent: () => import('./invoice.page').then((m) => m.McoInvoicePage),
  },
  {
    path: 'mco-fees/fee-listing',
    title: 'CIS - MCO Fee Listing',
    canActivate: [accessGuard],
    data: { access: Access.McoFeeListing, legacyForm: 'frmMCOFeeListNew' },
    loadComponent: () => import('./fee-listing.page').then((m) => m.McoFeeListingPage),
  },
  {
    path: 'mco-fees/refund-listing',
    title: 'CIS - MCO Refund Listing',
    canActivate: [accessGuard],
    data: { access: Access.McoFeeListing, legacyForm: 'frmMCORefundListing' },
    loadComponent: () => import('./refund-listing.page').then((m) => m.McoRefundListingPage),
  },
]
