import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Payment screens. Paths match layout/menu.ts. */
export const paymentRoutes: Routes = [
  {
    path: 'payment/clinic',
    title: 'CIS - Payment to Clinic',
    canActivate: [accessGuard],
    data: { access: Access.PaymentClinic, legacyForm: 'frmPaymentClinic' },
    loadComponent: () => import('./payment-clinic.page').then((m) => m.PaymentClinicPage),
  },
  {
    path: 'payment/clinic-invoice',
    title: 'CIS - Payment to Clinic - Invoice',
    canActivate: [accessGuard],
    data: { access: Access.PaymentClinicInvoice, legacyForm: 'frmPaymentCLNInvoice' },
    loadComponent: () => import('./payment-clinic-invoice.page').then((m) => m.PaymentClinicInvoicePage),
  },
  {
    path: 'payment/member',
    title: 'CIS - Payment to Member',
    canActivate: [accessGuard],
    data: { access: Access.PaymentMember, legacyForm: 'frmPaymentMemberAuto08' },
    loadComponent: () => import('./payment-member.page').then((m) => m.PaymentMemberPage),
  },
]
