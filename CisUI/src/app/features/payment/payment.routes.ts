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
  {
    path: 'payment/member-legacy',
    title: 'CIS - Payment to Member (old)',
    canActivate: [accessGuard],
    data: { access: Access.PaymentMember, legacyForm: 'frmPaymentMember' },
    loadComponent: () => import('./payment-member-legacy.page').then((m) => m.PaymentMemberLegacyPage),
  },
  {
    path: 'payment/post/clinic',
    title: 'CIS - Pay Clinic Post',
    canActivate: [accessGuard],
    data: { access: Access.AcExport, legacyForm: 'frmPayClinicPost' },
    loadComponent: () => import('./payment-clinic-post.page').then((m) => m.PaymentClinicPostPage),
  },
  {
    path: 'payment/tracking',
    title: 'CIS - Payment Tracking',
    canActivate: [accessGuard],
    data: { access: Access.PaymentTracking, legacyForm: 'FrmPayment', screenTitle: 'Payment Tracking' },
    loadComponent: () => import('./payment-tracking.page').then((m) => m.PaymentTrackingPage),
  },
  {
    path: 'payment/tracking-jupiter',
    title: 'CIS - Payment Tracking (Jupiter)',
    canActivate: [accessGuard],
    data: { access: Access.PaymentTrackingJupiter, legacyForm: 'frmPayment2', screenTitle: 'Payment Tracking (Jupiter)' },
    loadComponent: () => import('./payment-tracking.page').then((m) => m.PaymentTrackingPage),
  },
  {
    path: 'payment/ffs',
    title: 'CIS - FFS Payment',
    canActivate: [accessGuard],
    data: { access: Access.PaymentFfs, legacyForm: 'FrmPaymentFFS' },
    loadComponent: () => import('./payment-ffs.page').then((m) => m.PaymentFfsPage),
  },
]
