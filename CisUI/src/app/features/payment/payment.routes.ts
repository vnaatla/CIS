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
]
