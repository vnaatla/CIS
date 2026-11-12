import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Tracking and report screens (Phase 10). Paths match layout/menu.ts. */
export const reportsRoutes: Routes = [
  {
    path: 'tracking/clinic',
    title: 'CIS - Clinic Tracking',
    canActivate: [accessGuard],
    data: { access: Access.ClinicTracking, legacyForm: 'FrmTracking' },
    loadComponent: () => import('./clinic-tracking.page').then((m) => m.ClinicTrackingPage),
  },
  {
    path: 'invoice/tracking',
    title: 'CIS - Invoice Tracking',
    canActivate: [accessGuard],
    data: { access: Access.InvoiceTracking, legacyForm: 'FrmInvoiceTracking' },
    loadComponent: () => import('./invoice-tracking.page').then((m) => m.InvoiceTrackingPage),
  },
  {
    path: 'invoice/tracking-jupiter',
    title: 'CIS - Invoice Tracking (Jupiter)',
    canActivate: [accessGuard],
    data: { access: Access.InvoiceTrackingJupiter, legacyForm: 'FrmInvoiceTracking2', variant: 'jupiter' },
    loadComponent: () => import('./invoice-tracking.page').then((m) => m.InvoiceTrackingPage),
  },
  {
    path: 'invoice/email/payment',
    title: 'CIS - Email Notification (Payment)',
    canActivate: [accessGuard],
    data: { access: Access.EmailNotification, legacyForm: 'frmEmailNotification' },
    loadComponent: () => import('./email-payment.page').then((m) => m.EmailPaymentPage),
  },
  {
    path: 'invoice/email/providers',
    title: 'CIS - Email Notification (Providers)',
    canActivate: [accessGuard],
    data: { access: Access.EmailNotification, legacyForm: 'frmEmailNotificationProvider' },
    loadComponent: () => import('./email-providers.page').then((m) => m.EmailProvidersPage),
  },
  {
    path: 'invoice/email/rejection',
    title: 'CIS - Email Notification (Rejection)',
    canActivate: [accessGuard],
    data: { access: Access.EmailNotification, legacyForm: 'frmEmailRejectedClaim' },
    loadComponent: () => import('./email-rejection.page').then((m) => m.EmailRejectionPage),
  },
  {
    path: 'tracking/grp-company-clinic',
    title: 'CIS - Group Company Clinic Tracking',
    canActivate: [accessGuard],
    data: { access: Access.ClinicTracking, legacyForm: 'FrmTracking2' },
    loadComponent: () => import('./grp-company-clinic-tracking.page').then((m) => m.GrpCompanyClinicTrackingPage),
  },
  {
    path: 'invoice/claim-ascii/bintulu',
    title: 'CIS - Claim ASCII (bintulu)',
    canActivate: [accessGuard],
    data: { access: Access.ClaimAscii, legacyForm: 'frmClaimDetailsASCII', format: 'bintulu' },
    loadComponent: () => import('./claim-ascii.page').then((m) => m.ClaimAsciiPage),
  },
  {
    path: 'invoice/claim-ascii/toshiba',
    title: 'CIS - Claim ASCII (toshiba)',
    canActivate: [accessGuard],
    data: { access: Access.ClaimAscii, legacyForm: 'frmClaimASCIIToshiba', format: 'toshiba' },
    loadComponent: () => import('./claim-ascii.page').then((m) => m.ClaimAsciiPage),
  },
  {
    path: 'invoice/claim-ascii/singer',
    title: 'CIS - Claim ASCII (singer)',
    canActivate: [accessGuard],
    data: { access: Access.ClaimAscii, legacyForm: 'frmClaimASCIISinger', format: 'singer' },
    loadComponent: () => import('./claim-ascii.page').then((m) => m.ClaimAsciiPage),
  },
  {
    path: 'invoice/claim-ascii/tims',
    title: 'CIS - Claim ASCII (tims)',
    canActivate: [accessGuard],
    data: { access: Access.ClaimAscii, legacyForm: 'frmClaimASCIITIMS', format: 'tims' },
    loadComponent: () => import('./claim-ascii.page').then((m) => m.ClaimAsciiPage),
  },
  {
    path: 'invoice/claim-ascii/cimb',
    title: 'CIS - Claim ASCII (cimb)',
    canActivate: [accessGuard],
    data: { access: Access.ClaimAscii, legacyForm: 'frmClaimASCIICIMB', format: 'cimb' },
    loadComponent: () => import('./claim-ascii.page').then((m) => m.ClaimAsciiPage),
  },
  {
    path: 'invoice/claim-ascii/omni',
    title: 'CIS - Claim ASCII (omni)',
    canActivate: [accessGuard],
    data: { access: Access.ClaimAscii, legacyForm: 'frmClaimDataOMNI', format: 'omni' },
    loadComponent: () => import('./claim-ascii.page').then((m) => m.ClaimAsciiPage),
  },
  {
    path: 'invoice/claim-ascii/omni-providers',
    title: 'CIS - Med Providers (OMNI)',
    canActivate: [accessGuard],
    data: { access: Access.ClaimAscii, legacyForm: 'frmMedProviderOMNI', format: 'omniproviders' },
    loadComponent: () => import('./claim-ascii.page').then((m) => m.ClaimAsciiPage),
  },
  {
    path: 'invoice/claim-ascii/utusan',
    title: 'CIS - Claim ASCII (utusan)',
    canActivate: [accessGuard],
    data: { access: Access.ClaimAscii, legacyForm: 'frmClaimASCIIUtusan', format: 'utusan' },
    loadComponent: () => import('./claim-ascii.page').then((m) => m.ClaimAsciiPage),
  },
]
