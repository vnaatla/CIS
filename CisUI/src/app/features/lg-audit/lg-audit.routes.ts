import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** LG & Audit screens. Paths match layout/menu.ts. */
export const lgAuditRoutes: Routes = [
  {
    path: 'lg-audit/generate-lg',
    title: 'CIS - Generate LG',
    canActivate: [accessGuard],
    data: { access: Access.GenerateLg, legacyForm: 'frmLGGeneration' },
    loadComponent: () => import('./generate-lg.page').then((m) => m.GenerateLgPage),
  },
  {
    path: 'lg-audit/audit-review',
    title: 'CIS - Audit Review',
    canActivate: [accessGuard],
    data: { access: Access.AuditReview, legacyForm: 'frmAuditInvoice' },
    loadComponent: () => import('./audit-review.page').then((m) => m.AuditReviewPage),
  },
  {
    path: 'lg-audit/audit-tracking',
    title: 'CIS - Audit Tracking',
    canActivate: [accessGuard],
    data: { access: Access.AuditReview, legacyForm: 'FrmAuditInvoiceTracking' },
    loadComponent: () => import('./audit-tracking.page').then((m) => m.AuditTrackingPage),
  },
]
