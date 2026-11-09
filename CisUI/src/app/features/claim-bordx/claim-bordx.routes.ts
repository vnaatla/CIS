import { Routes } from '@angular/router'
import { Access } from '../../core/auth/access'
import { accessGuard } from '../../core/auth/guards'

/** Claim Bordx screens. Paths match layout/menu.ts. */
export const claimBordxRoutes: Routes = [
  {
    path: 'claim-bordx/insert-worksheet',
    title: 'CIS - Insert Worksheet',
    canActivate: [accessGuard],
    data: { access: Access.BordxInsert, legacyForm: 'frmWSBordx' },
    loadComponent: () => import('./insert-worksheet.page').then((m) => m.InsertWorksheetPage),
  },
  {
    path: 'claim-bordx/insert-invoice',
    title: 'CIS - Insert Invoice',
    canActivate: [accessGuard],
    data: { access: Access.BordxInsert, legacyForm: 'frmWSBordx2' },
    loadComponent: () => import('./insert-invoice.page').then((m) => m.InsertInvoicePage),
  },
  {
    path: 'claim-bordx/remove-worksheet',
    title: 'CIS - Remove Worksheet',
    canActivate: [accessGuard],
    data: { access: Access.BordxRemove, legacyForm: 'FrmBordxRevise' },
    loadComponent: () => import('./remove-worksheet.page').then((m) => m.RemoveWorksheetPage),
  },
  {
    path: 'claim-bordx/bordx-date',
    title: 'CIS - Bordx Date',
    canActivate: [accessGuard],
    data: { access: Access.BordxInsert, legacyForm: 'FrmBordxDate' },
    loadComponent: () => import('./bordx-date.page').then((m) => m.BordxDatePage),
  },
  ...(
    [
      ['send-to-payor', 'payor', Access.BordxSendToPayor, 'FrmSendtoPayor', 'Send to Grp. Company'],
      ['returned', 'returned', Access.BordxReturnedFromPayor, 'FrmReturnPayor', 'Returned Fr Grp. Company'],
      ['send-to-ac', 'ac', Access.BordxSendToAc, 'FrmSendtoAC', 'Send to A/C'],
      ['send-to-mnrb', 'mnrb', Access.BordxSendToPayor, 'FrmSendtoMNRB', 'Send to MNRB'],
    ] as const
  ).map(([path, kind, access, legacyForm, title]) => ({
    path: `claim-bordx/delivery/${path}`,
    title: `CIS - ${title}`,
    canActivate: [accessGuard],
    data: { access, legacyForm, kind },
    loadComponent: () => import('./delivery.page').then((m) => m.BordxDeliveryPage),
  })),
  {
    path: 'claim-bordx/ac-period',
    title: 'CIS - Bor Acc Period',
    canActivate: [accessGuard],
    data: { access: Access.BordxAcPeriod, legacyForm: 'FrmACPeriod', kind: 'acperiod' },
    loadComponent: () => import('./delivery.page').then((m) => m.BordxDeliveryPage),
  },
  {
    path: 'claim-bordx/print/standard',
    title: 'CIS - Print Bordx',
    canActivate: [accessGuard],
    data: { access: Access.BordxPrint, legacyForm: 'frmPrintBordxExcel' },
    loadComponent: () => import('./print-bordx.page').then((m) => m.PrintBordxPage),
  },
  ...(
    [
      ['seb', 'frmPrintBordxIBM', 'Print Bordx (SEB/INDAH WATER)'],
      ['prasarana', 'frmPrintBordxSPNB', 'Print Bordx (Prasarana)'],
      ['axa-affin', 'frmPrintBordxAXAAffin', 'Print Bordx (AXA Affin)'],
      ['time-dotcom', 'frmPrintBordxTimeDotcom', 'Print Bordx (Time Dotcom)'],
      ['bsn', 'frmPrintBordxBSN', 'Print Bordx (BSN)'],
      ['berjaya-sompo', 'frmPrintBordxBSI', 'Print Bordx (Berjaya Sompo)'],
      ['crystal', 'frmPrintBordx', 'Print Bordx (Crystal)'],
    ] as const
  ).map(([variant, legacyForm, screenTitle]) => ({
    path: `claim-bordx/print/${variant}`,
    title: `CIS - ${screenTitle}`,
    canActivate: [accessGuard],
    data: { access: Access.BordxPrint, legacyForm, variant, screenTitle },
    loadComponent: () => import('./print-bordx.page').then((m) => m.PrintBordxPage),
  })),
  ...(
    [
      ['control', Access.BordxControl, 'FrmBordxControl', 'Bordx Control'],
      ['outstanding', Access.BordxOutstanding, 'FrmOSBordx', 'Outstanding Bordx'],
      ['bordx-worksheet', Access.BordxWorksheet, 'FrmBordxWks', 'Bordx & Worksheet'],
    ] as const
  ).map(([kind, access, legacyForm, title]) => ({
    path: `claim-bordx/monitoring/${kind}`,
    title: `CIS - ${title}`,
    canActivate: [accessGuard],
    data: { access, legacyForm, kind },
    loadComponent: () => import('./monitoring.page').then((m) => m.BordxMonitoringPage),
  })),
]
