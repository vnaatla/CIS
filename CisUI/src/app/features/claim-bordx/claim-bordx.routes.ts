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
]
