import { Routes } from '@angular/router'
import { authGuard, guestGuard } from './core/auth/guards'

/**
 * Each migration phase adds its screens here as lazy routes under the module path, guarded by `accessGuard` with
 * `data: { access: Access.X, legacyForm: 'frmName' }` (legacyForm feeds the action log). Unknown paths inside a module show the
 * module overview.
 */
export const routes: Routes = [
  { path: 'login', canActivate: [guestGuard], loadComponent: () => import('./pages/login.page').then((m) => m.LoginPage), title: 'CIS - Login' },
  {
    path: '',
    canActivate: [authGuard],
    loadComponent: () => import('./layout/app-layout.component').then((m) => m.AppLayoutComponent),
    children: [
      { path: '', pathMatch: 'full', loadComponent: () => import('./pages/home.page').then((m) => m.HomePage), title: 'CIS' },
      {
        path: ':module',
        children: [
          { path: '', loadComponent: () => import('./pages/module.page').then((m) => m.ModulePage), title: 'CIS' },
          { path: '**', redirectTo: '' },
        ],
      },
    ],
  },
  { path: '**', redirectTo: '' },
]
