import { inject } from '@angular/core'
import { CanActivateFn, Router } from '@angular/router'
import { NotifyService } from '../notify/notify.service'
import { SessionService } from './session.service'

/** Pages behind the login. */
export const authGuard: CanActivateFn = () => {
  const session = inject(SessionService)
  return session.signedIn() ? true : inject(Router).parseUrl('/login')
}

/** The login page only when signed out. */
export const guestGuard: CanActivateFn = () => (inject(SessionService).signedIn() ? inject(Router).parseUrl('/') : true)

/**
 * A screen's menu right (route data `access`): the desktop answered "Access Denied!" and stayed where it was; the web shows the same
 * message and goes home.
 */
export const accessGuard: CanActivateFn = (route) => {
  const position = route.data['access'] as number | undefined
  if (position === undefined || inject(SessionService).hasAccess(position)) return true
  inject(NotifyService).notify('Access Denied!', 'error')
  return inject(Router).parseUrl('/')
}
