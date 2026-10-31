import { HttpErrorResponse, HttpInterceptorFn } from '@angular/common/http'
import { inject } from '@angular/core'
import { Router } from '@angular/router'
import { catchError, throwError, timeout } from 'rxjs'
import { SessionService } from '../auth/session.service'

/** Longest a request may take (the biggest tracking reports and exports run for a while). */
export const REQUEST_TIMEOUT_MS = 120_000

/** Adds the bearer token; a 401 outside /api/auth/ means the session expired: forget it and go to the login page. */
export const authInterceptor: HttpInterceptorFn = (request, next) => {
  const session = inject(SessionService)
  const router = inject(Router)
  const token = session.token
  const authorised = token && request.url.startsWith('/api/') ? request.clone({ setHeaders: { Authorization: `Bearer ${token}` } }) : request
  return next(authorised).pipe(
    timeout(REQUEST_TIMEOUT_MS),
    catchError((error: unknown) => {
      if (error instanceof HttpErrorResponse && error.status === 401 && !request.url.includes('/api/auth/')) {
        session.clear()
        void router.navigateByUrl('/login')
      }
      return throwError(() => error)
    }),
  )
}
