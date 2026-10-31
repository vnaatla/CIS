import { HttpClient } from '@angular/common/http'
import { Injectable, inject } from '@angular/core'
import { ActivatedRouteSnapshot, NavigationEnd, Router } from '@angular/router'
import { filter } from 'rxjs'
import { SessionService } from '../auth/session.service'

/** Route data key holding the legacy form name of a screen (e.g. 'frmMBMCLNEnq'). */
export const LEGACY_FORM = 'legacyForm'

/**
 * The desktop logged every form it opened (Form_Load → ActionLogUpdate(12, "frmName") → HISDB.dbo.tblActionLog). This posts the same
 * record when a route with `data.legacyForm` is shown. Failures are ignored, as in the desktop (logging never blocked a screen).
 */
@Injectable({ providedIn: 'root' })
export class ScreenAuditService {
  private readonly http = inject(HttpClient)
  private readonly session = inject(SessionService)

  start(router: Router): void {
    router.events.pipe(filter((e): e is NavigationEnd => e instanceof NavigationEnd)).subscribe(() => {
      const form = deepestForm(router.routerState.snapshot.root)
      if (form) this.screenOpened(form)
    })
  }

  screenOpened(screen: string): void {
    if (!this.session.signedIn()) return
    this.http.post('/api/audit/screen-opened', { screen }).subscribe({ error: () => undefined })
  }
}

function deepestForm(route: ActivatedRouteSnapshot): string | undefined {
  let form: string | undefined
  for (let r: ActivatedRouteSnapshot | null = route; r; r = r.firstChild) form = (r.data[LEGACY_FORM] as string | undefined) ?? form
  return form
}
