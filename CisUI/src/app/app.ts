import { ChangeDetectionStrategy, Component, inject } from '@angular/core'
import { Router, RouterOutlet } from '@angular/router'
import { ScreenAuditService } from './core/audit/screen-audit.service'
import { NotifyHostComponent } from './shared/notify-host.component'

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, NotifyHostComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<router-outlet /><app-notify-host />`,
})
export class App {
  constructor() {
    inject(ScreenAuditService).start(inject(Router))
  }
}
