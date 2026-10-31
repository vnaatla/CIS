import { ChangeDetectionStrategy, Component, inject } from '@angular/core'
import { RouterLink } from '@angular/router'
import { SessionService } from '../core/auth/session.service'
import { IconComponent } from '../shared/icon.component'
import { MENU, buildable } from '../layout/menu'

/** Landing page after login: one card per desktop top menu with how many of its screens are migrated. */
@Component({
  selector: 'app-home-page',
  imports: [RouterLink, IconComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <section>
      <h2>Welcome, {{ session.user()?.userName }}</h2>
      <p class="muted">
        The CIS desktop (MedixCIS) is being migrated module by module. Every desktop menu item is listed; open a module to see each screen,
        the VB6 form it replaces and its status.
      </p>
      <div class="home-modules">
        @for (m of modules; track m.key) {
          <a [routerLink]="m.path">
            <app-icon [name]="m.icon" />
            <span>{{ m.label }}</span>
            <span class="count">{{ migrated(m) }} / {{ total(m) }} migrated</span>
          </a>
        }
      </div>
    </section>
  `,
})
export class HomePage {
  protected readonly session = inject(SessionService)
  protected readonly modules = MENU
  protected total = (m: (typeof MENU)[number]) => buildable(m).length
  protected migrated = (m: (typeof MENU)[number]) => buildable(m).filter((i) => i.migrated).length
}
