import { ChangeDetectionStrategy, Component, inject } from '@angular/core'
import { RouterLink } from '@angular/router'
import { SessionService } from '../core/auth/session.service'
import { IconComponent } from '../shared/icon.component'
import { MENU, buildable } from '../layout/menu'

type Module = (typeof MENU)[number]

/** Landing page after login: a welcome banner with the overall progress, and one card per desktop top menu with its migrated screens. */
@Component({
  selector: 'app-home-page',
  imports: [RouterLink, IconComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <section class="home">
      <div class="home-hero">
        <div class="home-hero-text">
          <p class="home-date">{{ today }}</p>
          <h2>{{ greeting }}, {{ session.user()?.userName }}</h2>
          <p>
            The CIS desktop (MedixCIS) is being migrated module by module. Open a module to see each screen, the VB6 form it replaces and
            its status.
          </p>
        </div>
        <div class="home-progress" role="img" [attr.aria-label]="migratedAll + ' of ' + totalAll + ' screens migrated'">
          <svg viewBox="0 0 120 120" aria-hidden="true">
            <circle cx="60" cy="60" r="52" class="track" />
            <circle cx="60" cy="60" r="52" class="value" [attr.stroke-dasharray]="ring + ' 999'" />
          </svg>
          <div class="home-progress-label">
            <strong>{{ percent }}%</strong>
            <span>migrated</span>
          </div>
        </div>
      </div>

      <div class="home-stats">
        <div class="home-stat"><span><app-icon name="grid" /></span><div><strong>{{ modules.length }}</strong><small>Modules</small></div></div>
        <div class="home-stat"><span><app-icon name="check" /></span><div><strong>{{ migratedAll }}</strong><small>Screens migrated</small></div></div>
        <div class="home-stat"><span><app-icon name="list" /></span><div><strong>{{ totalAll }}</strong><small>Screens to build</small></div></div>
        <div class="home-stat"><span><app-icon name="shield" /></span><div><strong>{{ totalAll - migratedAll }}</strong><small>Remaining</small></div></div>
      </div>

      <h3 class="home-section">Modules</h3>
      <div class="home-modules">
        @for (m of modules; track m.key) {
          <a [routerLink]="m.path" [class.complete]="migrated(m) === total(m)">
            <span class="home-module-icon"><app-icon [name]="m.icon" [size]="20" /></span>
            <span class="home-module-body">
              <span class="home-module-name">{{ m.label }}</span>
              <span class="count">{{ migrated(m) }} / {{ total(m) }} migrated</span>
              <span class="home-bar" aria-hidden="true"><span [style.width.%]="total(m) ? (migrated(m) / total(m)) * 100 : 100"></span></span>
            </span>
            <span class="home-module-go" aria-hidden="true"><app-icon name="chevRight" [size]="16" /></span>
          </a>
        }
      </div>
    </section>
  `,
})
export class HomePage {
  protected readonly session = inject(SessionService)
  protected readonly modules = MENU
  protected total = (m: Module) => buildable(m).length
  protected migrated = (m: Module) => buildable(m).filter((i) => i.migrated).length
  protected readonly totalAll = MENU.reduce((n, m) => n + buildable(m).length, 0)
  protected readonly migratedAll = MENU.reduce((n, m) => n + buildable(m).filter((i) => i.migrated).length, 0)
  protected readonly percent = this.totalAll ? Math.round((this.migratedAll / this.totalAll) * 100) : 100
  /** Length of the progress arc (circumference 2·π·52 ≈ 326.7). */
  protected readonly ring = ((this.totalAll ? this.migratedAll / this.totalAll : 1) * 326.7).toFixed(1)
  protected readonly today = new Date().toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })
  protected readonly greeting = new Date().getHours() < 12 ? 'Good morning' : new Date().getHours() < 18 ? 'Good afternoon' : 'Good evening'
}
