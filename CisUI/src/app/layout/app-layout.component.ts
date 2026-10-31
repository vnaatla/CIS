import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core'
import { toSignal } from '@angular/core/rxjs-interop'
import { NavigationEnd, Router, RouterLink, RouterLinkActive, RouterOutlet } from '@angular/router'
import { filter, map } from 'rxjs'
import { SessionService } from '../core/auth/session.service'
import { IconComponent } from '../shared/icon.component'
import { MENU, MenuItem, MenuModule } from './menu'

const NAV_COLLAPSED = 'nav-collapsed'

/**
 * Main window (VB6 MDIMain): GENESIS topbar + icon rail + sidebar. The rail holds the desktop's top menus, the sidebar the items of the
 * module under the pointer or route, grouped by the desktop's sub-menus. Items not migrated yet are shown disabled, items without the
 * user's right are disabled with "Access Denied!" (the desktop's message).
 */
@Component({
  selector: 'app-layout',
  imports: [RouterOutlet, RouterLink, RouterLinkActive, IconComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="shell" [class.no-sidebar]="!section()" [class.nav-collapsed]="!open() && !!section()">
      <header class="topbar">
        <button type="button" class="nav-toggle" (click)="toggleNav()" [title]="open() ? 'Hide sidebar' : 'Show sidebar'" [attr.aria-label]="open() ? 'Hide sidebar' : 'Show sidebar'" [attr.aria-expanded]="open()">
          <app-icon [name]="open() ? 'chevLeft' : 'chevRight'" />
        </button>
        <span class="brand-mark"><app-icon name="cross" [size]="14" /></span>
        <strong>CIS</strong>
        <span class="muted">MediExpress Clinical Information System</span>
        <span class="spacer"></span>
        <button type="button" class="link" title="Print this screen (File > Print)" (click)="print()">Print</button>
        <span>{{ session.user()?.userName }} ({{ session.user()?.userCode }})</span>
        <button type="button" class="link" (click)="logout()">Log out</button>
      </header>

      <nav class="rail" aria-label="Modules" (mouseenter)="open.set(true)" (mouseleave)="navLeave($event)">
        <a routerLink="/" title="Home" aria-label="Home" [class.active]="path() === '/'"><app-icon name="home" /></a>
        @for (m of modules; track m.key) {
          <a [routerLink]="m.path" [title]="m.label" [attr.aria-label]="m.label" [class.active]="path().startsWith(m.path)" (mouseenter)="hover.set(m.key)">
            <app-icon [name]="m.icon" />
          </a>
        }
      </nav>

      @if (section(); as s) {
        <nav class="sidebar" aria-label="Pages" (mouseenter)="open.set(true)" (mouseleave)="navLeave($event)">
          <h4><a [routerLink]="s.path" routerLinkActive="active" [routerLinkActiveOptions]="{ exact: true }">{{ s.label }}</a></h4>
          @for (i of sidebarItems(s); track i.path; let n = $index) {
            @if (i.group && i.group !== sidebarItems(s)[n - 1]?.group) {
              <h5 class="nav-group">{{ i.group }}</h5>
            }
            @if (!i.migrated) {
              <span class="disabled" [title]="i.label + ' — not migrated yet (phase ' + i.phase + ')'">{{ i.label }}</span>
            } @else if (session.hasAccess(i.access)) {
              <a [routerLink]="i.path" routerLinkActive="active">{{ i.label }}</a>
            } @else {
              <span class="disabled" title="Access Denied!">{{ i.label }}</span>
            }
          }
        </nav>
      }

      <main class="content">
        <router-outlet />
      </main>
    </div>
  `,
})
export class AppLayoutComponent {
  protected readonly session = inject(SessionService)
  private readonly router = inject(Router)
  protected readonly modules = MENU
  protected readonly open = signal(localStorage.getItem(NAV_COLLAPSED) !== '1')
  protected readonly hover = signal<string | null>(null)
  protected readonly path = toSignal(
    this.router.events.pipe(
      filter((e): e is NavigationEnd => e instanceof NavigationEnd),
      map((e) => e.urlAfterRedirects.split('?')[0]),
    ),
    { initialValue: this.router.url.split('?')[0] },
  )

  protected readonly section = computed<MenuModule | undefined>(() => {
    const hovered = this.hover() && MENU.find((m) => m.key === this.hover())
    return hovered || MENU.find((m) => this.path().startsWith(m.path))
  })

  /** Sidebar entries: everything that needs a screen (menu items without a handler and a form are only on the module page). */
  protected sidebarItems(m: MenuModule): MenuItem[] {
    return m.items.filter((i) => i.desktop !== 'nothing')
  }

  protected toggleNav(): void {
    this.open.update((o) => {
      localStorage.setItem(NAV_COLLAPSED, o ? '1' : '0')
      return !o
    })
  }

  /** Hide when the pointer leaves the rail + sidebar (but not when moving between them). */
  protected navLeave(e: MouseEvent): void {
    if (!(e.relatedTarget instanceof Element && e.relatedTarget.closest('.rail, .sidebar'))) {
      this.open.set(false)
      this.hover.set(null)
    }
  }

  protected print(): void {
    window.print()
  }

  protected async logout(): Promise<void> {
    await this.session.logout()
    await this.router.navigateByUrl('/login')
  }
}
