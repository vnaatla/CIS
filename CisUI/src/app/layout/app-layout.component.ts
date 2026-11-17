import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core'
import { toSignal } from '@angular/core/rxjs-interop'
import { NavigationEnd, Router, RouterLink, RouterLinkActive, RouterOutlet } from '@angular/router'
import { filter, map } from 'rxjs'
import { SessionService } from '../core/auth/session.service'
import { IconComponent } from '../shared/icon.component'
import { MENU, MenuItem, MenuModule, nothingToBuild } from './menu'

/** Older setting (sidebar hidden); read once to choose the first pin state. */
const NAV_COLLAPSED = 'nav-collapsed'
const NAV_PINNED = 'nav-pinned'

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
    <div class="shell" [class.no-sidebar]="!docked()">
      <header class="topbar">
        <button type="button" class="nav-toggle" (click)="togglePin()" [title]="pinned() ? 'Unpin the sidebar' : 'Pin the sidebar'"
          [attr.aria-label]="pinned() ? 'Unpin the sidebar' : 'Pin the sidebar'" [attr.aria-pressed]="pinned()">
          <app-icon [name]="pinned() ? 'chevLeft' : 'chevRight'" />
        </button>
        <a routerLink="/" class="topbar-brand" aria-label="CIS home">
          <span class="brand-mark"><app-icon name="cross" [size]="16" /></span>
          <span class="topbar-title">
            <strong>CIS</strong>
            <small>MediExpress Clinical Information System</small>
          </span>
        </a>
        @if (section(); as s) {
          <span class="topbar-crumb"><app-icon name="chevRight" [size]="14" /> {{ s.label }}</span>
        }
        <span class="spacer"></span>
        <div class="topbar-actions">
          <button type="button" class="topbar-icon" title="Print this screen (File > Print)" aria-label="Print" (click)="print()">
            <app-icon name="printer" [size]="17" />
          </button>
          <span class="topbar-user" [title]="session.user()?.userName + ' (' + session.user()?.userCode + ')'">
            <span class="avatar" aria-hidden="true">{{ initials() }}</span>
            <span class="topbar-user-text">
              <strong>{{ session.user()?.userName }}</strong>
              <small>{{ session.user()?.userCode }}</small>
            </span>
          </span>
          <button type="button" class="link topbar-logout" (click)="logout()"><app-icon name="logout" [size]="16" /> Log out</button>
        </div>
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
        @if (docked() || open()) {
          <nav class="sidebar" [class.floating]="!docked()" aria-label="Pages" (mouseenter)="open.set(true)" (mouseleave)="navLeave($event)">
            <div class="sidebar-head">
              <a class="sidebar-module" [routerLink]="s.path" routerLinkActive="active" [routerLinkActiveOptions]="{ exact: true }">
                <span class="sidebar-module-icon"><app-icon [name]="s.icon" [size]="16" /></span>
                <span class="sidebar-module-text">
                  <strong>{{ s.label }}</strong>
                  <small>{{ migratedCount(s) }} / {{ sidebarItems(s).length }} screens</small>
                </span>
              </a>
              <button type="button" class="sidebar-pin" [class.on]="pinned()" (click)="togglePin()" [title]="pinned() ? 'Unpin the sidebar' : 'Pin the sidebar'"
                [attr.aria-label]="pinned() ? 'Unpin the sidebar' : 'Pin the sidebar'" [attr.aria-pressed]="pinned()">
                <app-icon [name]="pinned() ? 'pinOff' : 'pin'" [size]="15" />
              </button>
            </div>
            <div class="sidebar-items">
              @for (i of sidebarItems(s); track i.path; let n = $index) {
                @if (i.group && i.group !== sidebarItems(s)[n - 1]?.group) {
                  <h5 class="nav-group">{{ i.group }}</h5>
                }
                @if (!i.migrated) {
                  <span class="disabled" [title]="i.label + ' — not migrated yet (phase ' + i.phase + ')'">{{ i.label }}</span>
                } @else if (session.hasAccess(i.access)) {
                  <a [routerLink]="i.path" routerLinkActive="active">{{ i.label }}</a>
                } @else {
                  <span class="disabled" title="Access Denied!">{{ i.label }} <app-icon name="lockSmall" [size]="13" /></span>
                }
              }
            </div>
          </nav>
        }
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
  /** Pinned: the sidebar stays beside the content. Unpinned: it slides out over the content while the pointer is on the rail / sidebar. */
  protected readonly pinned = signal((localStorage.getItem(NAV_PINNED) ?? (localStorage.getItem(NAV_COLLAPSED) === '1' ? '0' : '1')) === '1')
  protected readonly open = signal(false)
  protected readonly hover = signal<string | null>(null)
  /** Up to two initials of the user's name, for the avatar. */
  protected readonly initials = computed(() =>
    (this.session.user()?.userName ?? this.session.user()?.userCode ?? '?')
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((w) => w[0]!.toUpperCase())
      .join(''),
  )
  protected readonly path = toSignal(
    this.router.events.pipe(
      filter((e): e is NavigationEnd => e instanceof NavigationEnd),
      map((e) => e.urlAfterRedirects.split('?')[0]),
    ),
    { initialValue: this.router.url.split('?')[0] },
  )

  /** The module of the current route. */
  private readonly current = computed(() => MENU.find((m) => this.path().startsWith(m.path)))

  /** The module whose items the sidebar shows: the one under the pointer, else the current one. */
  protected readonly section = computed<MenuModule | undefined>(() => {
    const hovered = this.hover() && MENU.find((m) => m.key === this.hover())
    return hovered || this.current()
  })

  /** Docked beside the content when pinned inside a module; otherwise the sidebar floats over the content while hovered. */
  protected readonly docked = computed(() => this.pinned() && !!this.current())

  /** Sidebar entries: everything that needs a screen (menu items without a handler and a form are only on the module page). */
  protected sidebarItems(m: MenuModule): MenuItem[] {
    return m.items.filter((i) => !nothingToBuild(i))
  }

  /** Migrated screens of a module (for the sidebar header). */
  protected migratedCount(m: MenuModule): number {
    return this.sidebarItems(m).filter((i) => i.migrated).length
  }

  /** Pin / unpin the sidebar (kept in localStorage). */
  protected togglePin(): void {
    this.pinned.update((p) => {
      localStorage.setItem(NAV_PINNED, p ? '0' : '1')
      return !p
    })
    this.open.set(false)
    this.hover.set(null)
  }

  /** Hide the unpinned sidebar when the pointer leaves the rail + sidebar (but not when moving between them). */
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
