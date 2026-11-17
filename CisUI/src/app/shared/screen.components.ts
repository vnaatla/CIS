import { ChangeDetectionStrategy, Component, computed, inject, input } from '@angular/core'
import { Router, RouterLink } from '@angular/router'
import { MENU } from '../layout/menu'
import { IconComponent } from './icon.component'

let headingSeq = 0

/**
 * Standard screen frame (GENESIS PageShell): a header card (the module's icon, Module › sub-menu breadcrumb, title, the replaced VB6
 * form as a badge) + the card container all screens share. The module and sub-menu come from the menu entry of the current route.
 */
@Component({
  selector: 'app-page-shell',
  imports: [RouterLink, IconComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <section class="registration case-reg screen" [attr.aria-labelledby]="headingId">
      <header class="screen-head">
        @if (place(); as p) {
          <span class="screen-head-icon" aria-hidden="true"><app-icon [name]="p.module.icon" [size]="22" /></span>
        }
        <div class="screen-head-text">
          @if (place(); as p) {
            <nav class="screen-crumbs" aria-label="Breadcrumb">
              <a routerLink="/">Home</a>
              <app-icon name="chevRight" [size]="12" />
              <a [routerLink]="p.module.path">{{ p.module.label }}</a>
              @if (p.group && p.group !== p.module.label) {
                <app-icon name="chevRight" [size]="12" />
                <span>{{ p.group }}</span>
              }
            </nav>
          }
          <h2 [id]="headingId">{{ title() }}</h2>
        </div>
        @if (subtitle()) {
          <p class="muted screen-head-badge"><app-icon name="file" [size]="13" /> {{ subtitle() }}</p>
        }
      </header>
      <ng-content />
    </section>
  `,
})
export class PageShellComponent {
  readonly title = input.required<string>()
  readonly subtitle = input<string | undefined>()
  protected readonly headingId = `page-heading-${++headingSeq}`
  private readonly url = inject(Router).url.split('?')[0]

  /** The menu module and sub-menu of this screen (the longest menu path the route starts with). */
  protected readonly place = computed(() => {
    let best: { module: (typeof MENU)[number]; group?: string; len: number } | undefined
    for (const m of MENU)
      for (const i of m.items)
        if ((this.url === i.path || this.url.startsWith(i.path + '/')) && i.path.length > (best?.len ?? 0)) best = { module: m, group: i.group, len: i.path.length }
    return best ?? (MENU.find((m) => this.url.startsWith(m.path)) ? { module: MENU.find((m) => this.url.startsWith(m.path))!, len: 0 } : undefined)
  })
}

/** Sticky command bar (GENESIS ActionBar; the VB6 command-button frame). Use <span class="spacer"></span> to split groups. */
@Component({
  selector: 'app-action-bar',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<div class="actions cr-actions" role="toolbar" [attr.aria-label]="label()"><ng-content /></div>`,
})
export class ActionBarComponent {
  readonly label = input.required<string>()
}

/** Small read-only counter (GENESIS Stat; VB6 "Records" / "Total Amt" labels). */
@Component({
  selector: 'app-stat',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<span class="stat"><span class="stat-label">{{ label() }}</span><strong class="stat-value">{{ value() }}</strong></span>`,
})
export class StatComponent {
  readonly label = input.required<string>()
  readonly value = input.required<string | number>()
}
