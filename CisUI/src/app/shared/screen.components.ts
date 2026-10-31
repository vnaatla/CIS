import { ChangeDetectionStrategy, Component, input } from '@angular/core'

let headingSeq = 0

/** Standard screen frame (GENESIS PageShell): heading + the card container all screens share. */
@Component({
  selector: 'app-page-shell',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <section class="registration case-reg screen" [attr.aria-labelledby]="headingId">
      <header class="screen-head">
        <h2 [id]="headingId">{{ title() }}</h2>
        @if (subtitle()) {
          <p class="muted">{{ subtitle() }}</p>
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
