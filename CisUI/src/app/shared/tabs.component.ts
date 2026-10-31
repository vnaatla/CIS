import { ChangeDetectionStrategy, Component, Directive, ElementRef, TemplateRef, contentChildren, inject, input, model, viewChildren } from '@angular/core'
import { NgTemplateOutlet } from '@angular/common'

/** One tab: `<ng-template appTab="Plan">...</ng-template>`. */
@Directive({ selector: 'ng-template[appTab]' })
export class TabDirective {
  readonly appTab = input.required<string>()
  readonly template = inject<TemplateRef<unknown>>(TemplateRef)
}

/**
 * Tabbed card (GENESIS Tabs, VB6 SSTab). Only the active panel is rendered. Left / Right / Home / End move between tabs.
 * Bind `[(active)]` to control it from the page.
 */
@Component({
  selector: 'app-tabs',
  imports: [NgTemplateOutlet],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="tabbed">
      <div class="tabs" role="tablist" [attr.aria-label]="label()" (keydown)="onKey($event)">
        @for (tab of tabs(); track $index; let i = $index) {
          <button
            #tabButton
            type="button"
            role="tab"
            [attr.aria-selected]="i === active()"
            [attr.tabindex]="i === active() ? 0 : -1"
            [class.active]="i === active()"
            (click)="select(i)"
          >
            {{ tab.appTab() }}
          </button>
        }
      </div>
      <div class="tab-body" role="tabpanel">
        @if (tabs()[active()]; as tab) {
          <ng-container [ngTemplateOutlet]="tab.template" />
        }
      </div>
    </div>
  `,
})
export class TabsComponent {
  readonly label = input.required<string>()
  readonly active = model(0)
  protected readonly tabs = contentChildren(TabDirective)
  private readonly buttons = viewChildren<ElementRef<HTMLButtonElement>>('tabButton')

  protected select(i: number): void {
    this.active.set(i)
    this.buttons()[i]?.nativeElement.focus()
  }

  protected onKey(e: KeyboardEvent): void {
    const last = this.tabs().length - 1
    const current = this.active()
    const next = ({ ArrowRight: current === last ? 0 : current + 1, ArrowLeft: current === 0 ? last : current - 1, Home: 0, End: last } as Record<string, number>)[e.key]
    if (next === undefined) return
    e.preventDefault()
    this.select(next)
  }
}
