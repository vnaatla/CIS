import { AfterViewInit, ChangeDetectionStrategy, Component, ElementRef, OnDestroy, input, output, viewChild } from '@angular/core'

const FOCUSABLE = 'input, select, textarea, button, [tabindex]'

/**
 * Focus-trapped dialog (GENESIS Modal; replaces the desktop's pop-up forms such as frmSearchMBM, frmSearchICD, frmCasList).
 * Escape closes it; focus returns to where it was when the dialog closes. Put `data-autofocus` on the control to focus first.
 */
@Component({
  selector: 'app-modal',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="modal-overlay" (keydown)="onKey($event)">
      <div #dialog [class]="wide() ? 'modal wide' : 'modal'" role="dialog" aria-modal="true" [attr.aria-label]="title()">
        <div class="modal-head">
          <h3>{{ title() }}</h3>
          <button type="button" class="link" aria-label="Close" (click)="closed.emit()">✕</button>
        </div>
        <div class="modal-body"><ng-content /></div>
      </div>
    </div>
  `,
})
export class ModalComponent implements AfterViewInit, OnDestroy {
  readonly title = input.required<string>()
  readonly wide = input(false)
  readonly closed = output<void>()
  private readonly dialog = viewChild.required<ElementRef<HTMLElement>>('dialog')
  private readonly previous = document.activeElement as HTMLElement | null

  ngAfterViewInit(): void {
    const el = this.dialog().nativeElement
    ;(el.querySelector<HTMLElement>('[data-autofocus]') ?? this.focusables()[0])?.focus()
  }

  ngOnDestroy(): void {
    this.previous?.focus()
  }

  protected onKey(e: KeyboardEvent): void {
    if (e.key === 'Escape') {
      e.stopPropagation()
      this.closed.emit()
      return
    }
    if (e.key !== 'Tab') return
    const items = this.focusables()
    if (items.length === 0) return
    const first = items[0]
    const last = items[items.length - 1]
    const active = document.activeElement as HTMLElement
    if (e.shiftKey ? active === first || !this.dialog().nativeElement.contains(active) : active === last) {
      e.preventDefault()
      ;(e.shiftKey ? last : first).focus()
    }
  }

  private focusables(): HTMLElement[] {
    return Array.from(this.dialog().nativeElement.querySelectorAll<HTMLElement>(FOCUSABLE)).filter((x) => !x.hasAttribute('disabled') && x.tabIndex >= 0)
  }
}
