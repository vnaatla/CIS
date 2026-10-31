import { Directive, computed, input } from '@angular/core'

export type ButtonVariant = 'default' | 'primary' | 'ghost' | 'link' | 'danger'

/**
 * The app's button (GENESIS components/Button): `<button appButton variant="primary" [busy]="saving()">Save</button>`.
 * Defaults to type="button" so a button inside a form never submits it by accident; set type="submit" for the form's own action.
 * While `busy`, the button is disabled and announced with aria-busy.
 */
@Directive({
  selector: 'button[appButton]',
  host: {
    '[attr.type]': 'type()',
    '[class]': 'cssClass()',
    '[disabled]': 'disabled() || busy()',
    '[attr.aria-busy]': 'busy() || null',
  },
})
export class ButtonDirective {
  readonly variant = input<ButtonVariant>('default')
  readonly type = input<'button' | 'submit' | 'reset'>('button')
  readonly busy = input(false)
  readonly disabled = input(false)
  protected readonly cssClass = computed(() => (this.variant() === 'default' ? '' : this.variant()))
}
