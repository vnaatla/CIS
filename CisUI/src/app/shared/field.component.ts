import { ChangeDetectionStrategy, Component, input } from '@angular/core'

/**
 * Label + control + hint / error (GENESIS components/Field). The label wraps the projected control, so the control is announced with
 * its label: `<app-field label="Payor" [error]="err"><select ...></select></app-field>`.
 */
@Component({
  selector: 'app-field',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <label class="form-field" [class.wide]="wide()" [class.span-2]="span() === 2" [class.invalid]="!!error()">
      <span>
        {{ label() }}
        @if (required()) {
          <abbr class="req" title="Required">*</abbr>
        }
      </span>
      <ng-content />
      @if (hint() && !error()) {
        <small class="hint">{{ hint() }}</small>
      }
      @if (error()) {
        <small class="field-error" role="alert">{{ error() }}</small>
      }
    </label>
  `,
  host: { style: 'display: contents' },
})
export class FieldComponent {
  readonly label = input.required<string>()
  readonly error = input<string | null | undefined>(null)
  readonly hint = input<string | undefined>()
  readonly wide = input(false)
  /** 2 = two grid columns (a long value that should not take the whole row). */
  readonly span = input<1 | 2>(1)
  readonly required = input(false)
}
