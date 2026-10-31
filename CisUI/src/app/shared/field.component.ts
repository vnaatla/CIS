import { ChangeDetectionStrategy, Component, input } from '@angular/core'

/**
 * Label + control + hint / error (GENESIS components/Field). The label wraps the projected control, so the control is announced with
 * its label: `<app-field label="Payor" [error]="err"><select ...></select></app-field>`.
 */
@Component({
  selector: 'app-field',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <label class="form-field" [class.wide]="wide()" [class.invalid]="!!error()">
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
  readonly required = input(false)
}
