import { ChangeDetectionStrategy, Component, forwardRef, input, signal } from '@angular/core'
import { ControlValueAccessor, NG_VALUE_ACCESSOR } from '@angular/forms'

export interface LookupOption {
  code: string
  name: string
}

/** Drop-down of code / name options shown as "code - name" (just the name when both are the same), like the desktop's combos. */
@Component({
  selector: 'app-code-select',
  changeDetection: ChangeDetectionStrategy.OnPush,
  providers: [{ provide: NG_VALUE_ACCESSOR, useExisting: forwardRef(() => CodeSelectComponent), multi: true }],
  template: `
    <select [value]="value()" [disabled]="disabled()" [attr.aria-label]="ariaLabel()" (change)="select($any($event.target).value)" (blur)="onTouched()">
      @if (allowBlank()) {
        <option value=""></option>
      }
      @for (o of options(); track o.code) {
        <option [value]="o.code" [selected]="o.code === value()">{{ o.code === o.name ? o.name : o.code + ' - ' + o.name }}</option>
      }
    </select>
  `,
  host: { style: 'display: contents' },
})
export class CodeSelectComponent implements ControlValueAccessor {
  readonly options = input<readonly LookupOption[]>([])
  readonly allowBlank = input(true)
  readonly ariaLabel = input<string | undefined>()
  protected readonly value = signal('')
  protected readonly disabled = signal(false)
  private onChange: (v: string) => void = () => undefined
  protected onTouched: () => void = () => undefined

  protected select(v: string): void {
    this.value.set(v)
    this.onChange(v)
  }

  writeValue(v: string | null): void {
    this.value.set(v ?? '')
  }

  registerOnChange(fn: (v: string) => void): void {
    this.onChange = fn
  }

  registerOnTouched(fn: () => void): void {
    this.onTouched = fn
  }

  setDisabledState(disabled: boolean): void {
    this.disabled.set(disabled)
  }
}
