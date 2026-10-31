import { ChangeDetectionStrategy, Component, forwardRef, input, signal } from '@angular/core'
import { ControlValueAccessor, NG_VALUE_ACCESSOR } from '@angular/forms'
import { IconComponent } from './icon.component'
import { dmyToIso, isoToDmy, maskDate } from './lib/masks'

/**
 * Masked dd/mm/yyyy text box with a calendar picker (GENESIS components/DateInput, VB6 MaskEdBox "##/##/####").
 * The form value stays dd/mm/yyyy text, as in the desktop; convert with `dmyToIso` before sending it to the API.
 */
@Component({
  selector: 'app-date-input',
  imports: [IconComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  providers: [{ provide: NG_VALUE_ACCESSOR, useExisting: forwardRef(() => DateInputComponent), multi: true }],
  template: `
    <span class="date-input">
      <input
        [value]="value()"
        placeholder="dd/mm/yyyy"
        maxlength="10"
        inputmode="numeric"
        [attr.aria-label]="ariaLabel()"
        [attr.aria-invalid]="invalid() || null"
        [disabled]="disabled()"
        (input)="typed($any($event.target))"
        (blur)="onTouched()"
      />
      @if (!disabled()) {
        <span class="cal-wrap" title="Pick a date">
          <app-icon name="calendar" [size]="14" />
          <input type="date" [attr.aria-label]="(ariaLabel() ?? 'Date') + ' picker'" [value]="iso()" (change)="picked($any($event.target).value)" />
        </span>
      }
    </span>
  `,
  host: { style: 'display: block' },
})
export class DateInputComponent implements ControlValueAccessor {
  readonly ariaLabel = input<string | undefined>()
  readonly invalid = input(false)
  protected readonly value = signal('')
  protected readonly disabled = signal(false)
  private onChange: (v: string) => void = () => undefined
  protected onTouched: () => void = () => undefined

  protected iso(): string {
    return dmyToIso(this.value())
  }

  protected typed(el: HTMLInputElement): void {
    const masked = maskDate(el.value)
    el.value = masked
    this.set(masked)
  }

  protected picked(iso: string): void {
    if (iso) this.set(isoToDmy(iso))
  }

  private set(v: string): void {
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
