import { Injectable } from '@angular/core'

/** Buttons that make a screen a data-entry screen (otherwise it is a search / filter screen). */
const ENTRY_BUTTONS = /^(save|update|edit|add|new|insert|create|amend|register)\b/i

/**
 * Gives every dropdown's blank first option a meaning instead of an empty line: "All" on search / filter screens (blank = no filter)
 * and "-- Select <field> --" on data-entry screens (blank = not chosen yet). The option's value stays empty, so searches, saves and
 * validation are unchanged. Options that already have text (e.g. "--Select--") are left alone.
 */
@Injectable({ providedIn: 'root' })
export class SelectPlaceholderService {
  private scheduled = false

  start(): void {
    // New dropdowns (childList) and form resets such as Clear (Angular switches the controls' ng-* classes).
    new MutationObserver(() => this.schedule()).observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class'] })
    document.addEventListener('change', () => this.schedule(), true)
    this.schedule()
  }

  private schedule(): void {
    if (this.scheduled) return
    this.scheduled = true
    requestAnimationFrame(() => {
      this.scheduled = false
      this.label()
    })
  }

  private label(): void {
    // The blank first option of each dropdown (value "" or a "none" value such as Report Type's -1).
    for (const option of Array.from(document.querySelectorAll<HTMLOptionElement>('select > option:first-child'))) {
      const select = option.parentElement as HTMLSelectElement
      // An unset model (undefined) selects nothing; show the prompt, which means the same (no value) - no change event, the model is untouched.
      if (option.dataset['ph'] && select.selectedIndex === -1) option.selected = true
      if (option.dataset['ph'] || option.textContent?.trim()) continue
      const screen = select.closest('.modal, section.screen') ?? document.body
      const entry = Array.from(screen.querySelectorAll('button')).some((b) => ENTRY_BUTTONS.test(b.textContent?.trim() ?? ''))
      const field = (select.closest('.form-field')?.querySelector(':scope > span')?.textContent ?? select.getAttribute('aria-label') ?? '').replace(/\*/g, '').trim()
      // A required field (or a dropdown that says so with data-prompt) must be chosen, even on a search screen.
      const required = !!select.closest('.form-field')?.querySelector('.req') || select.hasAttribute('data-required')
      const choose = entry || required
      option.textContent = select.dataset['prompt'] ?? (choose ? (field ? `-- Select ${field} --` : '-- Select --') : 'All')
      option.dataset['ph'] = choose ? 'select' : 'all'
      if (select.selectedIndex === -1) option.selected = true
    }
  }
}
