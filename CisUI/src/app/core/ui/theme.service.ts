import { Injectable, signal } from '@angular/core'

/** Browser key of the chosen theme (index.html applies it before the app starts). */
export const THEME_KEY = 'cis.theme'

export interface ThemeOption {
  id: string
  label: string
  /** Swatch colours shown in the menu: strong colour and accent. */
  swatch: [string, string]
}

/** Purple is the designed default; the navy colours come from the THEME TOKENS block in styles.css (scripts/theme-tokens.cjs). */
export const THEMES: ThemeOption[] = [
  { id: 'purple', label: 'Purple', swatch: ['#7c2172', '#9c2b8e'] },
  { id: 'navy', label: 'Navy blue', swatch: ['#0b2038', '#1f4e8c'] },
]
const DEFAULT = 'purple'

/** The colour theme of this browser: sets <html data-theme> (none for purple) and remembers the choice in localStorage. */
@Injectable({ providedIn: 'root' })
export class ThemeService {
  readonly themes = THEMES
  readonly theme = signal(DEFAULT)

  constructor() {
    this.apply(this.saved())
  }

  /** Switches to a theme and remembers it in this browser. */
  set(id: string): void {
    const theme = this.valid(id)
    this.apply(theme)
    try {
      if (theme === DEFAULT) localStorage.removeItem(THEME_KEY)
      else localStorage.setItem(THEME_KEY, theme)
    } catch {
      // storage unavailable (private mode): the theme still applies for this page
    }
  }

  private saved(): string {
    try {
      return this.valid(localStorage.getItem(THEME_KEY))
    } catch {
      return DEFAULT
    }
  }

  private valid(id: string | null): string {
    return THEMES.some((t) => t.id === id) ? (id as string) : DEFAULT
  }

  private apply(id: string): void {
    this.theme.set(id)
    const root = document.documentElement
    if (id === DEFAULT) delete root.dataset['theme']
    else root.dataset['theme'] = id
    const option = THEMES.find((t) => t.id === id) ?? THEMES[0]
    let meta = document.querySelector<HTMLMetaElement>('meta[name="theme-color"]')
    if (!meta) {
      meta = document.createElement('meta')
      meta.name = 'theme-color'
      document.head.appendChild(meta)
    }
    meta.content = option.swatch[0]
  }
}
