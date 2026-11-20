import { TestBed } from '@angular/core/testing'
import { THEME_KEY, ThemeService } from './theme.service'

describe('ThemeService', () => {
  const root = document.documentElement

  beforeEach(() => {
    localStorage.clear()
    delete root.dataset['theme']
    TestBed.resetTestingModule()
  })

  it('starts on purple without a data-theme attribute', () => {
    const service = TestBed.inject(ThemeService)
    expect(service.theme()).toBe('purple')
    expect(root.dataset['theme']).toBeUndefined()
  })

  it('switches to navy, remembers it, and back to purple clears both', () => {
    const service = TestBed.inject(ThemeService)
    service.set('navy')
    expect(service.theme()).toBe('navy')
    expect(root.dataset['theme']).toBe('navy')
    expect(localStorage.getItem(THEME_KEY)).toBe('navy')
    expect(document.querySelector<HTMLMetaElement>('meta[name="theme-color"]')?.content).toBe('#0b2038')

    service.set('purple')
    expect(root.dataset['theme']).toBeUndefined()
    expect(localStorage.getItem(THEME_KEY)).toBeNull()
  })

  it('applies the theme saved in this browser on start', () => {
    localStorage.setItem(THEME_KEY, 'navy')
    expect(TestBed.inject(ThemeService).theme()).toBe('navy')
    expect(root.dataset['theme']).toBe('navy')
  })

  it('falls back to purple for an unknown saved or requested theme', () => {
    localStorage.setItem(THEME_KEY, 'navy-classic')
    const service = TestBed.inject(ThemeService)
    expect(service.theme()).toBe('purple')
    service.set('green')
    expect(service.theme()).toBe('purple')
    expect(root.dataset['theme']).toBeUndefined()
  })
})
