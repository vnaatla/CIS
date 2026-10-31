import { provideHttpClient } from '@angular/common/http'
import { HttpTestingController, provideHttpClientTesting } from '@angular/common/http/testing'
import { TestBed } from '@angular/core/testing'
import { provideRouter } from '@angular/router'
import { LoginPage } from './login.page'

describe('LoginPage (frmLogin)', () => {
  beforeEach(() => {
    localStorage.clear()
    TestBed.configureTestingModule({ providers: [provideRouter([]), provideHttpClient(), provideHttpClientTesting()] })
  })

  it('asks for the user id and password before opening the change-password panel (cmdChange_Click)', async () => {
    const fixture = TestBed.createComponent(LoginPage)
    await fixture.whenStable()
    const el = fixture.nativeElement as HTMLElement
    Array.from(el.querySelectorAll('button')).find((b) => b.textContent?.includes('Change Password'))!.click()
    await fixture.whenStable()
    expect(el.querySelector('.error')?.textContent).toContain('Please Enter User ID and Password')
    expect(el.querySelector('fieldset.change-password')).toBeNull()
  })

  it('shows the API message when the login is refused', async () => {
    const fixture = TestBed.createComponent(LoginPage)
    await fixture.whenStable()
    const el = fixture.nativeElement as HTMLElement
    const [user] = Array.from(el.querySelectorAll<HTMLInputElement>('input'))
    user.value = 'NOBODY'
    user.dispatchEvent(new Event('input'))
    el.querySelector('form')!.dispatchEvent(new Event('submit'))

    TestBed.inject(HttpTestingController)
      .expectOne('/api/auth/login')
      .flush({ message: 'Invalid User Name!! Please enter again..' }, { status: 401, statusText: 'Unauthorized' })
    await new Promise((r) => setTimeout(r))
    await fixture.whenStable()
    expect(el.querySelector('.error')?.textContent).toContain('Invalid User Name!! Please enter again..')
  })
})
