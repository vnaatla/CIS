import { HttpErrorResponse, provideHttpClient } from '@angular/common/http'
import { HttpTestingController, provideHttpClientTesting } from '@angular/common/http/testing'
import { TestBed } from '@angular/core/testing'
import { errorMessage } from '../http/errors'
import { SESSION_KEY, Session, SessionService, loadSession } from './session.service'

const session = (hoursAhead = 1): Session => ({
  token: 'jwt',
  expiresUtc: new Date(Date.now() + hoursAhead * 3_600_000).toISOString(),
  user: { userCode: 'DEVADMIN', userName: 'Dev Admin', access: '0'.repeat(112) + '1' },
})

describe('SessionService', () => {
  let http: HttpTestingController
  let service: SessionService

  beforeEach(() => {
    localStorage.clear()
    TestBed.configureTestingModule({ providers: [provideHttpClient(), provideHttpClientTesting()] })
    http = TestBed.inject(HttpTestingController)
    service = TestBed.inject(SessionService)
  })

  afterEach(() => http.verify())

  it('stores the session after login and answers CheckAccess from the access flags', async () => {
    const login = service.login('devadmin', 'Gen12345')
    const req = http.expectOne('/api/auth/login')
    expect(req.request.body).toEqual({ userCode: 'devadmin', password: 'Gen12345' })
    req.flush(session())
    await login

    expect(service.signedIn()).toBe(true)
    expect(service.hasAccess(113)).toBe(true)
    expect(service.hasAccess(114)).toBe(false)
    expect(loadSession()?.user.userCode).toBe('DEVADMIN')
  })

  it('tells the API on logout and clears the session even when that call fails', async () => {
    localStorage.setItem(SESSION_KEY, JSON.stringify(session()))
    TestBed.resetTestingModule()
    TestBed.configureTestingModule({ providers: [provideHttpClient(), provideHttpClientTesting()] })
    http = TestBed.inject(HttpTestingController)
    service = TestBed.inject(SessionService)
    expect(service.signedIn()).toBe(true)

    const logout = service.logout()
    http.expectOne('/api/auth/logout').flush(null, { status: 401, statusText: 'Unauthorized' })
    await logout
    expect(service.signedIn()).toBe(false)
    expect(localStorage.getItem(SESSION_KEY)).toBeNull()
  })

  it('ignores an expired stored session', () => {
    localStorage.setItem(SESSION_KEY, JSON.stringify(session(-1)))
    expect(loadSession()).toBeNull()
  })

  it('shows the desktop message the API answered', async () => {
    const login = service.login('x', 'y')
    http.expectOne('/api/auth/login').flush({ message: 'Invalid User Name!! Please enter again..' }, { status: 401, statusText: 'Unauthorized' })
    await expect(login).rejects.toBeInstanceOf(HttpErrorResponse)
    await login.catch((e: unknown) => expect(errorMessage(e)).toBe('Invalid User Name!! Please enter again..'))
  })
})
