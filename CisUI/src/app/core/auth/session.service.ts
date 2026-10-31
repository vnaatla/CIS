import { HttpClient } from '@angular/common/http'
import { Injectable, computed, inject, signal } from '@angular/core'
import { firstValueFrom } from 'rxjs'
import { hasPosition } from './access'

export interface SessionUser {
  userCode: string
  userName: string
  access: string
}

export interface Session {
  token: string
  expiresUtc: string
  user: SessionUser
}

export const SESSION_KEY = 'cis.session'

/** The stored session, or null when there is none or it has expired. */
export function loadSession(storage: Storage = localStorage): Session | null {
  const raw = storage.getItem(SESSION_KEY)
  if (!raw) return null
  try {
    const session = JSON.parse(raw) as Session
    return new Date(session.expiresUtc) > new Date() ? session : null
  } catch {
    return null
  }
}

/** Signed-in user, token and menu rights (replaces the desktop's global Logon / Access variables). */
@Injectable({ providedIn: 'root' })
export class SessionService {
  private readonly http = inject(HttpClient)
  private readonly state = signal<Session | null>(loadSession())

  readonly session = this.state.asReadonly()
  readonly user = computed(() => this.state()?.user ?? null)
  readonly signedIn = computed(() => this.state() !== null)

  get token(): string | null {
    return this.state()?.token ?? null
  }

  /** VB6 CheckAccess(n). */
  hasAccess(position: number): boolean {
    return hasPosition(this.state()?.user.access, position)
  }

  /** frmLogin.cmdOK_Click. Errors carry the desktop's messages (Invalid User Name / Invalid Password / Your Account Been Suspended). */
  async login(userCode: string, password: string): Promise<void> {
    const session = await firstValueFrom(this.http.post<Session>('/api/auth/login', { userCode, password }))
    this.apply(session)
  }

  /** frmLogin.cmdSave_Click; resolves with the desktop's confirmation text. */
  async changePassword(userCode: string, password: string, newPassword: string, confirmPassword: string): Promise<string> {
    const answer = await firstValueFrom(
      this.http.post<{ message: string }>('/api/auth/change-password', { userCode, password, newPassword, confirmPassword }),
    )
    return answer.message
  }

  /** Exit > Logout / Close Connection: tells the API (USR.CIScnt - 1), then forgets the session even if the call fails. */
  async logout(): Promise<void> {
    if (this.state()) {
      try {
        await firstValueFrom(this.http.post('/api/auth/logout', null))
      } catch {
        // Expired token or API down: the local session is cleared anyway.
      }
    }
    this.clear()
  }

  /** Forgets the session without calling the API (401 from the server). */
  clear(): void {
    this.apply(null)
  }

  private apply(session: Session | null): void {
    if (session) localStorage.setItem(SESSION_KEY, JSON.stringify(session))
    else localStorage.removeItem(SESSION_KEY)
    this.state.set(session)
  }
}
