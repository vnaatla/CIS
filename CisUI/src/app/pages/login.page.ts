import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { Router } from '@angular/router'
import { ScreenAuditService } from '../core/audit/screen-audit.service'
import { errorMessage } from '../core/http/errors'
import { NotifyService } from '../core/notify/notify.service'
import { SessionService } from '../core/auth/session.service'
import { ButtonDirective } from '../shared/button.directive'
import { IconComponent } from '../shared/icon.component'

/**
 * frmLogin: user name + password, Enter in the user name moves to the password, Enter in the password logs in. "Change Password" opens
 * the new / re-type fields under the login box (cmdChange / cmdSave). Messages are the desktop's.
 */
@Component({
  selector: 'app-login-page',
  imports: [FormsModule, ButtonDirective, IconComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="login">
      <aside class="login-brand" aria-hidden="true">
        <div class="login-brand-inner">
          <div class="login-logo"><app-icon name="cross" [size]="30" /></div>
          <h1>CIS</h1>
          <p class="login-tagline">MediExpress Clinical Information System</p>
          <ul class="login-features">
            <li><span><app-icon name="users" /></span> Membership, cases and invoices</li>
            <li><span><app-icon name="file" /></span> Claims, bordereaux and payments</li>
            <li><span><app-icon name="activity" /></span> Postings, tracking and reports</li>
          </ul>
        </div>
        <p class="login-brand-foot">Replaces the MedixCIS desktop</p>
      </aside>

      <main class="login-panel">
        <form class="card login-card" (ngSubmit)="submit()">
          <div class="login-head">
            <div class="login-logo small"><app-icon name="cross" [size]="20" /></div>
            <div>
              <h2>{{ changing() ? 'Change password' : 'Sign in' }}</h2>
              <p class="muted">{{ changing() ? 'Your user name and current password, then the new one' : 'Use your CIS user name and password' }}</p>
            </div>
          </div>

          <label class="login-field">
            <span>User Name</span>
            <span class="login-input">
              <app-icon name="user" [size]="16" />
              <input #user name="userCode" autocomplete="username" maxlength="10" autofocus placeholder="User name" [(ngModel)]="userCode"
                (keydown.enter)="$event.preventDefault(); password.focus()" />
            </span>
          </label>
          <label class="login-field">
            <span>Password</span>
            <span class="login-input">
              <app-icon name="lock" [size]="16" />
              <input #password name="password" [type]="showPassword() ? 'text' : 'password'" autocomplete="current-password" maxlength="10" placeholder="Password"
                [(ngModel)]="passwordText" (keyup)="capsCheck($event)" (keydown)="capsCheck($event)" />
              <button type="button" class="login-eye" [attr.aria-label]="showPassword() ? 'Hide password' : 'Show password'" [attr.aria-pressed]="showPassword()"
                (click)="showPassword.set(!showPassword())">
                <app-icon [name]="showPassword() ? 'eyeOff' : 'eye'" [size]="16" />
              </button>
            </span>
          </label>
          @if (capsLock()) {
            <p class="login-caps" role="status"><app-icon name="alert" [size]="14" /> Caps Lock is on</p>
          }

          @if (changing()) {
            <fieldset class="change-password">
              <legend>Please enter a new Password:</legend>
              <label class="login-field">
                <span>New Password</span>
                <span class="login-input">
                  <app-icon name="lock" [size]="16" />
                  <input name="newPassword" type="password" autocomplete="new-password" maxlength="10" [(ngModel)]="newPassword" />
                </span>
              </label>
              <label class="login-field">
                <span>Re-Type New Password</span>
                <span class="login-input">
                  <app-icon name="lock" [size]="16" />
                  <input name="confirmPassword" type="password" autocomplete="new-password" maxlength="10" [(ngModel)]="confirmPassword" />
                </span>
              </label>
            </fieldset>
          }

          @if (error()) {
            <p class="error login-message" role="alert"><app-icon name="alert" [size]="16" /> {{ error() }}</p>
          }
          @if (info()) {
            <p class="success login-message" role="status"><app-icon name="check" [size]="16" /> {{ info() }}</p>
          }

          <div class="login-actions">
            @if (changing()) {
              <button appButton variant="primary" type="submit" [busy]="busy()">{{ busy() ? 'Saving…' : 'Save' }}</button>
              <button appButton variant="link" (click)="cancelChange()">Cancel</button>
            } @else {
              <button appButton variant="primary" type="submit" [busy]="busy()">{{ busy() ? 'Signing in…' : 'Sign in' }}</button>
              <button appButton variant="link" (click)="startChange()">Change Password</button>
            }
          </div>
        </form>
        <p class="login-foot">Authorised users only. Your sign-in is recorded.</p>
      </main>
    </div>
  `,
})
export class LoginPage {
  private readonly session = inject(SessionService)
  private readonly router = inject(Router)
  private readonly notify = inject(NotifyService)
  private readonly audit = inject(ScreenAuditService)

  protected userCode = ''
  protected passwordText = ''
  protected newPassword = ''
  protected confirmPassword = ''
  protected readonly changing = signal(false)
  protected readonly busy = signal(false)
  protected readonly error = signal('')
  protected readonly info = signal('')
  protected readonly showPassword = signal(false)
  protected readonly capsLock = signal(false)

  protected capsCheck(e: KeyboardEvent): void {
    if (typeof e.getModifierState === 'function') this.capsLock.set(e.getModifierState('CapsLock'))
  }

  /** cmdChange_Click: needs a user name or password first. */
  protected startChange(): void {
    this.clearMessages()
    if (!this.userCode && !this.passwordText) {
      this.error.set('Please Enter User ID and Password')
      return
    }
    this.changing.set(true)
  }

  protected cancelChange(): void {
    this.changing.set(false)
    this.newPassword = this.confirmPassword = ''
    this.clearMessages()
  }

  protected async submit(): Promise<void> {
    this.clearMessages()
    this.busy.set(true)
    try {
      if (this.changing()) {
        const message = await this.session.changePassword(this.userCode, this.passwordText, this.newPassword, this.confirmPassword)
        this.notify.notify(message, 'success')
        this.info.set(message)
        this.changing.set(false)
        this.passwordText = this.newPassword = this.confirmPassword = ''
        this.showPassword.set(false)
      } else {
        await this.session.login(this.userCode, this.passwordText)
        this.audit.screenOpened('MDIMain')
        await this.router.navigateByUrl('/')
      }
    } catch (e) {
      this.error.set(errorMessage(e))
      this.passwordText = ''
    } finally {
      this.busy.set(false)
    }
  }

  private clearMessages(): void {
    this.error.set('')
    this.info.set('')
  }
}
