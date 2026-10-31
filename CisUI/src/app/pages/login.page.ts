import { ChangeDetectionStrategy, Component, inject, signal } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { Router } from '@angular/router'
import { ScreenAuditService } from '../core/audit/screen-audit.service'
import { errorMessage } from '../core/http/errors'
import { NotifyService } from '../core/notify/notify.service'
import { SessionService } from '../core/auth/session.service'
import { ButtonDirective } from '../shared/button.directive'

/**
 * frmLogin: user name + password, Enter in the user name moves to the password, Enter in the password logs in. "Change Password" opens
 * the new / re-type fields under the login box (cmdChange / cmdSave). Messages are the desktop's.
 */
@Component({
  selector: 'app-login-page',
  imports: [FormsModule, ButtonDirective],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="login">
      <form class="card" (ngSubmit)="submit()">
        <h2>CIS</h2>
        <p class="muted">MediExpress Clinical Information System</p>
        <label>
          User Name
          <input #user name="userCode" autocomplete="username" maxlength="10" autofocus [(ngModel)]="userCode" (keydown.enter)="$event.preventDefault(); password.focus()" />
        </label>
        <label>
          Password
          <input #password name="password" type="password" autocomplete="current-password" maxlength="10" [(ngModel)]="passwordText" />
        </label>

        @if (changing()) {
          <fieldset class="change-password">
            <legend>Please enter a new Password:</legend>
            <label>
              New Password
              <input name="newPassword" type="password" autocomplete="new-password" maxlength="10" [(ngModel)]="newPassword" />
            </label>
            <label>
              Re-Type New Password
              <input name="confirmPassword" type="password" autocomplete="new-password" maxlength="10" [(ngModel)]="confirmPassword" />
            </label>
          </fieldset>
        }

        @if (error()) {
          <p class="error" role="alert">{{ error() }}</p>
        }
        @if (info()) {
          <p class="success" role="status">{{ info() }}</p>
        }

        <div class="card-actions">
          @if (changing()) {
            <button appButton variant="primary" type="submit" [busy]="busy()">{{ busy() ? 'Saving…' : 'Save' }}</button>
            <span class="spacer"></span>
            <button appButton variant="link" (click)="cancelChange()">Cancel</button>
          } @else {
            <button appButton variant="primary" type="submit" [busy]="busy()">{{ busy() ? 'Signing in…' : 'OK' }}</button>
            <span class="spacer"></span>
            <button appButton variant="link" (click)="startChange()">Change Password</button>
          }
        </div>
      </form>
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
