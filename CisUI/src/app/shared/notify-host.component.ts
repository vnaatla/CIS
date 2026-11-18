import { ChangeDetectionStrategy, Component, inject } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { NotifyService } from '../core/notify/notify.service'
import { ButtonDirective } from './button.directive'
import { ModalComponent } from './modal.component'

/** Mounted once by the root component: the toast stack, the pending Yes/No confirmation and the pending text question. */
@Component({
  selector: 'app-notify-host',
  imports: [FormsModule, ModalComponent, ButtonDirective],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="toasts" aria-live="polite">
      @for (n of notify.notices(); track n.id) {
        <div [class]="'toast ' + n.kind" [attr.role]="n.kind === 'error' ? 'alert' : 'status'">
          <span>{{ n.text }}</span>
          <button type="button" aria-label="Dismiss" (click)="notify.dismiss(n.id)">✕</button>
        </div>
      }
    </div>
    @if (notify.pending(); as pending) {
      <app-modal title="Please confirm" icon="alert" (closed)="notify.answer(false)">
        <p class="confirm-text">{{ pending.text }}</p>
        <div class="modal-actions">
          <button appButton variant="primary" data-autofocus (click)="notify.answer(true)">Yes</button>
          <button appButton variant="ghost" (click)="notify.answer(false)">No</button>
        </div>
      </app-modal>
    }
    @if (notify.prompting(); as p) {
      <app-modal title="Please enter" icon="file" (closed)="notify.reply(null)">
        <p class="confirm-text">{{ p.text }}</p>
        <label class="form-field prompt-field">
          <span>{{ p.label }}</span>
          <input name="promptValue" data-autofocus [(ngModel)]="p.value" (keydown.enter)="notify.reply(p.value)" />
        </label>
        <div class="modal-actions">
          <button appButton variant="primary" (click)="notify.reply(p.value)">OK</button>
          <button appButton variant="ghost" (click)="notify.reply(null)">Cancel</button>
        </div>
      </app-modal>
    }
  `,
})
export class NotifyHostComponent {
  protected readonly notify = inject(NotifyService)
}
