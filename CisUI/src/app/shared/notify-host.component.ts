import { ChangeDetectionStrategy, Component, inject } from '@angular/core'
import { NotifyService } from '../core/notify/notify.service'
import { ModalComponent } from './modal.component'

/** Mounted once by the root component: the toast stack and the pending Yes/No confirmation. */
@Component({
  selector: 'app-notify-host',
  imports: [ModalComponent],
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
      <app-modal title="Please confirm" (closed)="notify.answer(false)">
        <p class="confirm-text">{{ pending.text }}</p>
        <div class="modal-actions">
          <button type="button" data-autofocus (click)="notify.answer(true)">Yes</button>
          <button type="button" class="link" (click)="notify.answer(false)">No</button>
        </div>
      </app-modal>
    }
  `,
})
export class NotifyHostComponent {
  protected readonly notify = inject(NotifyService)
}
