import { ChangeDetectionStrategy, Component, OnInit, inject, input, output } from '@angular/core'
import { FormsModule } from '@angular/forms'
import { NotifyService } from '../../core/notify/notify.service'
import { ButtonDirective } from '../../shared/button.directive'
import type { LookupOption } from '../../shared/code-select.component'
import { FieldComponent } from '../../shared/field.component'
import { ModalComponent } from '../../shared/modal.component'
import { ActionBarComponent } from '../../shared/screen.components'

/** The pending details kept by frmCasePending. */
export interface PendingDetails {
  reasonIds: string[]
  reasonNames: string[]
  questions: string[]
  remark: string
}

/** frmCasePending: pending reasons, questions to the hospital and an additional remark (editable only while editing the case). */
@Component({
  selector: 'app-case-pending-modal',
  imports: [FormsModule, ModalComponent, FieldComponent, ButtonDirective, ActionBarComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <app-modal title="Case Pending" (closed)="cancel()">
      @if (note()) {
        <p class="notice" role="alert">{{ note() }}</p>
      }
      <fieldset aria-label="Pending reasons" [disabled]="!editable()">
        <legend>Pending Reasons</legend>
        @for (r of reasons(); track r.code) {
          <label class="check-field"><input type="checkbox" [checked]="ids.includes(r.code)" (change)="toggle(r)" /> {{ r.name }}</label>
        } @empty {
          <p>No pending reasons to choose.</p>
        }
      </fieldset>
      <fieldset aria-label="Questions" [disabled]="!editable()">
        <legend>Questions to Hospital</legend>
        <app-field label="Question" [wide]="true"><input name="question" [(ngModel)]="question" /></app-field>
        <app-action-bar label="Question actions">
          <button appButton (click)="add()">Add</button>
          <button appButton [disabled]="selected < 0" (click)="remove()">Remove</button>
        </app-action-bar>
        <ul class="pick-list" aria-label="Questions list">
          @for (q of questions; track $index) {
            <li [class.active]="$index === selected" (click)="selected = $index">{{ q }}</li>
          }
        </ul>
        <app-field label="Additional Remark" [wide]="true"><input name="addPendRemark" maxlength="2000" [(ngModel)]="remark" /></app-field>
      </fieldset>
      <app-action-bar label="Pending actions">
        <button appButton variant="primary" [disabled]="!editable()" (click)="save()">Save</button>
        <button appButton (click)="cancel()">Cancel</button>
      </app-action-bar>
    </app-modal>
  `,
})
export class CasePendingModalComponent implements OnInit {
  private readonly notify = inject(NotifyService)
  readonly reasons = input<LookupOption[]>([])
  readonly details = input.required<PendingDetails>()
  readonly editable = input(false)
  readonly note = input('')
  readonly saved = output<PendingDetails>()
  readonly closed = output<void>()
  protected ids: string[] = []
  protected names: string[] = []
  protected questions: string[] = []
  protected remark = ''
  protected question = ''
  protected selected = -1

  ngOnInit(): void {
    const d = this.details()
    this.ids = [...d.reasonIds]
    this.names = [...d.reasonNames]
    this.questions = [...d.questions]
    this.remark = d.remark
  }

  protected toggle(r: LookupOption): void {
    const i = this.ids.indexOf(r.code)
    if (i >= 0) {
      this.ids.splice(i, 1)
      this.names.splice(this.names.indexOf(r.name), 1)
    } else {
      this.ids.push(r.code)
      this.names.push(r.name)
    }
  }

  /** cmdaddreason(0): no duplicates. */
  protected add(): void {
    const q = this.question
    if (q.length >= 1 && !this.questions.includes(q)) this.questions.push(q)
    this.question = ''
  }

  protected remove(): void {
    this.questions.splice(this.selected, 1)
    this.selected = -1
  }

  /** CmdSave_Click. */
  protected save(): void {
    if (this.question.trim()) return this.notify.notify('You have typed the question but not added to the list. Please add it or remove it!', 'info')
    this.saved.emit({ reasonIds: this.ids, reasonNames: this.names, questions: this.questions, remark: this.remark.trim() })
  }

  /** cmdcancel_Click. */
  protected async cancel(): Promise<void> {
    if (this.editable() && !(await this.notify.ask('Your current selection of Pending reasons and Questions will be reset to original values, are you sure?'))) return
    this.closed.emit()
  }
}
