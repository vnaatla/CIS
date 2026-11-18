import { Injectable, signal } from '@angular/core'

export type NoticeKind = 'info' | 'success' | 'warn' | 'error'

export interface Notice {
  id: number
  kind: NoticeKind
  text: string
}

export interface PendingConfirm {
  text: string
  resolve: (ok: boolean) => void
}

/** A text question (replaces window.prompt / VB6 InputBox): OK returns the text (possibly empty), Cancel returns null. */
export interface PendingPrompt {
  text: string
  label: string
  value: string
  resolve: (value: string | null) => void
}

/** Toasts and Yes/No confirmations (replaces the desktop's MsgBox pop-ups); rendered once by NotifyHostComponent. */
@Injectable({ providedIn: 'root' })
export class NotifyService {
  private seq = 1
  readonly notices = signal<Notice[]>([])
  readonly pending = signal<PendingConfirm | null>(null)

  /** Toast notification; errors stay a little longer. */
  notify(text: string, kind: NoticeKind = 'warn'): void {
    const id = this.seq++
    this.notices.update((list) => [...list, { id, kind, text }])
    setTimeout(() => this.dismiss(id), kind === 'error' ? 9000 : 6000)
  }

  dismiss(id: number): void {
    this.notices.update((list) => list.filter((n) => n.id !== id))
  }

  /** Yes/No confirmation (MsgBox vbYesNo). */
  ask(text: string): Promise<boolean> {
    return new Promise((resolve) => this.pending.set({ text, resolve }))
  }

  answer(ok: boolean): void {
    this.pending()?.resolve(ok)
    this.pending.set(null)
  }

  readonly prompting = signal<PendingPrompt | null>(null)

  /** Text question (InputBox): resolves to the typed text on OK (empty string allowed) or null on Cancel. */
  prompt(text: string, label = 'Value', value = ''): Promise<string | null> {
    return new Promise((resolve) => this.prompting.set({ text, label, value, resolve }))
  }

  reply(value: string | null): void {
    this.prompting()?.resolve(value)
    this.prompting.set(null)
  }
}
