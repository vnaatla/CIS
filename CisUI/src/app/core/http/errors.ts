import { HttpErrorResponse } from '@angular/common/http'

/** The text to show for a failed request: the API's Problem Details `message` (the desktop's MsgBox text) when there is one. */
export function errorMessage(error: unknown): string {
  if (error instanceof HttpErrorResponse) {
    const body = error.error as { message?: string; title?: string } | string | null
    if (body && typeof body === 'object') {
      if (body.message) return body.message
      if (body.title) return body.title
    }
    if (typeof body === 'string' && body.length > 0 && body.length < 300) return body
    if (error.status === 0) return 'Cannot reach the CIS API. Is it running?'
    if (error.status === 504) return 'The request took too long. Try again, or narrow the search.'
    return `Request failed (${error.status}).`
  }
  return error instanceof Error ? error.message : 'Unexpected error'
}

/** For downloads (responseType 'blob'): the Problem Details body arrives as a Blob; read its message. */
export async function blobErrorMessage(error: unknown): Promise<string> {
  if (error instanceof HttpErrorResponse && error.error instanceof Blob) {
    const text = await error.error.text()
    try {
      const body = JSON.parse(text) as { message?: string; title?: string }
      return body.message ?? body.title ?? text
    } catch {
      return text
    }
  }
  return errorMessage(error)
}
