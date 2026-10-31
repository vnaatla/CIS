import { HttpClient, HttpParams } from '@angular/common/http'
import { Injectable, inject } from '@angular/core'
import { Observable, shareReplay } from 'rxjs'
import type { LookupOption } from '../shared/code-select.component'

/** Shared drop-down lists (GET /api/lookups/*). Lists that do not depend on a filter are cached for the session. */
@Injectable({ providedIn: 'root' })
export class LookupsService {
  private readonly http = inject(HttpClient)
  private readonly cache = new Map<string, Observable<LookupOption[]>>()

  payors(healthType?: string): Observable<LookupOption[]> {
    return this.cached(`payors:${healthType ?? ''}`, '/api/lookups/payors', { healthType })
  }

  healthTypes(): Observable<LookupOption[]> {
    return this.cached('health-types', '/api/lookups/health-types')
  }

  insuredTypes(): Observable<LookupOption[]> {
    return this.cached('insured-types', '/api/lookups/insured-types')
  }

  groupCompanies(payor?: string): Observable<LookupOption[]> {
    return this.cached(`group-companies:${payor ?? ''}`, '/api/lookups/group-companies', { payor })
  }

  /** Not cached: the clinic list is searched (8,000+ rows). */
  clinics(panelOnly: boolean, search?: string): Observable<LookupOption[]> {
    return this.http.get<LookupOption[]>('/api/lookups/clinics', { params: params({ panelOnly: String(panelOnly), search }) })
  }

  private cached(key: string, url: string, query: Record<string, string | undefined> = {}): Observable<LookupOption[]> {
    let list = this.cache.get(key)
    if (!list) {
      list = this.http.get<LookupOption[]>(url, { params: params(query) }).pipe(shareReplay({ bufferSize: 1, refCount: false }))
      this.cache.set(key, list)
    }
    return list
  }
}

function params(query: Record<string, string | undefined>): HttpParams {
  let p = new HttpParams()
  for (const [k, v] of Object.entries(query)) if (v) p = p.set(k, v)
  return p
}
