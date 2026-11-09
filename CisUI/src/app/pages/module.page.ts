import { ChangeDetectionStrategy, Component, computed, inject, input } from '@angular/core'
import { RouterLink } from '@angular/router'
import { SessionService } from '../core/auth/session.service'
import { DESKTOP_STATE_TEXT, MENU, MenuItem, nothingToBuild } from '../layout/menu'
import { PageShellComponent } from '../shared/screen.components'

interface Group {
  name: string
  items: MenuItem[]
}

/**
 * Module overview (route `/:module`): every desktop menu item of the module, the VB6 form it replaces, the access position, how it
 * appears in the desktop and the migration status. The traceability view of docs/menu-mapping.md inside the app.
 */
@Component({
  selector: 'app-module-page',
  imports: [RouterLink, PageShellComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    @if (module(); as m) {
      <app-page-shell [title]="m.label" subtitle="Desktop menu items of this module and their migration status">
        <div class="module-items">
          @for (g of groups(); track g.name) {
            <fieldset>
              <legend>{{ g.name }}</legend>
              <div class="table-wrap">
                <table class="grid" [attr.aria-label]="m.label + ' — ' + g.name">
                  <thead>
                    <tr>
                      <th scope="col">Menu item</th>
                      <th scope="col">VB6 form</th>
                      <th scope="col" class="num">Access</th>
                      <th scope="col">Desktop</th>
                      <th scope="col">Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    @for (i of g.items; track i.path) {
                      <tr>
                        <td>
                          @if (i.migrated && session.hasAccess(i.access)) {
                            <a [routerLink]="i.path">{{ i.label }}</a>
                          } @else {
                            {{ i.label }}
                          }
                        </td>
                        <td class="form">{{ i.form }}</td>
                        <td class="num">{{ i.access }}</td>
                        <td>{{ desktop(i) }}</td>
                        <td>
                          @if (nothingToBuild(i)) {
                            <span class="badge blue">Nothing to build</span>
                          } @else if (i.migrated) {
                            <span class="badge green">Migrated</span>
                          } @else {
                            <span class="badge amber">Phase {{ i.phase }}</span>
                          }
                          @if (!session.hasAccess(i.access)) {
                            <span class="badge red" title="Your USRAccess does not have this position">No access</span>
                          }
                        </td>
                      </tr>
                    }
                  </tbody>
                </table>
              </div>
            </fieldset>
          }
        </div>
      </app-page-shell>
    } @else {
      <section>
        <h2>Page not found</h2>
        <p class="muted">There is no CIS module "{{ moduleKey() }}".</p>
      </section>
    }
  `,
})
export class ModulePage {
  /** Route parameter (component input binding). */
  readonly moduleKey = input.required<string>({ alias: 'module' })
  protected readonly session = inject(SessionService)
  protected readonly module = computed(() => MENU.find((m) => m.key === this.moduleKey()))
  protected readonly groups = computed<Group[]>(() => {
    const groups: Group[] = []
    for (const item of this.module()?.items ?? []) {
      const name = item.group ?? this.module()!.label
      const group = groups.find((g) => g.name === name) ?? groups[groups.push({ name, items: [] }) - 1]
      group.items.push(item)
    }
    return groups
  })

  protected readonly nothingToBuild = nothingToBuild

  protected desktop(i: MenuItem): string {
    return DESKTOP_STATE_TEXT[i.desktop ?? 'visible']
  }
}
