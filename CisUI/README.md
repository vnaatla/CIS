# CisUI: Angular front end of the CIS migration

Angular 22 (standalone components, signals, lazy routes) with the GENESIS look: `src/styles.css` is the GENESIS stylesheet plus a small
"CIS additions" section, and the shared components in `src/app/shared` mirror the GENESIS React components.

```powershell
npm ci
npm start          # http://localhost:4200, proxies /api to the Cis API on :5100
npm run lint       # oxlint
npm test           # unit tests (Vitest via the Angular CLI)
npm run build      # production build (dist/cis-ui; the API serves it from wwwroot in production)
npm run e2e        # browser smoke test against a running API + UI (see e2e/README.md)
```

| Folder | What |
|---|---|
| `src/app/core` | session / access rights, HTTP interceptor, guards, notifications, action log, lookups |
| `src/app/shared` | reusable UI (button, field, date input, code select, data grid, tabs, modal, page shell, icons) and pure helpers (`lib`) |
| `src/app/layout` | app layout (topbar, rail, sidebar) and `menu.ts`, the full CIS desktop menu with status per item |
| `src/app/pages` | login, home, module overview |
| `src/app/features` | one folder per module (maintenance, membership, invoice, claim, claim-bordx, lg-audit, mco-fees, payment, receipt, ac-export, reports) with its pages and routes |
| `e2e` | browser checks (see `e2e/README.md`) |
| `scripts/menu-mapping.mjs` | regenerates `../Cis/docs/menu-mapping.md` from `menu.ts` |
| `scripts/test-guide-coverage.cjs` | lists migrated menu items that no test guide in `../Cis/docs/testing` mentions |
| `scripts/theme-tokens.cjs` | colour themes (Purple / Navy blue): turns brand colours in `src/styles.css` into tokens and regenerates the THEME TOKENS block; `--check` audits that no brand colour is typed in (see `../Cis/docs/architecture.md`, "Colour themes") |

Colour themes: the palette button in the top bar switches Purple / Navy blue (`core/ui/theme.service.ts`, kept per browser in
`localStorage` `cis.theme`). After editing colours in `src/styles.css`, run `node scripts/theme-tokens.cjs` (and `--check`).

Node.js ≥ 22.22.3 is required. Dependencies are pinned to exact versions (`.npmrc` `save-exact=true`).
