# Colour theme (top bar) - manual test guide

New in the web app (no desktop equivalent): the palette button in the top bar, next to Print, switches **Purple** (default) and
**Navy blue**. The choice is kept in the browser (`localStorage` key `cis.theme`); nothing is written to the database.
Sign in as DEVADMIN. Automated: `$env:E2E_PASSWORD='Gen12345'; $env:E2E_STEPS='theme'; node e2e/only.mjs`.

| # | Steps | Expected |
|---|---|---|
| T1 | Any screen, e.g. Membership > MBM Adjustment → palette button | Menu "Theme" with Purple (check mark) and Navy blue, each with a colour swatch |
| T2 | Choose **Navy blue** | Every colour changes at once, without a reload: top bar line and logo, menu icon tiles and active item, page header icon, section markers, buttons (deep navy #0B2038 / accent #1F4E8C), tabs, dropdown arrows, grid header, pop-ups, home banner |
| T3 | Open other modules (Invoice > Clinic, Claim Bordx > Update Claims Data, Payment > To Clinic, Maintenance > Entity, MBM Tracking) | No purple left anywhere; red / green / amber messages and statuses keep their colours |
| T4 | Reload the page (F5) | Still Navy blue, with no flash of purple while loading |
| T5 | Log out | The login page is Navy blue too; sign in again → still Navy blue |
| T6 | Keyboard: Tab to the palette button → Arrow Down → Arrow Up / Down → Enter; open again → Esc | Arrow Down opens the menu on the current theme, arrows move, Enter picks and closes, Esc closes; focus returns to the button |
| T7 | Open the menu → click elsewhere on the page | The menu closes without changing the theme |
| T8 | Choose **Purple** | Back to the original colours; `localStorage` no longer has `cis.theme` |
| T9 | Another browser or a private window | Purple (the choice is per browser) |

Clean-up: choose Purple again (or remove `cis.theme` from the browser's local storage).
