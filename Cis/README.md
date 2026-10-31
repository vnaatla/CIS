# Cis: Web API of the CIS migration

ASP.NET Core 10 Web API that, with the Angular app in `../CisUI`, replaces the VB6 desktop **MedixCIS** (`../CIS VB`).

| Folder | What |
|---|---|
| `src/Cis.Api` | the API (layered: controllers → validators → services → repositories → `cis.*` procedures) |
| `tests/Cis.Api.Tests` | unit, service, architecture, integration and database tests |
| `sql` | additive SQL scripts for CISDB (`cis-main.sql` is generated) |
| `docs` | [architecture](docs/architecture.md), [onboarding](docs/onboarding.md), [menu mapping](docs/menu-mapping.md) (every desktop menu item and its status), [screen checklists](docs/screens) |

Quick start: [docs/onboarding.md](docs/onboarding.md).
