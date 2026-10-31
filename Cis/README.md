# Cis: Web API of the CIS migration

ASP.NET Core 10 Web API that, with the Angular app in `../CisUI`, replaces the VB6 desktop **MedixCIS** (`../CIS VB`).

| Folder | What |
|---|---|
| `src/Cis.Domain` | entities, business rules, domain errors (no dependencies) |
| `src/Cis.Application` | use cases (services), ports (repository / infrastructure interfaces), request / response models, validators |
| `src/Cis.Infrastructure` | adapters: Dapper repositories on `cis.*` procedures, SQL sessions, Excel / PDF, e-mail, file shares |
| `src/Cis.Api` | HTTP host and composition root: controllers, JWT, access filter, error handling, Swagger |
| `tests/Cis.Api.Tests` | unit, service, architecture, integration and database tests |
| `sql` | additive SQL scripts for CISDB (`cis-main.sql` is generated) |
| `docs` | [architecture](docs/architecture.md), [onboarding](docs/onboarding.md), [menu mapping](docs/menu-mapping.md) (every desktop menu item and its status), [screen checklists](docs/screens) |

Quick start: [docs/onboarding.md](docs/onboarding.md).
