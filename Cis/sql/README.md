# Database scripts

The API calls stored procedures in a `cis` schema in **CISDB** instead of sending SQL text. This folder holds them.
They are new objects only: no existing table, view or procedure is changed, so the VB6 CIS and HIS desktops and GENESIS keep
working against the same databases.

| File | What it is |
| --- | --- |
| `main/000_setup.sql` | Checks the server and database, creates the `cis` schema and the synonyms to the other legacy databases |
| `main/010_auth.sql` | Login / logout / change password, error log, action log (frmLogin, MDIMain, Medix.bas) |
| `main/020_lookups.sql` | Shared drop-down lists (payor, health type, insured type, group company, clinic) |
| `cis-main.sql` | **Generated** from `main/*.sql` in name order. Do not edit by hand |

## How to run

```powershell
sqlcmd -S localhost -E -d CISDB -i Cis\sql\cis-main.sql
```

Safe to run again: it only creates or replaces `cis.*` procedures and synonyms.

## Synonyms

The procedures run on one connection to CISDB. Tables of the other databases are reached through `cis.<prefix>_<table>` synonyms, so a
save that touches several databases is one transaction:

| Prefix | Database (variable in `000_setup.sql`) |
| --- | --- |
| `Maint_` | HISMaintenance (`@MaintenanceDb`) |
| `Hdb_` | HISDB (`@HisDb`) |
| `Ws_` | CISWorksheet (`@CisWorksheetDb`) |
| `Pay_` | CISPayment (`@CisPaymentDb`) |
| `HWs_` | HISWorkSheet (`@HisWorksheetDb`) |
| `HPay_` | HISPayment (`@HisPaymentDb`) |
| `Arch_` | HISDMArchive (`@ArchiveDb`) |

## Rules for new scripts

- One file per screen: `main/NNN_<screen>.sql`, procedures named `cis.<Screen>_<Action>`.
- Parameters only. Legacy procedures that append a raw `@strAppend` string into dynamic SQL are never called; they are reproduced as
  parameterised procedures.
- CISDB runs at compatibility level 120, CISWorksheet at 100: avoid `STRING_AGG`, `TRIM`, `CONCAT_WS`, `STRING_SPLIT` and other
  post-2014 functions in code that runs in those databases.
- After editing `main/*.sql`, rebuild the combined script: `$env:CIS_WRITE_SQL='1'; dotnet test --filter Combined_script_is_up_to_date`.
