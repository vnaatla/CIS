# Database scripts

The API calls stored procedures in a `cis` schema in **CISDB** instead of sending SQL text. This folder holds them.
They are new objects only: no existing table, view or procedure is changed, so the VB6 CIS and HIS desktops and GENESIS keep
working against the same databases.

| File | What it is |
| --- | --- |
| `main/000_setup.sql` | Checks the server and database, creates the `cis` schema and the synonyms to the other legacy databases |
| `main/010_auth.sql` | Login / logout / change password, error log, action log (frmLogin, MDIMain, Medix.bas) |
| `main/020_lookups.sql` | Shared drop-down lists (payor, health type, insured type, group company, clinic) |
| `main/100`–`220_*.sql` | Maintenance: list reports, LG templates, benefits (clinic, member, dental), entity, plan code / premium / limits / MCO |
| `main/300`–`365_*.sql` | Membership: member search, top-up, case transfer, registration, general exclusions, MBM adjustment and enquiry |
| `main/400`–`530_*.sql` | Invoice: batch transfer, invoice update, OP batch, pop-ups, pass to claim, worksheet, clinic / member invoice, enquiry, adjustments |
| `main/600`–`630_*.sql` | LG & Audit: audit review and tracking, LG generation |
| `main/700`–`740_*.sql` | MCO Fees: fee and refund listings, MCO invoice, GST, GST zeroised |
| `main/800`–`866_*.sql` | Claim Bordx: insert worksheet / invoice, remove, bordx date, delivery, monitoring, print, BSN summary |
| `main/900`–`940_*.sql` | Payment: to clinic, clinic invoice, to member, tracking, pay clinic post, FFS |
| `main/950`–`990_*.sql` | Receipt: from member, payor, FFS, upload, MNRB |
| `main/1000`–`1050_*.sql` | A/C Export: SUN postings (invoice, member, payment, bordx, receipt) |
| `main/1100`–`1154_*.sql` | Tracking and reports: clinic tracking, Claim ASCII, invoice tracking, e-mail notifications, MBM / Claim tracking engines (1140–1154 are **generated**, see below) |
| `cis-main.sql` | **Generated** from `main/*.sql` in name order. Do not edit by hand |
| `tools/` | Generators of the tracking engines and the desktop procedures they are generated from (`*.source.sql`) |

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
- Values in generated or copied SQL use single quotes: double-quoted values only work under QUOTED_IDENTIFIER OFF (scripts start with
  `SET QUOTED_IDENTIFIER ON`).

## Generated tracking engines (`tools/`)

The MBM and Claim Tracking report engines are generated from the desktop procedures, with each case's query kept and the typed criteria as
parameters. Edit the generators, not the output:

| Command (node, from the folder holding the form extracts) | Output |
| --- | --- |
| `gen-trk.js mbm` (CISMBMTracking, filters `build-mbmtrk-filters.js`) | `main/1140_mbm_tracking.sql` (`cis.MbmTrk_Run`) |
| `gen-trk.js mbmsimple` (CISMBMTracking_Simple) | `main/1142_mbm_tracking_simple.sql` (`cis.MbmTrkSimple_Run`) |
| `gen-trk.js claim` (CISClaimTracking, `build-clmtrk-filters.js`) | `main/1150_claim_tracking.sql` (`cis.ClmTrk_Run`) |
| `gen-trk.js claimsimple` (CISClaimTracking_Simple) | `main/1152_claim_tracking_simple.sql` (`cis.ClmTrkSimple_Run`) |
| `gen-trk.js claim2` (CISClaimTracking2, `build-ct2-filters.js`) | `main/1154_claim_tracking2.sql` (`cis.ClmTrk2_Run`) |
| `gen-mbm-all.sh <form> <catalogue>`, `analyse-clmtrk.js` (TRK_FORM / TRK_CAT), `analyse-ct2.js` | the report catalogues in `src/Cis.Application/Reports/*.catalogue.json` |

`subdiff.js <form1> <form2>` lists the routines that differ between two form extracts (used to find what a variant form changes).
