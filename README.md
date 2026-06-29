# PBC Tour App

A web app for the Providence Baptist College tour team.

## Files

- `index.html` — Team-facing form app (Church Visit, Treasurer, Maverick Club, Social Photos, Tour Schedule)
- `dashboard.html` — Treasurer/Leader dashboard with live data, schedule management, CSV export
- `Code.gs` — Google Apps Script backend (paste into Apps Script editor in Google Sheets)

## Deployment

- **Form**: https://calebr13.github.io/pbc-tour-app/
- **Dashboard**: https://calebr13.github.io/pbc-tour-app/dashboard.html
- **GitHub Repo**: https://github.com/calebr13/pbc-tour-app

## Google Services

- **Google Sheet**: Travel Team Data
- **Apps Script URL**: https://script.google.com/macros/s/AKfycbwWf-ONXMIPyuBm4IP0YQJy7Ea4w0sUIo4kvYaaei407FxBMPlUgPrwLtirI2aoL4fk/exec
- **Tour Receipts folder**: https://drive.google.com/drive/folders/1tCBpsUqS4FUyo00X8Nus6gg3HoAcjBvq
- **Tour Maverick Club folder**: https://drive.google.com/drive/folders/1Ip2NYKB9NBlE1rxPcHenk6ZCGGEA59t6
- **Tour Socials folder**: https://drive.google.com/drive/folders/1QLUh0LEwvmzciJmcpXGp2kFSWW7qmSxS

## Google Sheet Tabs

| Tab | Description |
|-----|-------------|
| Church Evaluations | Church visit evaluations |
| Expenses | Trip expenses with receipts |
| Income | Income/offerings by type |
| Interest Cards | Student Maverick Club cards |
| Social Photos | Social media photo log |
| Tour Schedule | Daily tour schedule (both tours) |
| Tour Lodging | Nightly lodging per tour group |
| Submission Log | All form submissions |

## Apps Script Functions

- `setupSheets()` — Safe: only creates missing tabs, never deletes data
- `addChurchLocationColumns()` — Safe: adds structured `Church Name`, `Church City`, `Church State` columns where missing
- `migrateChurchLocationFields()` — Safe: backfills structured church location columns from existing combined church strings when possible
- `seedLadiesSchedule()` — Populates Ladies tour schedule
- `seedMensSchedule()` — Populates blank Men's schedule structure
- `seedLadiesLodging()` — Populates Ladies lodging data
- `DANGER_deleteAndRebuildAllSheets()` — ⚠️ Destructive: requires confirmation

## Church Location Update

The app now stores church information in both formats:

- Legacy combined field: `Church Name / City / State`
- Structured fields: `Church Name`, `Church City`, `Church State`

Recommended rollout:

1. Deploy the updated `Code.gs`
2. Run `addChurchLocationColumns()` once
3. Run `migrateChurchLocationFields()` once
4. Manually review rows where city/state could not be inferred from older freeform entries

This migration is non-destructive: it preserves the original combined church field and only fills the new structured columns.

## Schedule Evaluation Flag

The `Tour Schedule` sheet now supports a `Needs Evaluation` column.

- `Yes` — show this stop in the Church Evaluation picker if no evaluation has been completed yet
- blank or `No` — do not show it in the picker

The app will still hide a stop after a matching evaluation has been submitted.

## Tour Groups

- **Men's names**: Putnam, Walker, Tan, McDonnell, Davadilla, Wilbar, Rardin
- **Ladies' names**: Bartels, Bessine, Bottrell, Chase, Miller, Palmer, Rardin

## Tour Dates
May 16 — August 2, 2026
