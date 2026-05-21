# PBC Tour App

A web app for the Providence Baptist College tour team.

## Files

- `index.html` — Team-facing form app (Church Visit, Treasurer, Maverick Club, Social Photos, Tour Schedule)
- `dashboard.html` — Treasurer/Leader dashboard with live data, schedule management, CSV export
- `Code.gs` — Google Apps Script backend (paste into Apps Script editor in Google Sheets)

## Deployment

- **Form + Dashboard**: https://calebr13.github.io/pbc-tour-app/
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
- `seedLadiesSchedule()` — Populates Ladies tour schedule
- `seedMensSchedule()` — Populates blank Men's schedule structure
- `seedLadiesLodging()` — Populates Ladies lodging data
- `DANGER_deleteAndRebuildAllSheets()` — ⚠️ Destructive: requires confirmation

## Tour Groups

- **Men's names**: Putnam, Walker, Tan, McDonnell, Davadilla, Wilbar, Rardin
- **Ladies' names**: Bartels, Bessine, Bottrell, Chase, Miller, Palmer, Rardin

## Tour Dates
May 16 — August 2, 2026
