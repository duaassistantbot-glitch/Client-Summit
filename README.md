# Rev.io Client Summit Dashboard

Static dashboard for the 2026 Rev.io Client Summit.

## Build

```bash
npm install
npm run build
npm run agenda:pdf
```

The current build reads the vFairs/registration Excel export from:

```text
assets/source/registration-report-deduped-2026-08-12.xlsx
```

The committed dashboard intentionally excludes attendee emails, phone numbers, and source user IDs.

## Updating registration data

Do **not** replace the full workbook with each new vFairs export. Referral attribution can change or disappear in later exports, so updates must append only net-new registrants by email and leave all existing rows untouched.

```bash
cd revio-client-summit-dashboard
INCOMING_XLSX=/path/to/latest-vfairs-export.xlsx node scripts/append-new-registrants-only.js
node scripts/build-dashboard.js
```

The append script writes a timestamped backup before changing the current workbook.

## Future HubSpot Source

Replace `loadWorkbookRows()` in `scripts/build-dashboard.js` with a HubSpot fetch layer that returns the same normalized row fields.

## Agenda

The Summit agenda tab and shareable PDF are generated from `scripts/agenda-data.js`.

The shareable PDF is written to:

```text
assets/Revio-Summit-2026-Agenda.pdf
```
