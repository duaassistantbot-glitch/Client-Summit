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
/home/openclaw/.openclaw/media/inbound/me-registration-report-27962-154818---2ba3c48a-dfc9-4c94-906a-c8f3eae1ae3c.xlsx
```

The committed dashboard intentionally excludes attendee emails, phone numbers, and source user IDs.

## Future HubSpot Source

Replace `loadWorkbookRows()` in `scripts/build-dashboard.js` with a HubSpot fetch layer that returns the same normalized row fields.

## Agenda

The Summit agenda tab and shareable PDF are generated from `scripts/agenda-data.js`.

The shareable PDF is written to:

```text
assets/Revio-Summit-2026-Agenda.pdf
```
