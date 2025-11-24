# Dublin Cleaners Request Manager docs

## Experience overview
The Request Manager surfaces a single-page workspace for Supplies, IT, and Maintenance teams. The hero section now spotlights the Dublin Cleaners logo and a reassurance banner that every submission is auto-saved to Google Sheets. If the backing spreadsheet cannot be found, the Apps Script automatically provisions a new workbook named `RequestManager` and seeds required headers, keeping catalog, requests, logs, and notes aligned.

## Authentication & authorization
- Users are identified by `Session.getActiveUser()` and emails are normalized before use.
- Status changes and note updates still rely on the approver allowlist stored in the `SUPPLIES_TRACKING_STATUS_EMAILS` script property (with sensible defaults and owner inclusion).
- All server entry points flow through `handleServerCall_`, which ensures setup, wraps correlation IDs, and centralizes error logging.

## Frontend design updates
- Added a brand ribbon with the official Dublin Cleaners stacked logo plus sheet health badges.
- Included timezone-aware messaging via `RequestsApp.refreshStorageBanner` so users see when autosave data will be timestamped.
- Supplemental styles live in `styles.html` to keep the main template lean; scripts in `scripts.html` stay under the `window.RequestsApp` namespace.

## Testing notes
Local helper tests run with `npm test` from the repo root. A dedicated `tests/package.json` mirrors the commands for developers who prefer running the suite directly from the `tests/` directory.
