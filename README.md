# Dublin Cleaners Request Manager (pre-refactor experience)

A Google Apps Script web app that restores the original multi-track request workflow for Dublin Cleaners. Team members can submit
supplies, IT, and maintenance tickets while managers monitor queues, manage catalog inventory, and leave follow-up notes.

## Feature highlights

- **Three request types:** dedicated forms and dashboards for supplies purchases, IT incidents, and maintenance work orders.
- **Live catalog browsing:** filterable inventory metadata (category, supplier, estimated cost, usage counts) that can be attached to
  supplies requests or overridden with custom line items.
- **Status automation:** consistent status keys (`pending`, `approved`, `ordered`, `in_progress`, `completed`, `declined`) with
  audit trails and optional ETA updates for supplies.
- **Operational guardrails:** device-based rate limiting, duplicate request blocking via `clientRequestId`, centralized logging, and
  configurable notification allowlists for status changes.
- **Notes & collaboration:** IT and maintenance queues support threaded request notes so approvers can capture troubleshooting
  context without leaving the app.

## Apps Script architecture

`Code.gs` defines the original helper ecosystem that predates the UX/security refactor:

- **Sheets:**
  - `SuppliesRequests`, `ITRequests`, `MaintenanceRequests` for ticket storage.
  - `Catalog` for inventory metadata and usage analytics.
  - `Logs`, `StatusLog`, `RequestNotes` for auditing and collaboration history.
- **Script properties:**
  - `SUPPLIES_TRACKING_SHEET_ID` – backing spreadsheet id (auto-created if missing).
  - `SUPPLIES_TRACKING_SETUP_VERSION` – bootstrapping guard.
  - `SUPPLIES_TRACKING_STATUS_EMAILS` – comma-separated list for status notifications (defaults to developer allowlist).
- **Caching:** Apps Script cache stores catalog listings, request pages per type, processed client request ids, and status email
  lists.
- **Notifications:** Status transitions trigger batched Gmail notifications using the configured allowlist.

The HTML frontend (`index.html`) mirrors the historical single-page experience: a dashboard landing section, tabbed views for each
request category, catalog search with quick-add cards, and inline status management controls. The new helper shim at the bottom of
the file exports `window.RequestsAppHelpers` to keep the current Node-based unit tests satisfied.

## Setup

1. Install local tooling dependencies:
   ```bash
   npm install
   ```
2. Push `Code.gs`, `index.html`, and `appsscript.json` to your Apps Script project using [clasp](https://github.com/google/clasp)
   or the Apps Script editor.
3. Ensure the backing spreadsheet contains the sheets listed above. The script auto-creates headers when missing and seeds catalog
   defaults the first time it runs.
4. Populate script properties as needed:
   - `SUPPLIES_TRACKING_STATUS_EMAILS`
   - (Optional) override the default spreadsheet id via `SUPPLIES_TRACKING_SHEET_ID`.
5. Deploy as a **web app** that executes as the accessing user with limited sharing (only requesters/approvers).

## Local development & testing

- `npm test` runs the lightweight helper validation suite used by CI.
- `npm run lint` (if configured) keeps the legacy code style consistent.

Because this branch intentionally mirrors the pre-refactor experience, newer guard helpers (`handleServerCall_`, role namespaces,
 etc.) are not present. Review the follow-up refactor branch when you are ready to reapply the hardened UX and authorization layer.
