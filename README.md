# Dublin Cleaners Request Manager

A Google Apps Script web app that collects Supplies, IT, and Maintenance requests into a single, mobile-friendly workspace. The refreshed hero experience highlights the official Dublin Cleaners logo plus a reassurance banner that every submission is auto-saved to Google Sheets—creating the workbook on first run if needed.

## Feature highlights
- **Three request lanes:** Dedicated forms and dashboards for supplies purchases, IT incidents, and maintenance work orders with inline validation and tab-aware gradients.
- **Auto-provisioned storage:** `handleServerCall_` enforces setup before every server call, generating the `RequestManager` spreadsheet (and its headers) when missing. Requests, catalogs, logs, and notes live in separate sheets and stay cache-friendly.
- **Authorization guardrails:** Status updates and notes respect the approver allowlist (`SUPPLIES_TRACKING_STATUS_EMAILS`), while active-user emails are normalized to avoid leaks. Correlation IDs accompany every response for reliable logging.
- **Collaboration extras:** Catalog usage metrics, status automation, and request notes keep managers and requesters aligned without leaving the page.

## Apps Script layout
- **Runtime files:** `Code.gs`, `index.html`, `styles.html`, and `scripts.html` are deployed. The template uses `include()` to inline supplemental styles/scripts while keeping helper tests intact.
- **Server helpers:** Entry points wrap through `handleServerCall_`, which ensures setup, applies locks where needed, and emits JSON envelopes shaped as `{ ok, cid, ... }` on success or `{ ok: false, code, message }` on failure.
- **Storage model:** Spreadsheet tabs: `SuppliesRequests`, `ITRequests`, `MaintenanceRequests`, `Catalog`, `Logs`, `StatusLog`, and `RequestNotes`. Headers auto-heal if they drift.

## Deployment & setup
1. Install dependencies locally: `npm install`.
2. Push `Code.gs`, `index.html`, `styles.html`, `scripts.html`, and `appsscript.json` to your Apps Script project (via clasp or the editor).
3. Deploy as a **web app** that executes as the accessing user. The script will create the backing sheet if `SUPPLIES_TRACKING_SHEET_ID` is unset.
4. Optional script properties:
   - `SUPPLIES_TRACKING_SHEET_ID` – override the auto-created spreadsheet id.
   - `SUPPLIES_TRACKING_STATUS_EMAILS` – comma/JSON list of approvers for status edits.

## Testing & tooling
- Run helper tests with `npm test` from the repo root or directly inside `tests/` (see `tests/package.json`).
- Lint and format helpers live in `npm run lint` and `npm run format` (HTML-aware ESLint + Prettier).

Additional implementation notes and flow diagrams live in `docs/README.md`.
