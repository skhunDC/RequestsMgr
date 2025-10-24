# Request Manager Web App

The Request Manager project delivers a Google Apps Script web application that centralizes supplies, IT, and maintenance requests. The solution pairs a mobile-friendly HTMLService UI with guarded server utilities that read and write to the `Orders` spreadsheet while enforcing authorization allowlists.

## Project layout

The Apps Script deployment only needs the files listed below. `Code.gs` exposes the server entry points, while the HTML files render the UI surfaces.

| File | Purpose |
| --- | --- |
| `Code.gs` | Apps Script backend that normalizes requests, enforces authorization, and wraps all entry points with telemetry and caching helpers. |
| `index.html` | Primary single-page workspace for submitting and managing requests. Includes hero, dashboards, request forms, and queue views. |
| `print.html` | Dedicated printable snapshot of the queues. Accepts query parameters (`type`, `scope`) to select which list to render. |
| `scripts.html` | Shared JavaScript payload injected into `index.html`. Establishes the `RequestsApp` namespace, fetches data from the server, and drives all interactive behavior. |
| `styles.html` | Global CSS design system used by both HTML views. |
| `appsscript.json` | Script manifest (timezone, OAuth scopes). |

Supporting documentation and tooling live alongside the runtime files:

- `/docs/README.md` (this file) and `/docs/AGENTS.md` capture design notes, authentication flow, and engineering guidelines.
- `/tests` contains Node-based unit tests that exercise pure helper functions.
- `/package.json` exposes linting, formatting, and test commands for local development.

## Frontend surfaces

### Main workspace (`index.html`)

The primary workspace is a single-page app structured around three request types (supplies, IT, maintenance). Key behaviors include:

- Tabbed hero to jump between request categories.
- Inline validation, toast notifications, and accessible status chips for each request item.
- Dashboard summary cards sourced from cached server metrics.
- Catalog lookups for supplies to suggest SKUs and vendor details.
- Status management controls exposed when the signed-in user appears on the `REQUESTS_APP_MANAGERS` allowlist.

The HTML template uses `<?!= include('styles'); ?>` and `<?!= include('scripts'); ?>` to pull in shared assets. All client logic lives inside the `window.RequestsApp` namespace defined in `scripts.html` to avoid global leaks.

### Print workspace (`print.html`)

`print.html` offers a printer-friendly version of the queues. It shares the global styles but ships a lightweight script dedicated to fetching the chosen queue and rendering a static table. Features include:

- Dropdowns for queue type (`supplies`, `it`, `maintenance`) and scope (`all`, `mine`).
- Iterative paging via `google.script.run.listRequests` to collect the full dataset.
- Inline note rendering with author/timestamp metadata.
- A `Print page` action that simply triggers the browser’s print dialog.

Serve this view by visiting the web app URL with `?view=print`. Optional `type` and `scope` parameters default to `supplies` and `all` respectively.

### Shared assets (`scripts.html`, `styles.html`)

- `scripts.html` defines `SESSION`, the `RequestsApp` controller, and the exported helper utilities under `window.RequestsAppHelpers`. The helpers remain accessible to Node-based tests.
- `styles.html` centralizes the design tokens, layout primitives, request card styles, and the print-specific layout classes.

## Backend services (`Code.gs`)

`Code.gs` manages all interactions with Google Sheets and encapsulates the application’s business logic:

- `doGet(e)` inspects `e.parameter.view` to select either the main (`index`) or print (`print`) template before embedding the authenticated session context.
- Every RPC exposed to the client funnels through `handleServerCall_` to ensure consistent logging, locking, and error envelopes shaped as `{ ok, ... }`.
- Request metadata is normalized via `REQUEST_TYPES` definitions, which specify sheet names, header layouts, and summary/detail builders.
- Caching layers (`CacheService`) and runtime memoization keep catalog, dashboard, and request listings responsive.
- Allowlist enforcement relies on Script Properties (`REQUESTS_APP_DEVELOPERS`, `REQUESTS_APP_MANAGERS`, `REQUESTS_APP_REQUESTERS`). The active user is derived from `Session.getActiveUser()` and logged with correlation IDs for traceability.

## Authentication & authorization flow

1. `doGet` calls `ensureSetup_()` to verify sheet schemas and bootstrap script properties.
2. `getStatusAuthContext_()` inspects the allowlists to determine whether the active user can update statuses. The resulting object (`session.statusAuth`) flows to both HTML templates for feature gating.
3. Each RPC invoked via `google.script.run` passes through `handleServerCall_`, which validates the caller, logs the correlation ID, and wraps the result in a standard `{ ok }` envelope.
4. Mutations acquire a lock, update Sheets, refresh caches, and broadcast toast messages back to the UI via the resolved promise.

All sensitive values (script properties, sheet IDs) remain server-side. Templates only receive the active email and authorization flags required for conditional rendering.

## Local development & testing

1. Install dependencies: `npm install` (root).
2. Lint and format: `npm run lint` and `npm run format` now include the split HTML partials.
3. Run tests: `npm test` delegates to `tests/package.json`, which executes Node’s built-in test runner (`node --test "*.test.js"`). Tests mock the browser environment to exercise helper functions without Apps Script APIs.

The repository intentionally avoids bundlers. When editing `scripts.html`, keep exports under `window.RequestsApp` / `window.RequestsAppHelpers`. For new shared assets, prefer using `include('fileName')` from Apps Script templates to stay consistent.

## Deployment tips

- Only deploy `Code.gs`, `index.html`, `print.html`, `scripts.html`, `styles.html`, and `appsscript.json` to Apps Script.
- Publish the web app as the service account that owns the Sheets so that caching and locking work reliably.
- Update allowlists through Script Properties (`Project Settings → Script properties`) before inviting new managers or requesters.
- After deploying, visit `/exec?view=print` to confirm the print snapshot loads with your account.

For deeper engineering notes and collaboration conventions, see [`docs/AGENTS.md`](./AGENTS.md).
