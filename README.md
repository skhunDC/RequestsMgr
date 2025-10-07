# Request Manager (Google Apps Script)

A premium-feeling request intake and tracking tool for Dublin Cleaners, delivered with Google Apps Script + HTML Service. The app keeps operations simple for requesters while enforcing role-based access and audit logging for managers and developers.

## Architecture

- **Apps Script backend** (`Code.gs`) manages authorization, validation, Google Sheet storage, caching, and structured JSON responses.
- **HTMLService frontend** (`index.html`) renders a responsive single-page experience with inline validation, toasts, keyboard-friendly controls, and status management.
- **Google Sheet data** is auto-provisioned as `Orders` within the bound spreadsheet (script property `REQUESTS_APP_SHEET_ID`). Logs are written to the `Logs` sheet in the same spreadsheet for audit trails.

## Access control & roles

The app authenticates with the active Google Workspace session (`Session.getActiveUser`). Every server entry point enforces an allowlist.

| Role        | Purpose             | Script property           | Default                                                            | Capabilities                                              |
| ----------- | ------------------- | ------------------------- | ------------------------------------------------------------------ | --------------------------------------------------------- |
| `developer` | Trusted maintainers | `REQUESTS_APP_DEVELOPERS` | Seeded with Dublin Cleaners developer emails                       | Full access, can edit statuses, and access admin features |
| `manager`   | Operations managers | `REQUESTS_APP_MANAGERS`   | Defaults to developer allowlist if unset                           | Create requests and update statuses                       |
| `requester` | Standard teammates  | `REQUESTS_APP_REQUESTERS` | Defaults to developer allowlist if unset (configure before launch) | Create requests and view their submissions                |

To configure, open **Project Settings → Script properties** and set comma-separated email lists for the keys above. Users without a listed email see an access denied page with a correlation ID.

## Deploying the web app

1. Install dependencies for local linting/tests:
   ```bash
   npm install
   ```
2. Use [clasp](https://github.com/google/clasp) or the Apps Script editor to push `Code.gs` and `index.html` to your project.
3. When the script first runs it will create a Spreadsheet (if `REQUESTS_APP_SHEET_ID` is unset) and seed the `Orders` header. Share the spreadsheet with collaborating managers if needed.
4. Deploy as a **web app**:
   - **Execute as:** User accessing the web app.
   - **Who has access:** Only specific people or groups. Add the same allowlisted users for predictable behavior.
5. Test in an incognito Chrome window to confirm OAuth prompts show the signed-in email and the UI reflects the correct role.

## Local tooling

| Command          | Description                                                     |
| ---------------- | --------------------------------------------------------------- |
| `npm run lint`   | ESLint on `Code.gs`, `index.html`, and tests for early feedback |
| `npm run format` | Prettier write-mode for project files                           |
| `npm test`       | Node-based unit tests for pure helper utilities                 |

All scripts run locally; no bundlers or transpilers are required.

## Running unit tests

Pure helper functions used on the client live in `tests/helpers.test.js`. They validate payload sanitization and ID generation logic without calling Apps Script APIs:

```bash
npm test
```

## Operational notes

- All mutations enforce a `clientRequestId` for idempotency. Duplicate submissions surface the existing record.
- Server-side logs capture denied access attempts with timestamps, correlation IDs, and required roles for review.
- Cache entries for request lists expire automatically and are invalidated on writes to keep the UI in sync.

## Future extensibility ideas

1. **Approval routing:** add approver assignments and multi-step statuses (e.g., “Awaiting approval”).
2. **Notifications:** integrate Gmail or Chat webhooks triggered from Apps Script when statuses change.
3. **Catalog support:** attach optional inventory SKUs via an additional `Catalog` sheet and expose filtered selectors client-side.
4. **Metrics dashboard:** surface charts using cached aggregates for managers without leaving the app.
