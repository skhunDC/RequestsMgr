# Request Manager

A Google Apps Script web application that centralizes supplies, IT, and maintenance requests for Dublin Cleaners. The refreshed layout keeps the production UX while reorganizing the runtime into reusable partials and adds a printable snapshot of each queue.

- **Modern single-page workspace** — `index.html` renders responsive request forms, dashboards, catalog lookups, and status management controls.
- **Printable summaries** — `print.html` loads a lightweight view driven by `google.script.run.listRequests` so supervisors can export queue snapshots with a single click.
- **Shared assets** — `styles.html` and `scripts.html` are injected with `include('fileName')`, keeping the design system and SPA logic in one place.
- **Hardened backend** — `Code.gs` normalizes input, enforces Script Property allowlists, wraps RPCs with `handleServerCall_`, and caches frequently accessed data.

For detailed architecture notes, see [`docs/README.md`](docs/README.md). Engineering guardrails live in [`docs/AGENTS.md`](docs/AGENTS.md).

## Deploying to Apps Script

1. Install dependencies for local tooling:
   ```bash
   npm install
   ```
2. Push the runtime files (`Code.gs`, `index.html`, `print.html`, `scripts.html`, `styles.html`, `appsscript.json`) to your Apps Script project.
3. Ensure the backing spreadsheet contains the sheets referenced in `REQUEST_TYPES` plus `Catalog`, `Logs`, `StatusLog`, and `RequestNotes`. `ensureSetup_()` will seed headers when missing.
4. Configure script properties as needed:
   - `SUPPLIES_TRACKING_SHEET_ID`
   - `SUPPLIES_TRACKING_SETUP_VERSION`
   - `SUPPLIES_TRACKING_STATUS_EMAILS`
   - Authorization allowlists: `REQUESTS_APP_DEVELOPERS`, `REQUESTS_APP_MANAGERS`, `REQUESTS_APP_REQUESTERS`
5. Deploy as a web app executing as the script owner. Visit the published URL normally for the main UI, or append `?view=print&type=supplies` for the printable snapshot.

## Local development

```bash
npm run lint    # ESLint across Apps Script and HTML partials
npm run format  # Prettier for scripts, styles, and docs
npm test        # Runs Node-based unit tests from tests/package.json
```

Tests rely on the helper exports defined in `scripts.html`, so keep shared utilities under `window.RequestsAppHelpers`. When adding pure functions, mirror them with small Node tests under `/tests`.

## Support & next steps

- Extend `print.html` if supervisors need additional columns or filtering options.
- Update the allowlists in Script Properties before onboarding new managers or requesters.
- Review the docs in `/docs` for deeper context on authentication, caching, and UI patterns before making significant changes.
