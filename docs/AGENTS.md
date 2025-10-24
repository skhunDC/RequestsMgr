# Engineering notes

This document supplements the root `AGENTS.md` with context on the refreshed file layout and day-to-day development workflows.

## Templates and includes

- `doGet(e)` now selects `index.html` or `print.html` based on `e.parameter.view`. Both templates call `<?!= include('styles'); ?>` and, when needed, `<?!= include('scripts'); ?>` to keep shared assets in one place.
- Use `include('fileName')` for any future partials instead of duplicating markup. The helper lives in `Code.gs`.
- `scripts.html` must keep the `window.RequestsApp` namespace intact. Add new client modules under that object or `window.RequestsAppHelpers` so tests can continue to load them.

## Frontend expectations

- Preserve the mobile-first layout defined in `styles.html`. New components should follow the existing utility classes (`.card`, `.meta`, `.request-item`, etc.).
- Avoid mutating the DOM with raw HTML strings when working with user data—stick to `textContent`/`setAttribute` for sanitization.
- The print view (`print.html`) intentionally ships a lighter script. If you need main-app utilities there, consider extracting them into a shared helper rather than importing the entire SPA bundle.

## Backend guardrails

- All RPCs should flow through `handleServerCall_` so that logging, locking, and error envelopes remain consistent.
- Continue returning `{ ok, ... }` objects to the client. Use the existing request definitions in `REQUEST_TYPES` for normalization.
- Authorization decisions rely on the script property allowlists (`REQUESTS_APP_DEVELOPERS`, `REQUESTS_APP_MANAGERS`, `REQUESTS_APP_REQUESTERS`). Never expose those values in templates.

## Tooling & tests

- Run `npm run lint`, `npm run format`, and `npm test` before committing. `npm test` now delegates to `tests/package.json`, which uses the Node `test` runner.
- Tests load helper functions directly from `scripts.html`. Keep helper exports co-located with the UI code so they stay in sync.
- If you add new pure functions, mirror them with small Node tests under `/tests` to ensure they remain Apps Script agnostic.

## Deployment checklist

- Only the runtime files (`Code.gs`, `index.html`, `print.html`, `scripts.html`, `styles.html`, `appsscript.json`) should be pushed to Apps Script.
- After deploying, validate both the main workspace and the print view (`?view=print&type=supplies`) using an account that reflects the target authorization tier.
