# Guidance for Agents

This project ships a production-ready Google Apps Script web app with a refined single-page UX and hardened access controls.

## Core Expectations

- **Only Apps Script runtime files** (`Code.gs`, `index.html`, `print.html`, `scripts.html`, `styles.html`) are deployed. Support docs and Node-based tooling live alongside them for local development.
- Keep the backend lean, fast (<5s per call), and always return JSON envelopes shaped as `{ ok, ... }`.
- All server entry points must call `handleServerCall_` (or an equivalent guard) so that authorization, correlation IDs, and logging are consistent.
- Frontend scripts should stay namespaced under `window.RequestsApp` / `window.RequestsAppHelpers` to avoid global leaks.
- UX must remain mobile-first, ADA-aware, and avoid `innerHTML` for user data.

## Data & Security

- Data persists in the auto-provisioned `Orders` sheet (header defined in `Code.gs`). Mutations must acquire a lock and update the cache + client request dedupe map.
- Authorization is enforced via script property allowlists (`REQUESTS_APP_DEVELOPERS`, `REQUESTS_APP_MANAGERS`, `REQUESTS_APP_REQUESTERS`). Log denied access attempts with correlation IDs.
- Never expose sensitive values in the HTML. Derive the active user from `Session.getActiveUser()`.

## Frontend Conventions

- Keep the app single-page with inline validation, toast notifications, and accessible status badges.
- Use the handcrafted utility-class CSS already provided; extend by following the same naming pattern (e.g., `.card`, `.field`).
- Client/server contracts: pass `cid` (correlation id) and `clientRequestId` for mutations. Expect responses shaped like `{ ok: true, ... }` or `{ ok: false, code, message }`.

## Tooling

- Local commands (`npm run lint`, `npm run format`, `npm test`) must stay green before shipping.
- Unit tests in `tests/` should focus on pure helpers—no GAS APIs.

## Documentation

- Keep `README.md` aligned with current auth flow, deployment steps, and future roadmap hints.
- Document any new script properties or environment expectations before handing work off.
