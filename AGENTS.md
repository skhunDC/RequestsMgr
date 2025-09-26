# Guidance for Agents

This project is a lightweight supplies request system built on Google Apps Script.

## Apps Script Delivery Expectations
- You are a senior Google Apps Script engineer. Update the project to ship a mobile-first web app with exactly two files: `Code.gs` (backend) and `index.html` (frontend via HTML Service). No extra files, no frameworks, no bundlers.
- Goals: fast on 4G, single-page UX, resilient to errors.
- Maintain clear client ↔ server contracts using JSON DTOs only.
- Respect Apps Script quotas and timeouts.

### Hard Constraints
- Only `doGet(e)` serves `index.html`. Client calls server via `google.script.run`.
- All server functions must finish in under 5 seconds; split heavy flows into steps.
- Data store may be Sheets, Drive, or UrlFetch; cache hot reads with `CacheService` (120–300s).
- Pass and return plain JSON only. Do not return HTML from the server.

### Frontend (`index.html`)
- Include `<meta name="viewport" content="width=device-width,initial-scale=1">`.
- Use system fonts, 44px touch targets, and card/list UI instead of wide tables.
- Keep state client-side (in-memory plus small `localStorage` cache).
- Show skeleton/disabled state during requests and use optimistic UI when safe.
- Sanitize user data by writing via `textContent` (avoid `innerHTML`).
- Provide a central `handleError(err)` that shows a toast/alert, offers retry, and logs via a server method.
- For pagination, request `{ nextToken }`, render the slice, and show a “Load more” control.

### Backend (`Code.gs`)
- `doGet(e)` must return `HtmlService.createTemplateFromFile('index').evaluate()`.
- Export server methods used by the client; each returns `{ ok: true, ... }` or `{ ok: false, code, message }`.
- Use `CacheService` on hot list reads and invalidate cache entries on writes.
- Ensure idempotency for mutations by requiring `clientRequestId` and deduplicating via cache (`rid:<id>`).
- Store configuration, flags, and secrets in `PropertiesService` (Script/User), never in HTML.

### Security
- Prefer Execute-as-user access unless kiosk/public flows require “execute as me”.
- Validate and whitelist resource IDs server-side and use least-privilege access.
- If exposing `doPost` endpoints, require a nonce and verify origin.

### Performance Budgets
- Initial HTML TTFB (warm) under 500 ms.
- First JSON payload under 150 KB.
- Interaction roundtrip target under 800 ms.
- Batch reads/writes (e.g., Sheets `batchGet`/`batchUpdate`). Return only needed fields.

### Logging & Observability
- Client attaches `cid` (correlation id) to every call.
- Server `try/catch` blocks must log `{ ts, fn, cid, message, stack, reqSummary }` via `Logger` and optionally append to a “Logs” sheet.

## Data Model
- **Orders**: `id | ts | requester | description | qty | status | approver`
- **Catalog**: `sku | description | category | archived`

Only the fields above are stored. Pricing and budget logic are intentionally omitted.

## Roles
- Any user may submit requests and view their own history.
- Approval and catalog features are accessible to all users.

## Conventions
- Keep code lean and mobile-first.
- Use `google.script.run` for all client ↔ server communication.
- Wrap sheet mutations with `withLock_` to avoid race conditions.

## Programmatic Checks
Run ESLint before committing:
```bash
npm test
```
