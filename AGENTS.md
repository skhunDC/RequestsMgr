# Guidance for Agents

This project is a lightweight supplies request system built on Google Apps Script.

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
