# Guidance for Agents

This project implements a centralized supplies ordering and tracking system using Google Apps Script and HTMLService.

## Data Model
- **Orders**: `id | ts | requester | item | qty | est_cost | status | approver | decision_ts | override? | justification | cost_center`
- **Catalog**: `sku | desc | vendor | price | override_required | threshold | gl_code | cost_center`
- **Audit**: append-only JSON diff per state change.

## Roles
- Leadership Team users may submit requests and view their own history.
- Approvers can bulk approve/deny requests.
- Developers manage catalog, budgets, and role assignments via the Dev Console.

## Approval Flow
1. User submits request (optional override flag + justification when required).
2. Approver decides PENDING requests → `APPROVED`, `DENIED`, or `ON-HOLD`.
3. Budget guardrail warns at 80 %, blocks at 100 % unless super-admin override.
4. Audit sheet records each change.

## Override Logic
Items with `override_required=true` demand a Yes/No override toggle and 40‑character justification before submission.

Please keep documentation helpful but not restrictive; avoid language that blocks future features.
