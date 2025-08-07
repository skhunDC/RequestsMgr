# SuppliesTracking

Centralized Supplies Ordering & Tracking System for Dublin Cleaners' Leadership Team.

## Setup
1. Clone repository and install dependencies:
   ```bash
   npm install
   ```
2. Enable Google Apps Script API and configure [`clasp`](https://github.com/google/clasp) for deployment.
3. Create the Google Sheets with tabs `Orders`, `Catalog`, and `Audit` following the column layouts in [AGENTS.md](AGENTS.md).
4. Update `CHAT_WEBHOOK` in `Code.gs` with the Google Chat incoming webhook URL.

## Local UI Testing
The UI can be served locally with Vite for rapid development. HTML/JS is later copied into Apps Script.
```bash
npm run dev
```
This serves `index.html` and enables ES modules. No build step is required for Apps Script deployment.

## Deployment
1. Run `clasp push` to upload `Code.gs` and `index.html` to Apps Script.
2. Visit the Apps Script project and set up a daily time trigger for `sendDailyDigest`.
3. Share the underlying Google Sheet with all Leadership Team members.

## Mobile‑First Guidelines
- Phone (≤ 480px) is the primary breakpoint. A sticky bottom nav is used.
- Tablet/desktop show a left sidebar. Components scale using flexbox and utility classes.
- Maintain WCAG AA contrast; interactive elements follow Google Material patterns.

## Auth Flow
- The app relies on `Session.getActiveUser().getEmail()` to identify the signed-in Google account.
- Only the following Leadership Team email addresses can access the main UI by default:
  `skhun@dublincleaners.com`, `ss.sku@protonmail.com`, `brianmbutler77@gmail.com`, `brianbutler@dublincleaners.com`,
  `rbrown5940@gmail.com`, `rbrown@dublincleaners.com`, `davepdublincleaners@gmail.com`, `lisamabr@yahoo.com`,
  `dddale40@gmail.com`, `nismosil85@gmail.com`, `mlackey@dublincleaners.com`, `china99@mail.com`.
  Update `LT_EMAILS` in `Code.gs` to modify this list.
- The Developer Console is restricted to predefined developer emails or those added via the console.

## Testing
Run ESLint before committing:
```bash
npm run lint
```
