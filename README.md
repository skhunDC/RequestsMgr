# SuppliesTracking

A minimal supplies request and approval workflow built with Google Apps Script and HTMLService.

## Features
- Mobile‑first request form with stock search, category filter, custom line items, and cart submission.
- All Team Requests view showing the last 90 days of orders from everyone.
- Admin views for pending approvals and catalog management (add / archive items).

## Setup
1. Install dependencies:
   ```bash
   npm install
   ```
2. Use [clasp](https://github.com/google/clasp) to push `Code.gs` and `index.html` to an Apps Script project tied to a Google Sheet.
3. The script creates `Orders` and `Catalog` sheets if missing and seeds the catalog with default stock items.

## Development
Serve the HTML locally with Vite for quick iteration:
```bash
npm run dev
```

## Testing
Run ESLint before committing:
```bash
npm test
```
