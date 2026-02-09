# Project Summary

## Overview
Chrome extension that collects product parameters on Gomer pages, finds matching offers in price feeds, and appends rows to Google Sheets. It supports multiple Gomer contexts (on‑moderation, active, changes, item‑details, binding‑attribute‑page) and keeps task/category context when navigating to binding pages.

## Main Features
- **Google Sheets export** with formatted hyperlinks (HYPERLINK formula + blue underline styling).
- **Offer search in price feeds** by `offerId + paramName + paramValue` (strict match).
- **Offer ID collection** from:
  - `on‑moderation`
  - `active`
  - `changes`
- **Binding page support**:
  - Opens binding page in a new tab.
  - Preserves task ID and category ID between pages.
  - Uses correct attribute selector for search vs. file output.
- **Value handling**:
  - Extracts `ru:` value from multilingual text.
  - Uses selected text (if any) for file value/link.
  - Uses full value for price search.
- **Toast flow**:
  - Single toast updated through stages (search → not found → add to file).
  - Red toast for not found/price unavailable.
- **Price availability handling**:
  - If price link is `javascript:void(0)` or `title="Api virtual source"` → shows “Прайс недоступен” and still writes row (without product link).

## Current Structure (params/)
- `manifest.json` — MV3 manifest, permissions, background service worker.
- `src/content/content.js` — UI injection, DOM parsing, binding-page flow, messaging to background.
- `src/background/background.js` — Sheets API, price feed parsing, matching logic.
- `src/popup/popup.html` / `src/popup/popup.js` — UI to save Google Sheet URL.
- `README.md` — Quick setup and usage.
- `.gitignore` — excludes `.pem` and `.DS_Store`.

## Core Flow
1. User saves Google Sheet URL in popup.
2. On Gomer pages, user clicks the Excel button.
3. Extension:
   - Reads category, attribute, value.
   - Collects offer IDs from proper list page (on‑moderation / active / changes).
   - Searches price feed for matching offer (by offerId + param + value).
   - Appends row to Google Sheets.

## Key Selectors
### Binding page (search vs. file)
- **Search param name**: `#pv_id_7 > span > div` → strip id `(####)`
- **File attribute (full text)**: `#pv_id_8 > span > div`

### Item-details category
- On‑moderation category: `td:nth-child(6) > a:nth-child(1)` (Rozetka category link)
- Price category (seller pin): `td:nth-child(5) > span` (title with id)

## Storage Keys
- `sheetUrl` — Google Sheets URL
- `bindingPageTaskId`
- `bindingPageCategoryId`
- `bindingPageSourceType` (on‑moderation / active / changes)

## Notes
- Price search is **strict** by `<param name="X">Y</param>`.
- Offer ID collection uses **bpm_number** filter (task ID) as before.
- If price not found: row still added but without product link.
- `extension.pem` is kept locally for stable extension ID; excluded from Git.
