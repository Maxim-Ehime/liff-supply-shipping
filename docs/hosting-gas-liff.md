# Hosting, GAS, and LIFF Notes

## Current Shape

- Static pages are hosted outside Google Apps Script, currently intended for Cloudflare Pages.
- Form submissions call Google Apps Script Web App directly through each page's `GAS_URL`.
- LIFF initialization uses each page's `LIFF_ID`.
- Login return routing is handled by `index.html`.
- `router.html` is not part of the active flow.

## Active Forms

- `index.html`: shipping request, action `liff_shipping`.
- `order_form.html`: supply order, action `liff_order`.
- `product_request.html`: product request with optional images, action `liff_product_request`.
- `history.html`: request history, action `liff_history`.

## Manual Environment Switch

`.clasp.json` must stay valid JSON. Do not leave commented script IDs in this file.

To switch clasp targets, edit only `scriptId` manually:

```json
{
  "scriptId": "SCRIPT_ID_HERE"
}
```

Keep the rest of the file unchanged unless the clasp project structure changes.

The HTML files also contain `GAS_URL` and `LIFF_ID`. When switching production/test behavior, update those constants consistently.

## OAuth Authorization

Google OAuth consent cannot be pre-approved by code. Each Apps Script project and deploying account must approve newly required scopes at least once.

Keep all expected scopes declared in `appsscript.json` before production rollout. This prevents surprise scope additions during normal operation.

When a new production script project is prepared:

1. Push the latest code and `appsscript.json`.
2. Run `authorizeRequiredScopes` once from the Apps Script editor.
3. Approve the OAuth consent screen.
4. Deploy a new Web App version.

`authorizeRequiredScopes` intentionally creates and removes a temporary Drive folder so the same Drive write/delete permissions used by product image upload and cleanup are requested without touching real request data.

This is not required for every deploy. Repeat it only when:

- a new Apps Script project is used
- a different deploying account is used
- `appsscript.json` adds new scopes
- authorization was revoked from the Google account security settings

## Failed To Fetch Triage

1. Confirm the page is using the intended `GAS_URL`.
2. POST a small `text/plain;charset=utf-8` JSON body to the GAS `/exec` URL.
3. A healthy Web App should return JSON, even for an unsupported action.
4. If the response is an Apps Script HTML error page, fix the GAS project/deployment before investigating Cloudflare Pages.
5. If the response is JSON from command line but browser fetch fails, then inspect browser console/network details for CORS, mixed environment constants, or LIFF redirect URL issues.

Expected fetch shape:

```js
fetch(GAS_URL, {
  method: 'POST',
  headers: { 'Content-Type': 'text/plain;charset=utf-8' },
  body: JSON.stringify({ action: 'liff_shipping', data: payload })
});
```

Using `text/plain` avoids a CORS preflight for these GAS Web App requests.

## LIFF Routing

Rich menu URLs should enter through the LIFF endpoint and pass the intended page with `page`.

Examples:

- Shipping request: `https://liff.line.me/LIFF_ID/?page=shipping`
- Supply order: `https://liff.line.me/LIFF_ID/?page=order`
- Product request: `https://liff.line.me/LIFF_ID/?page=product`
- Request history: `https://liff.line.me/LIFF_ID/?page=history`

`page=shipping` stays on `index.html` because the shipping request form is the top page. `page=order`, `page=product`, and `page=history` are loaded by the `index.html` SPA entry without changing the visible URL to the individual HTML page. If the SPA load fails, `index.html` falls back to `location.replace()` for the requested page.

Secondary forms call `liff.login({ redirectUri })` with `./index.html?form=...`.

`index.html` redirects based on:

- direct `page` query
- direct `form` query
- `liff.state`
- short-lived `sessionStorage` key `pendingLiffForm`

This keeps routing in one place and avoids a separate static router page.
