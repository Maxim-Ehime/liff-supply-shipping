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

Secondary forms call `liff.login({ redirectUri })` with `./index.html?form=...`.

`index.html` redirects based on:

- direct `form` query
- `liff.state`
- short-lived `sessionStorage` key `pendingLiffForm`

This keeps routing in one place and avoids a separate static router page.
