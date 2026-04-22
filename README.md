# liff-supply-shipping

Netlify + LIFF + GAS で運用する依頼フォームです。  
フロントは Netlify Function を経由して GAS へ送信します。

## Netlify 環境変数

同じキー名で、Production / Deploy Preview / Branch Deploy ごとに値を設定します。

- `GAS_WEBAPP_ID`: 送信先 GAS Web アプリ ID
- `LIFF_ID`: LIFF アプリ ID

`Site settings -> Environment variables` でキーを作成し、各 context に値を割り当ててください。

## GAS Script Properties

GAS 側には機密/運用値を設定します。

- `LINE_TOKEN`
- `MASTER_SHEET_NAME`
- `SUPPLY_SHEET_NAME`
- `TARGET_USER_ID`

`LIFF_ID` はフロント初期化用として Netlify 環境変数側で管理します。

## ルーティング

- フロント送信先: `/api/gas`（`netlify/functions/gas-proxy.js`）
- フロント設定取得: `/env.js`（`netlify/functions/env-js.js`）

`/env.js` は `window.APP_CONFIG` を返し、`index.html` / `order_form.html` が参照します。
