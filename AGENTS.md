# AGENTS.md

## Instructions

- コミットメッセージは日本語で書く。
- `.clasp.json` はコメント不可の JSON として維持する。
- 本番/テストの clasp 切替は `.clasp.json` の `scriptId` を手動で書き換える。
- LIFF フロントは Cloudflare Pages などの静的ホスティングから配信し、送信先は各 HTML の `GAS_URL` で指定する。
- `failed to fetch` 調査では、まず `GAS_URL` の `/exec` に `text/plain;charset=utf-8` の POST が JSON を返すか確認する。
- `router.html` は使わない。フォーム振り分けは `index.html` の `form` / `liff.state` / `pendingLiffForm` 処理で行う。

