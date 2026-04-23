# このリポジトリ向け Copilot 指示

## 前提条件

- **回答は必ず日本語でしてください。**
- コードの変更が必要な場合は、まず計画を立てて実装方法を2つ以上提示して、メリットデメリットを含めて説明してください。
- 変更内容がプロジェクトの規約やアーキテクチャに影響する場合は、必ずその点を指摘してください。

## ビルド・テスト・Lint

このリポジトリには、ローカル実行用の build / test / lint コマンドは定義されていません。プロジェクト直下に `package.json`、`clasp.json`、テストランナー設定もありません。

そのため、ツールが追加されるまでは検証は手動前提です。

- フロントエンド変更時: `index.html` と `order_form.html` をブラウザまたは LIFF 上で確認する
- GAS 変更時: デプロイ済み Web アプリと紐付いたスプレッドシートでリクエスト処理を確認する
- ログイン認証を通せないため自動テストはありません、単体テスト 1 件だけを実行するコマンドも存在しない

## 全体アーキテクチャ

このプロジェクトは、LIFF + Google Apps Script で 2 種類の依頼を扱う小規模ワークフローです。

- `index.html`: 送り依頼フォーム
- `order_form.html`: 備品注文フォーム
- `src\webapp.js`: Apps Script Web アプリの POST エントリーポイント
- `src\core.js`: バリデーション、スプレッドシート書き込み、LINE 通知を行う共通 GAS ヘルパー

ブラウザ側の 2 ページは、独立した LIFF フォームとして動作します。

1. 固定の `LIFF_ID` で LIFF を初期化する
2. LINE プロフィールまたは `localStorage` から依頼者名を初期表示する
3. デプロイ済み Apps Script の `GAS_URL` に JSON を送信する

Apps Script 側は `doPost(e)` のみを受け付けます。受信する JSON は次のどちらかです。

- `action: "liff_shipping"` と、送り依頼フォームに対応した `data`
- `action: "liff_order"` と、`userName` と `items` を含む `data`

`src\webapp.js` は薄いルーターとして実装されており、リクエスト解析・`action` 分岐のみを行い、実処理は `src\core.js` に委譲しています。

`src\core.js` には実際の業務ルールが集約されています。

- `LIFF_ID`、`LINE_TOKEN`、`MASTER_SHEET_NAME`、`SUPPLY_SHEET_NAME`、`TARGET_USER_ID` を Script Properties から取得する
- 受信 payload を検証・正規化する
- 対象シートの 2 行目に新しい依頼を挿入する
- 行の最終列にチェックボックスを追加し、初期値を `false` にする
- シート書き込み成功後に LINE のテキスト通知を送る

スプレッドシートへの書き込みは `SpreadsheetApp.getActiveSpreadsheet()` を使っているため、このプロジェクトは **スプレッドシートにバインドされた Apps Script** 前提です。運用モデルを変えない限り、安易に Spreadsheet ID 指定方式へ変更しないでください。

## 重要な規約

### リクエスト契約

フロントエンドとバックエンドは `action` の文字列と payload 形状で密結合しています。片方を変えるときは、もう片方も同じ変更内で揃えてください。

- 送り依頼フォームは `action: "liff_shipping"` を送る
- 備品注文フォームは `action: "liff_order"` を送る

### シート書き込み列は配列順で決まる

バックエンドは、値をオブジェクトではなく配列としてそのままシートへ書き込みます。

- `toShippingRow_()` が送り依頼シートの列順を定義する
- `toSupplyOrderRow_()` が備品注文シートの列順を定義する

スプレッドシート側の列構成を変える場合は、まずこれらの関数を更新してください。`insertRequestRow_()` は、行配列の最終列に対して常にチェックボックスを追加します。

### 表示名と保存名の扱い

`order_form.html` の備品名は UI 表示用に番号付きですが、`stripNumber_()` によりシート保存時と LINE 通知時には番号を除去します。今後も番号付き表示を続けるなら、この挙動を維持してください。

### GAS 内部ヘルパーの命名

Apps Script の内部ヘルパー関数は、`appendShippingToSheet_` や `normalizeOrderPayload_` のように末尾アンダースコア付きです。`doPost` のようなエントリーポイント以外はこの命名規則に合わせてください。

### 同時実行時の扱い

シート挿入は `LockService.getScriptLock()` で保護され、常に 2 行目へ挿入されます。つまり最新依頼が見出し直下に並ぶ前提です。運用側が明示的に末尾追加を求めない限り、この挙動を維持してください。

### 依頼者名の保存挙動

両方の LIFF ページは `requestUserName` を共通の `localStorage` キーとして使い、あわせて `shippingUserName` もフォールバック参照しています。依頼者名の保持方法は 2 フォーム間で揃えてください。

### エラーハンドリング

GAS のエンドポイントは、失敗時も含めて常に JSON を返します。形式は `{ ok: true }` または `{ ok: false, error }` です。フロント側はこの契約を前提に `error` の文言をそのままアラート表示します。
