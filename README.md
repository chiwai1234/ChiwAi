# マンション巡回チェック

マンションの巡回業務を記録し、リーダーが状況を把握・集計できる、インストール不要の Web ツールです。
`index.html` 1 枚で完結し、GitHub Pages 上でそのまま動作します。

- 公開URL: https://chiwai.jp/

## 特長

- **巡回員ログイン** — Firebase Authentication によるメール／パスワード認証（未設定時はローカルモードで動作）。
- **サーバー不要・インストール不要** — スマートフォンのブラウザで開くだけ。
- **記録は端末内に保存** — 記録はブラウザの `localStorage` に保存されます（外部送信なし）。
- **CSV をメール送信** — 共有シート（設定不要）と、Google Apps Script による自動送信（設定後）の 2 方式。

## 画面構成

### ログイン
- Firebase 設定済みなら、メール／パスワードでログイン（新規登録・パスワード再設定つき）。
- 未設定なら「ローカルモード」（巡回員名の入力のみ）で利用できます。

### 1. 巡回入力
- マンション名・巡回者名・巡回日・時間帯を入力（巡回者はログイン名を自動入力）。
- 各チェックポイントを「異常なし / 要確認 / 異常あり」で記録（状態選択で時刻を自動記録、メモ可）。
- 件数のリアルタイム表示、全体所見・引き継ぎ欄あり。

### 2. 記録・集計（リーダー向け）
- 巡回記録数・要確認あり・異常ありをサマリー表示。マンション名 / 日付 / ステータスで絞り込み。
- **⬇ CSV書き出し** … 端末にダウンロード。
- **✉ メールで送る** … 共有シートで CSV を添付してメール送信（対応外の端末ではダウンロード＋メール作成にフォールバック）。
- **📤 自動送信** … Google Apps Script 設定時、指定アドレスへ自動でメール送信。

### 3. 設定
- アカウント情報 / ログアウト、メール送信の設定状態表示。
- チェックポイントの追加・改名・並べ替え・削除、初期リセット、全記録削除。

---

## セットアップ

`index.html` 冒頭の設定ブロックを書き換えるだけです（コード変更は最小限）。

```js
var FIREBASE_CONFIG = { apiKey:"", authDomain:"", projectId:"", appId:"" };
var GAS_ENDPOINT = "";   // 例: https://script.google.com/macros/s/XXXX/exec
var GAS_TO       = "";   // 例: leader@example.com
```

いずれも未設定のままでも動作します（認証＝ローカルモード、自動送信＝無効。共有シート送信は常時利用可）。

### ① Firebase Authentication（本格ログイン）

1. [Firebase コンソール](https://console.firebase.google.com/) でプロジェクトを作成。
2. 「Authentication」→「ログイン方法」で **メール／パスワード** を有効化。
3. 「Authentication」→「設定」→「承認済みドメイン」に公開ドメイン（例：`chiwai.jp`）を追加。
4. 「プロジェクトの設定」→「マイアプリ（ウェブ）」を追加し、表示される `firebaseConfig` を
   `index.html` の `FIREBASE_CONFIG` に貼り付け。
5. 巡回員のアカウントは、アプリのログイン画面「新規登録」から作成するか、
   Firebase コンソールの Authentication 画面で管理者が追加できます。

> `apiKey` 等はクライアントに公開される前提の公開情報です（秘密鍵ではありません）。
> アクセス制御は承認済みドメインと認証で行います。

### ② Google Apps Script（CSV 自動メール送信）

1. [script.google.com](https://script.google.com/) で新規プロジェクトを作成し、以下を貼り付け。

   ```js
   function doPost(e) {
     var data = JSON.parse(e.postData.contents);
     var to = data.to || 'leader@example.com';
     var filename = data.filename || 'patrol.csv';
     var csv = data.csv || '';
     var blob = Utilities.newBlob('﻿' + csv, 'text/csv', filename);
     MailApp.sendEmail({
       to: to,
       subject: '巡回記録 ' + (data.subject || ''),
       body: '巡回記録CSVを添付します。\n送信者: ' + (data.from || ''),
       attachments: [blob]
     });
     return ContentService
       .createTextOutput(JSON.stringify({ ok: true }))
       .setMimeType(ContentService.MimeType.JSON);
   }
   ```

2. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」。
   - 実行するユーザー：自分
   - アクセスできるユーザー：全員
3. 発行された **ウェブアプリ URL** を `GAS_ENDPOINT`、送信先を `GAS_TO` に設定。

> 送信はブラウザから `no-cors` の POST で行うため、応答は取得しません（送信結果は GAS 側のメールで確認）。

---

## CSV の列

`巡回日, 時間帯, マンション名, 巡回者, ログインユーザー, チェックポイント, 状態, 記録時刻, メモ, 全体所見, 保存日時`

（チェックポイント 1 件につき 1 行の縦持ち形式。Excel 日本語対応の BOM 付き UTF-8。）

## 運用の流れ（例）

1. 巡回員がログインし、現地でスマホから各箇所をチェックして保存。
2. 「記録・集計」でサマリー・異常内容を確認し、「メールで送る」または「自動送信」でリーダーへ CSV を送信。
3. リーダーは受け取った CSV を Excel 等で集計・保管。

## 注意事項

- 記録は**開いた端末・ブラウザ内**にのみ保存されます（端末をまたいだ共有は CSV 送信で行う設計）。
- ブラウザのデータ消去やプライベートモードでは記録が保持されない場合があります。早めに CSV を送信・保管してください。
- 共有シート送信（Web Share API）は HTTPS かつ対応ブラウザ（主にスマートフォン）で利用できます。非対応時はダウンロード＋メール作成にフォールバックします。
