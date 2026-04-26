# Google Sheets 連携

このフォルダの [CallTrackerSync.gs](/mnt/c/users/kawai/onedrive/デスクトップ/count/integrations/google-sheets/CallTrackerSync.gs) を Google Apps Script に貼り付けて使います。

## 手順

1. Google Apps Script を新規作成
2. `Code.gs` の中身を `CallTrackerSync.gs` で置き換える
3. `デプロイ -> 新しいデプロイ`
4. 種類は `ウェブアプリ`
5. 実行ユーザーは自分
6. アクセスできるユーザーは `全員`
7. 発行された `exec` URL をこのサイトの `Apps Script URL` に貼る

## サイト側で入れるもの

- `Apps Script URL`
- `スプレッドシートURL`
- `対象シート名（メンバー名タブ）`
- `Slack Webhook URL（任意）`

月名の入力は不要です。サイトの日付から `2026年4月｜Q1` のような月ブロックを自動で探します。

## 自動同期されるタイミング

- `テンプレートをコピー`
- `Slackへ送信`

## 自動同期される項目

- `コール数 実`
- `2回目架電数 実`
- `担当者接続数 実`
- `サンプル送付数 実`
- `導入数(新規成約) 実`

## Slack送信

- `Slack Webhook URL` を入れると `Slackへ送信` で報告テンプレート本文をチャンネルへ投稿できます
- Slack送信時は同時にスプレッドシート同期も実行します
