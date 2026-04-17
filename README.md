# GAS マスターコード

## ルール

**`Code.js` が唯一の正（Single Source of Truth）です。**

- GASのコードを修正・追加する場合は、**必ずこのファイルを編集**してください
- 編集後、Google Apps Script エディタにコピペ → 「**新しいバージョン**」でデプロイ
- 各アプリのフォルダにGASコードのコピーを置かないでください

## 対象アプリ

このGASは以下のアプリのバックエンドを兼任しています：

| アプリ | フォルダ | APIプレフィックス |
|--------|---------|-----------------|
| 消防設備点検アプリ | `消防設備点検システム関連/` | `getTodaySites`, `saveSiteDetails`, etc. |
| 校正機器管理システム | `calibration-manager/` | `getAllCalibrationData`, etc. |
| Google Chat通知 | （共通） | `sendChatNotification` |

## デプロイ先

- **スプレッドシート**: 消防設備点検システム_DB
- **GAS Web App URL**: 各アプリの `.env` → `VITE_GAS_API_URL` に記載
- **デプロイ設定**: 「自分として実行」「全員がアクセス可能」

## 更新手順

1. `Code.js` を編集
2. Google Apps Script エディタにコピペ
3. 「デプロイ」→「新しいデプロイ」→ ウェブアプリ → デプロイ
4. ⚠️ 既存バージョンの上書きは不可。必ず「新しいデプロイ」
5. 新しいURLが発行されたら、各アプリの `.env` を更新

## 将来の移行計画

GAS + Spreadsheet → **Supabase** への段階的移行を予定。
移行時もこのフォルダにマイグレーションスクリプト等を配置する。
