# セッション記録: PROCESSED フォルダ移動運用への切替

- 日付: 2026-07-02
- スコープ: FR-JOB-003 / NFR-OPS-001

## 背景・課題

- v2026.1 受入時は `PROCESSED_FOLDER_ID` 未設定のため、処理済みファイルは Drive 説明欄 `[processed:...]` フラグで運用していた
- `INPUT_FOLDER_ID` を変更し `PROCESSED_FOLDER_ID` を Script Properties に追加。フォルダ移動モードへ切替
- コード側の対応要否を確認する必要があった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| Pipeline 後処理 | 既存 `markFileProcessed_` の `moveTo`（設定済み時） | 新規後処理ロジックの追加 |
| INPUT 走査時の説明欄スキップ | 見送り（新 INPUT フォルダで運用開始） | `[processed:` 説明欄の自動スキップ |
| 設定の正本 | GAS Script Properties（`.env` は控え） | `.env` からの自動同期 |

## 実装サマリ

- **GAS 本体の変更は `Main.gs` の `smokeTestConfig` 拡張のみ**（PROCESSED ログ・同一 ID 検証・Drive アクセス確認）
- `gas/README.md` にフォルダ移動モードの実機確認手順を追記
- `06_ROADMAP.md`・`07_CHANGELOG.md` に運用切替を記録

## 学び

- `INPUT_FOLDER_ID` / `PROCESSED_FOLDER_ID` は `Config.gs` で既に読み込み済みのため、設定変更だけなら Pipeline 改修は不要
- `clasp run` は API 実行可能デプロイが別途必要。設定確認は GAS エディタで `smokeTestConfig` を実行する運用が現実的
- 昇格候補: なし

## 変更ファイル（主要）

- `gas/src/Main.gs`
- `gas/README.md`
- `doc/specs/06_ROADMAP.md`
- `doc/specs/07_CHANGELOG.md`

## 検証

### 自動・半自動（実施済み）

- `pnpm run gas:push` — `Main.gs` をクラウドへ反映
- `Config.gs` / `Pipeline.gs` の静的確認 — `processedFolderId` 設定時に `markFileProcessed_` が `moveTo` する

### 利用者実施（GAS エディタ / スプレッドシート）

1. GAS エディタで `smokeTestConfig` 実行 → ログに INPUT / PROCESSED のアクセス OK
2. 新 INPUT に未処理 1 件 → メニュー「未処理ファイルを取り込む」→ PROCESSED へ移動
3. 同ファイルを INPUT に戻して再実行 → スキップ（重複）
