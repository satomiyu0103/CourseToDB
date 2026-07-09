# セッション記録: INPUT 空化（処理成功時）

- 日付: 2026-07-03
- スコープ: FR-JOB-003

## 背景・課題

- HTML 取込成功後も INPUT フォルダにファイルが残る
- v2026.1 受入時は `PROCESSED_FOLDER_ID` 未設定の説明欄フラグ運用で、成功後も INPUT に物理的に残る
- `gas/src/*.js` が `.gs` と二重デプロイされ、クラウド上の挙動が正本と乖離するリスクがあった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| 成功後の INPUT 除去 | `PROCESSED_FOLDER_ID` 必須 + `moveTo` のみ | 説明欄フラグフォールバック |
| 重複・日次上限スキップ | INPUT に残す（現仕様維持） | スキップ時も移動 |
| デプロイ正本 | stale `.js` 削除 + `.clasp.json` は `.gs` のみ | `.js` 併用継続 |

## 実装サマリ

- `Config.gs`: `PROCESSED_FOLDER_ID` 未設定・INPUT 同一 ID 時に起動エラー
- `Pipeline.gs`: `markFileProcessed_` を `moveTo` のみに簡素化。後処理ログを「PROCESSED へ移動」に統一
- `Main.gs`: `smokeTestConfig` から説明欄モード分岐を削除
- stale `gas/src/*.js` 7 件削除、`.clasp.json` の `scriptExtensions` を `[".gs"]` に限定
- `gas/README.md`・要件・設計・機能一覧・`.env.example` を更新

## 学び

- 「処理済み」と「INPUT から消えた」は別概念。説明欄フラグは重複判定に使わず、INPUT 空化にはフォルダ移動が必須
- clasp で `.js` と `.gs` を同時 push するとグローバル `var` が上書きされうる。正本は `.gs` に一本化する

## 変更ファイル（主要）

- `gas/src/Config.gs`, `Pipeline.gs`, `Main.gs`
- `.clasp.json`
- `gas/README.md`
- `doc/specs/02_要件定義.md`, `03_システム設計.md`, `04_機能一覧.md`, `07_CHANGELOG.md`
- `config/.env.example`

## 検証

### 自動・半自動

- コード静的確認: `markFileProcessed_` は `processedFolderId` 必須前提で `moveTo` のみ
- stale `.js` 削除後、`gas/src/` に `.gs` のみが Pipeline / Config / Main を定義

### 利用者実施（GAS エディタ / スプレッドシート）

1. Script Properties に `PROCESSED_FOLDER_ID` を設定（未設定なら `loadRuntimeConfig` がエラー）
2. `pnpm run gas:push` 後、`smokeTestConfig` → INPUT / PROCESSED アクセス OK
3. INPUT に未処理 HTML 1 件 → 取込 → **INPUT 空**・PROCESSED にファイル・シート 1 行
4. 同ファイルを INPUT に戻して再実行 → スキップ（INPUT に残るのは OK）
