# セッション記録: 重複キーへ期限日を含める（FR-SHT-002 改訂）

- 日付: 2026-07-01
- スコープ: FR-SHT-002

## 背景・課題

- 同一 JAN（または商品コード）で期限が異なる PDF を再投入すると、2 件目が重複スキップされ別行として残らなかった
- 複合キー経路のみ `expiration_date` を含んでいたが、実運用では JAN 優先が多くボトルネックだった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| 期限比較 | 全キー種別（JAN / 商品コード / 複合）で `expiration_date` を比較 | JAN のみ従来どおり |
| 期限なし | 期限不問として既存行と重複 | 期限ありと別行 |
| 実装形 | `baseValue` + `expirationIso` + `isDuplicateProductRecord_` | 文字列連結のみ |

## 実装サマリ

- `schema.js`: `buildDuplicateBaseValue_`・`isDuplicateProductRecord_`・`formatDuplicateStrategyForLog_`。戦略返却を `{ keyType, baseValue, expirationIso }` に変更
- `spreadsheet.js`: `findExistingProductRow_` をペア比較に接続。ログを新形式に更新
- `scripts/verify-duplicate-check.mjs`: ロジック同期・期限違い/期限なしテスト追加

## 変更ファイル（主要）

- `gas/src/schema.js` `spreadsheet.js`
- `scripts/verify-duplicate-check.mjs`
- `doc/specs/02_`〜`04_` `07_` `03_システム設計.md`
- `doc/reference/setup/product-schema-design.md`
- `gas/README.md`

## 検証

- ローカル: `node scripts/verify-duplicate-check.mjs` — OK（2026-07-01）
- GAS: `npm run clasp:push` — 12 ファイル同期済み（2026-07-01）
- GAS 実機: 同一商品名・期限違い PDF の別行追記を確認（2026-07-01）
- 追記修正: シート Date 型の期限読み取り不具合を `sheetCellToRecordString_` で修正後、再確認 OK
