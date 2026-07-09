# セッション記録: 重複チェック（FR-SHT-002）

- 日付: 2026-06-28
- スコープ: FR-SHT-002（提案 A-3）

## 背景・課題

- Phase 2 残タスクとして `appendProductRecord_` が常に追記し、同一 JAN の二重登録を防げなかった
- `main.js`・`ui.js` は `skipped` / `duplicate_of_row` を受け取れる状態だったが、spreadsheet 層の検索が未実装だった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| キー優先順 | JAN → 商品コード → 複合キー | 単一キーのみ |
| 重複時 PDF | 処理済みフォルダへ移動継続 | インプットに残す |
| 判定不能（全キー空） | 追記許可 + warn | エラー停止 |
| 重複時通知 | なし（バッチ継続） | Slack 通知 |

## 実装サマリ

- `schema.js`: `normalizeTextForDuplicateCheck_`・`buildCompositeDuplicateKey_`・`resolveDuplicateCheckStrategy_`
- `spreadsheet.js`: `loadProductSheetDataRows_`・`findExistingProductRow_`・`appendProductRecord_` ガード
- `main.js`: スキップ時 WARNING ログ
- `ui.js`: 一括サマリに重複スキップ件数
- `scripts/verify-duplicate-check.mjs`: ローカル単体検証

## 変更ファイル（主要）

- `gas/src/schema.js` `spreadsheet.js` `main.js` `ui.js`
- `scripts/verify-duplicate-check.mjs`
- `doc/specs/04_`〜`07_` `06_ROADMAP` `gas/README.md` `doc/reference/migration/meishi-to-product.md`

## 検証

- ローカル: `node scripts/verify-duplicate-check.mjs` — OK
- GAS: ダミー行追記 — スクリプト正常終了を確認（2026-06-28）
- GAS: メニュー「インプット PDF 一括処理」— 正常動作・重複スキップを確認（2026-06-28）
