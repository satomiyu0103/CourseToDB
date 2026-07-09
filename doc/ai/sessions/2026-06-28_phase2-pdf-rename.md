# セッション記録: Phase 2 PDF リネーム・期限正規化

- 日付: 2026-06-28
- スコープ: FR-PDF-001 / FR-PDF-003 / FR-PDF-004 / FR-PDF-002（期限 YY.MM）

## 背景・課題

- スマホ Drive スキャン後の PDF 名が汎用のまま残り手動リネームが発生していた
- Phase 2 の drive 列挙・移動は未実装のままだった
- 裏面刻印 `28.08/+DFL/B` の期限が Gemini で空になるケースがあった
- Plan 実装を master 直コミットしており、今後のビジネス開発向けにブランチ戦略を見直す必要があった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| リネームタイミング | Gemini 抽出成功後・移動直前 | スキャン直後 |
| ジャンル | 固定マスタ + Gemini 選択 | 自由記述 |
| 日付プレフィックス | 期限日優先、なければ処理日 | 処理日のみ |
| `28.08` 解釈 | YY.MM → 当月末（2028-08-31） | DD.MM（欧州式） |
| Phase 2 完了 | コアパイプライン一段落 | Slack・重複・UI は Phase 2 残 |
| Git | モード A ソロ / モード B 本番。エージェントは B 既定 | master 直コミット常態化 |

## 実装サマリ

- `genreList.js`・`drive.js` リネーム/移動/列挙・`schema` 25 列化
- `debugPocExtractPdf` 等と Script Property `DEBUG_PDF_FILE_ID`
- `normalizeExpirationDateToIso_` で YY.MM 刻印正規化
- `git_workflow.mdc` に Plan 連携ブランチ戦略を追記

## 変更ファイル（主要）

- `gas/src/drive.js` `genreList.js` `main.js` `gemini.js` `schema.js` `spreadsheet.js` `config.js`
- `doc/specs/02_`〜`07_` `06_ROADMAP` Phase 2 一段落
- `.cursor/rules/git_workflow.mdc` `implementation-phase2/SKILL.md`
- `scripts/verify-*.mjs`

## 検証

- GAS: `debugPocExtractFirstInputPdf` — 期限 `2028-08-31`、JAN・商品名抽出 OK
- ローカル: `node scripts/verify-expiration-date.mjs` `node scripts/verify-drive-filename.mjs`
