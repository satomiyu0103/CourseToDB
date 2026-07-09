# セッション記録: ControlPanel サイドバー UI

- 日付: 2026-06-28
- スコープ: FR-PDF-001（サイドバー運用導線）

## 背景・課題

- `ControlPanel.html` は雛形のみで、日常運用はメニュー + alert に限定されていた
- サイドバー初回表示で `script.container.ui` 不足エラー、再認可後も汎用エラーが出るケースがあった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| サイドバー操作 | 一括処理・シート作成・ヘッダー再適用の 3 ボタン | PoC fileId 入力・デバッグボタン |
| `onOpen` | メニューのみ（自動サイドバー表示なし） | 起動時自動表示 |
| API パターン | `menu*` string 返却 + `menu*WithAlert` | サイドバー専用関数の重複実装 |
| 権限 | `appsscript.json` に `script.container.ui` 明示 + `authorizeContainerUi` | スコープ暗黙付与 |

## 実装サマリ

- `ui.js`: `showControlPanelSidebar` / `openControlPanelSidebar_` / `authorizeContainerUi` / `runWithAlert_`
- `ControlPanel.html`: ボタン 3 つ + `#status` 結果表示
- `appsscript.json`: `script.container.ui` スコープ追加

## 変更ファイル（主要）

- `gas/src/ui.js` `gas/ControlPanel.html` `gas/appsscript.json`
- `doc/specs/06_ROADMAP.md` `doc/specs/07_CHANGELOG.md` `gas/README.md`
- `doc/ai_guidelines/試験実装のエラー.md`（サイドバー権限エラー）

## 検証

- `pnpm exec clasp push` — 12 ファイル同期 OK
- 実機: `authorizeContainerUi` で UI 権限承認後、サイドバー表示を確認
- 実機: 一括処理・商品DB シート作成・ヘッダー行再適用の 3 操作をサイドバーから確認（2026-06-28）
