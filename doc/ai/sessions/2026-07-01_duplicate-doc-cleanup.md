# セッション記録: 重複ドキュメント削除

- 日付: 2026-07-01
- スコープ: 運用基盤（NFR-OPS-002）

## 背景・課題

- `doc/ai/` 統合後も旧パス stub・索引スタブ 6 件・`agent_workflows.mdc` が残り、正本が二重化していた

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| 旧パス | 即時削除（`doc/ai_guidelines/`・`doc/records/`） | リダイレクト stub 維持 |
| 索引スタブ | `doc/ai/guidelines/` から 6 件削除 | `.cursor` との二重索引維持 |
| 後方互換 | `agent_workflows.mdc` 削除 | スタブ残置 |
| 履歴 | sessions 本文・CHANGELOG 過去行はそのまま | 履歴の書き換え |

## 実装サマリ

- ルート・テンプレ双方で旧ディレクトリ・スタブ 6 件・`agent_workflows.mdc` を `git rm`
- `Project_map.md`・`doc/ai/README.md`・`template_sync.mdc`・`AGENTS.md`・`05_ディレクトリ構成.md` 等の参照を `doc/ai/` と `.cursor/` に一本化
- `TPL_実機FB記録.md` の分割先を `doc/ai/sessions/` に更新

## 変更ファイル（主要）

- 削除: `doc/ai_guidelines/**`・`doc/records/**`・`doc/ai/guidelines/` スタブ 6 件・`.cursor/rules/agent_workflows.mdc`
- 更新: `doc/ai/guidelines/Project_map.md`・`doc/ai/README.md`・`doc/specs/07_CHANGELOG.md` 他
- `ai-agent-devenv-template/` 同期一式

## 検証

- 削除パスの `Test-Path` 確認
- `rg` で現行参照 grep（sessions 履歴・CHANGELOG 過去行除く）
