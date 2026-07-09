# セッション記録: doc/ai 統合と意思決定ログ運用

- 日付: 2026-07-01
- スコープ: 運用基盤（NFR-OPS-002）

## 背景・課題

- エージェント知識が `.cursor/`・`doc/ai_guidelines/`・`doc/records/agent_sessions/` に分散していた
- 設計のみの意思決定記録に SKILL がなく、比較検討の横断索引もなかった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| 知識層の置き場 | `doc/ai/`（guidelines・sessions・decisions） | `.cursor/doc/ai/` への統合 |
| 実行層 | `.cursor/` 維持（rules・skills・memory_stream） | rules/skills の doc 配下移動 |
| 設計セッション記録 | `design-decision-record` SKILL 新設 | agent-session-record のみ |
| 旧パス | リダイレクト stub 残置 | 即時削除 |

## 実装サマリ

- `doc/ai_guidelines/` → `doc/ai/guidelines/`、`doc/records/agent_sessions/` → `doc/ai/sessions/` に移動
- `doc/ai/README.md`・`runtime.md`・`decisions/README.md` 新設
- `design-decision-record` SKILL 新設。`agent-session-record`・`template_sync`・`phase3`・`memory_logger` 更新
- `00_開発日誌.md` に記入ルール追記。`ai-agent-devenv-template/` へフル同期

## 変更ファイル（主要）

- `doc/ai/**`（新規・移動）
- `.cursor/skills/design-decision-record/SKILL.md`
- `.cursor/skills/agent-session-record/SKILL.md`
- `.cursor/rules/template_sync.mdc`
- `doc/specs/00_開発日誌.md`・`05_ディレクトリ構成.md`・`07_CHANGELOG.md`
- `doc/ai/guidelines/Project_map.md`・`AGENTS.md`
- `ai-agent-devenv-template/`（同期一式）

## 検証

- 旧パス参照の grep 確認（stub・履歴セッション除く）
- テンプレ `project_identity.mdc` の `{{PROJECT_DESCRIPTION}}` プレースホルダ維持を確認
