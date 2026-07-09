# セッション記録: memory_logger・template_sync とテンプレ同期

- 日付: 2026-06-28
- スコープ: `NFR-OPS-002`

## 背景・課題

- タスク完了時の再利用ファクトを `.cursor/doc/memory_stream.md` に蓄積する仕組みが必要
- `food-label-pdf-gas` と `ai-agent-devenv-template/` で共有ファイルを二重管理しないよう、テンプレ同期ルールを常時適用化

## 意思決定

- **採用**: `memory_logger.mdc`（alwaysApply）で追記トリガー・フォーマット・50KB 圧縮を定義
- **採用**: `template_sync.mdc`（alwaysApply）で `.cursor/`・`doc/ai_guidelines/`・`AGENTS.md` の同期対象を明示
- **非採用**: テンプレ単体リポジトリの `agent_core` に template_sync 義務行を入れる（埋め込み先のみ）

## 実装サマリ

- 親リポ: `memory_logger.mdc`・`template_sync.mdc`・`memory_stream.md` 新設。`agent_core`・`AGENTS.md`・`Project_map`・`05_ディレクトリ構成` 更新
- テンプレ: 同一ファイルを同期。`template_sync.mdc` に単体リポ向け注記。`TEMPLATE_SETUP`・CHANGELOG 追記

## 変更ファイル一覧

- `.cursor/rules/memory_logger.mdc`（新規）
- `.cursor/rules/template_sync.mdc`（新規）
- `.cursor/doc/memory_stream.md`（新規）
- `.cursor/rules/agent_core.mdc`
- `AGENTS.md`・`doc/ai_guidelines/Project_map.md`・`doc/specs/05_ディレクトリ構成.md`・`doc/specs/07_CHANGELOG.md`
- `ai-agent-devenv-template/` 配下の対応ファイル一式

## 検証

- 親・テンプレとも `.cursor/rules/*.mdc` が 14 本で一致
- `memory_logger.md` 誤参照なし（正本は `.mdc`）
