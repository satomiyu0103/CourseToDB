# セッション記録: ログ更新時のテンプレ同期方針

- 日付: 2026-06-28
- スコープ: `NFR-OPS-002`

## 背景・課題

- ログ更新（開発日誌・memory_stream・セッション記録）時にテンプレ正本が追随しない懸念
- テンプレは各プロジェクト開始時の git clone 正本であり、Agent AI 運用をプロジェクト間で揃える目的がある

## 意思決定

- **採用**: 運用基盤（Rules・Skills・memory_stream 雛形・記入ルール）はテンプレ同期、日次データはプロジェクト固有のまま
- **採用**: `template_sync.mdc` に目的とログ更新時の同期表を明記
- **非採用**: 開発日誌エントリ・agent_sessions 本文・memory_stream ファクトのテンプレコピー

## 実装サマリ

- `template_sync` / `memory_logger` / `agent-session-record` / `phase3-doc-updates` / `agent_core` を更新
- テンプレに `doc/records/` 雛形、`00_開発日誌` 記入ルール追記、README・TEMPLATE_SETUP の目的説明を追加

## 変更ファイル一覧

- `.cursor/rules/template_sync.mdc` / `memory_logger.mdc` / `agent_core.mdc`
- `.cursor/skills/agent-session-record/SKILL.md` / `phase3-doc-updates/SKILL.md`
- `ai-agent-devenv-template/` 配下の対応ファイル一式

## 検証

- テンプレ `memory_stream.md` はファクト空の雛形のまま
