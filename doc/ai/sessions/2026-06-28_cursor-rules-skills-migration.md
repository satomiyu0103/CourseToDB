# Agent Session: doc/ai_guidelines → .cursor 再配置

日付: 2026-06-28  
スコープ: ドキュメント整備（NFR-OPS-002 運用・保守）

## 背景・課題

`doc/ai_guidelines/` のガイドラインが 1 ファイルあたりの指示量が多く、Cursor Agent の遵守率低下リスクがあった。Rules（常時制約）と Skills（タスク手順）への分解が必要だった。

## 意思決定

- **採用**: `.cursor/rules/*.mdc`（11 ファイル）+ `.cursor/skills/*/SKILL.md`（17 本）
- **採用**: `doc/ai_guidelines/` ガイド系は索引スタブ化、checklists・試験実装のエラーはデータ正本として doc に残す
- **採用**: `agent_implement.mdc` を廃止し `agent_implement_entry.mdc` に置換
- **非採用**: `ai-agent-devenv-template/` 側の同期更新（本リポジトリのみ）

## 実装サマリ

- Rules: 命名・Git・テスト・外部 API・神クラス監視・安全運用コア等を glob 別に分割
- Skills: implementation-phase1/2、phase3-doc-updates、監査 7 本、CHANGELOG・セッション記録等
- doc スタブ: 実装規約、agent_phase3_dod、エージェント実装記録、リファクタリング判断基準、安全運用ガイド、ai_setup_check_list
- 参照更新: Project_map、AGENTS.md、checklists 全件、05_ディレクトリ構成、git 早見表

## 変更ファイル一覧

- `.cursor/rules/` — 9 新規 + agent_core/agent_workflows 更新、agent_implement.mdc 削除
- `.cursor/skills/` — 17 ディレクトリ新規
- `doc/ai_guidelines/` — 7 スタブ化 + Project_map 更新
- `doc/ai_guidelines/checklists/` — 根拠リンク更新
- `AGENTS.md`, `doc/specs/05_ディレクトリ構成.md`, `doc/specs/07_CHANGELOG.md`
- `doc/reference/cheatsheets/git.md`

- `ai-agent-devenv-template/` — 同構成に同期（`project_identity.mdc` は `{{PROJECT_DESCRIPTION}}` 維持）

## 検証コマンド

```bash
# 旧 agent_implement.mdc が残っていないこと
test ! -f .cursor/rules/agent_implement.mdc

# Skills 数（17 本）
find .cursor/skills -name SKILL.md | wc -l
```
