# セッション記録: task-value-first Skill 新設

- 日付: 2026-07-02
- スコープ: NFR-OPS-002（エージェント運用基盤）

## 背景・課題

- Kent Beck の「タスク完了数ではなく学びと波及で評価される」という先輩エンジニア向けアドバイスを、本テンプレの Phase フローに組み込みたかった
- 既存の `implementation-phase1/2`・`git_workflow`・`agent-session-record` はあるが、着手前の価値検証と完了時の学び抽出が明示されていなかった
- Rule に全文を入れるとコンテキストを圧迫するため、B/C 下限は Rule、A 型の判断手順は Skill に分離する方針とした

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| 新規 Skill | `task-value-first`（A-0 要否・A-1 差分計画・A-2 実施・A-3 完了接続） | 既存 phase1/2 本文への長文追記 |
| 新規 Rule | `engineer_signals_core`（C 回避・B 下限、globs 付き） | `alwaysApply: true` で毎セッション注入 |
| 学びの正本 | `agent-session-record` に `## 学び` 必須項目 | memory_stream のみ |
| 背景文書 | `doc/ai/guidelines/task-value-first.md` | Skill 本文に Beck 全文 |

## 実装サマリ

- `.cursor/skills/task-value-first/SKILL.md` と `checklist.md` を新設
- `.cursor/rules/engineer_signals_core.mdc` を新設
- `agent_implement_entry.mdc` の Phase 表に task-value-first を接続
- `agent-session-record` に `## 学び` 必須項目を追加、`_example_session.md` を更新
- `AGENTS.md`・`Project_map.md`・`decisions/README.md` に索引を追加

## 学び

- キャリアアドバイスをエージェント運用に載せるときは「憲法（Rule）」と「手順（Skill）」と「背景（guideline）」の三層に分けると、既存 Phase フローを壊さず差し込める
- A シグナルの「学びの可視化」は sessions の固定見出しにすると、Phase 3 で毎回抜けにくい
- 昇格候補: なし（本実装で Skill / Rule へ初回昇格済み）

## 変更ファイル（主要）

- `.cursor/skills/task-value-first/SKILL.md`
- `.cursor/skills/task-value-first/checklist.md`
- `.cursor/rules/engineer_signals_core.mdc`
- `.cursor/rules/agent_implement_entry.mdc`
- `.cursor/skills/agent-session-record/SKILL.md`
- `doc/ai/guidelines/task-value-first.md`
- `doc/ai/sessions/_example_session.md`
- `AGENTS.md`
- `doc/ai/guidelines/Project_map.md`
- `doc/ai/decisions/README.md`
- `doc/specs/07_CHANGELOG.md`
- `doc/specs/00_開発日誌.md`
- `.cursor/doc/memory_stream.md`

## 検証

- `agent_implement_entry` の Phase 表から新 Skill・Rule へリンクできること
- `engineer_signals_core` が `alwaysApply: false` かつ `src/**` 等の globs であること
- `_example_session.md` に `## 学び` があること
