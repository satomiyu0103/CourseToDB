---
name: phase3-doc-updates
description: 実装完了後の doc 更新トリガー判定と更新手順。Phase 3 DoD の正本。
---

# Phase 3 — 完了後 DoD

**該当する doc のみ**更新する（全件更新しない）。

## 更新トリガー一覧

| トリガー | 更新先 | 必須 |
|---|---|:---:|
| 実装・修正・doc 整備完了 | `doc/specs/07_CHANGELOG.md` `[Unreleased]` | ○ |
| チャットで実装完了 | `doc/ai/sessions/` + `00_開発日誌.md` 索引 | ○ |
| memory_stream 追記 or 上記セッション記録 | `ai-agent-devenv-template/` へ [template_sync](../../rules/template_sync.mdc) で同時コピー | ○ |
| [template_sync](../../rules/template_sync.mdc) 同期対象（`.cursor/`・`doc/ai/`・`AGENTS.md` 等）の変更 | `ai-agent-devenv-template/` へ **編集直後・完了報告前** に同期 | ○ |
| 利用者の「終了」等（セッション完了） | [memory_logger](../../rules/memory_logger.mdc) の締めで commit / push（両リポジトリ） | ○ |
| 新機能 or 実装状況変更 | `doc/specs/04_機能一覧.md` | 条件 |
| ファイル・ディレクトリ追加削除移動 | `05_ディレクトリ構成.md` + `src/` README ツリー | 条件 |
| 要件・MoSCoW 変更 | `02_要件定義.md` | 条件 |
| 設計・API・画面変更 | `03_システム設計.md` | 条件 |
| 再利用価値あるエラー | `doc/ai/guidelines/試験実装のエラー.md` | 条件 |
| 技術選定の決定 | `doc/adr/NNNN-タイトル.md` 新規 | 条件 |
| doc 新規 md or 索引変更 | `doc/ai/guidelines/Project_map.md` / `doc/ai/README.md` | 条件 |
| 意思決定の横断索引 | `doc/ai/decisions/README.md` | 条件 |
| doc 表記（利用者の呼び方） | `.cursor/rules/documentation_wording.mdc` に準拠 | ○ |
| 日本語の技術文（段落以上の執筆・推敲） | `.cursor/rules/japanese_tech_writing.mdc` に従い Skill を読了 | 条件 |
| ログ運用基盤の変更 | 上記 [template_sync](../../rules/template_sync.mdc) 同期対象行を参照 | ○ |
| 生成物・秘密ファイルパターン増 | `.gitignore` | 条件 |
| フェーズ計画変更 | `06_ROADMAP.md` | 条件 |
| 日誌等肥大 | `.cursor/skills/records-split/SKILL.md` | 条件 |

## 必須サブスキル

1. `.cursor/skills/changelog-entry/SKILL.md`
2. `.cursor/skills/agent-session-record/SKILL.md`（チャット実装時）または `design-decision-record`（設計のみ時）
3. 条件該当時: `known-error-entry` / `records-split`。意思決定昇格時: `doc/ai/decisions/README.md`
