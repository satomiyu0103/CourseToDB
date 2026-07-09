# Agent 運用フロー

本ファイルは **エージェント向けルーター** です。高レベルのルールを示し、詳細手順は適切な skill / rule へ誘導します。

> **常時**: [`.cursor/rules/agent_core.mdc`](.cursor/rules/agent_core.mdc)  
> **実装時**（`src/`・`doc/` 等）: [`.cursor/rules/agent_implement_entry.mdc`](.cursor/rules/agent_implement_entry.mdc)  
> **Phase 3 詳細**: [`.cursor/skills/phase3-doc-updates/SKILL.md`](.cursor/skills/phase3-doc-updates/SKILL.md)

---

## ルーティング表

| タスク | Skill / Rule |
|------|-------|
| 実装着手（Phase 1） | [implementation-phase1](.cursor/skills/implementation-phase1/SKILL.md) |
| 実装中（Phase 2） | [implementation-phase2](.cursor/skills/implementation-phase2/SKILL.md) |
| 完了・doc 更新（Phase 3） | [phase3-doc-updates](.cursor/skills/phase3-doc-updates/SKILL.md) |
| 価値検証・差分計画 | [task-value-first](.cursor/skills/task-value-first/SKILL.md) |
| CHANGELOG 追記 | [changelog-entry](.cursor/skills/changelog-entry/SKILL.md) |
| セッション記録 | [agent-session-record](.cursor/skills/agent-session-record/SKILL.md) |
| 設計判断の記録 | [design-decision-record](.cursor/skills/design-decision-record/SKILL.md) |
| 既知エラー登録 | [known-error-entry](.cursor/skills/known-error-entry/SKILL.md) |
| 監査（ドキュメント） | [audit-documentation](.cursor/skills/audit-documentation/SKILL.md) |
| 監査（実装） | [audit-implementation](.cursor/skills/audit-implementation/SKILL.md) |
| 監査（運用） | [audit-operations](.cursor/skills/audit-operations/SKILL.md) |
| 監査（プロトタイプ後） | [audit-post-prototype](.cursor/skills/audit-post-prototype/SKILL.md) |
| 監査（リファクタ） | [audit-refactoring](.cursor/skills/audit-refactoring/SKILL.md) / [audit-refactor-full](.cursor/skills/audit-refactor-full/SKILL.md) |
| 監査（セキュリティ） | [audit-security](.cursor/skills/audit-security/SKILL.md) |
| 安全なバッチ・自動化 | [safe-operations-detail](.cursor/skills/safe-operations-detail/SKILL.md) · [safe_operations_core.mdc](.cursor/rules/safe_operations_core.mdc) |
| 日本語技術文の推敲 | [japanese-tech-writing](.cursor/skills/japanese-tech-writing/SKILL.md) |
| Web / UI デザイン方針 | [DESIGN.md](DESIGN.md) · [doc/design/ai-design-brief.md](doc/design/ai-design-brief.md) |
| Git・ブランチ・PR | [git_workflow.mdc](.cursor/rules/git_workflow.mdc) |
| テンプレ同期 | [template_sync.mdc](.cursor/rules/template_sync.mdc) |
| Drive 参照（MCP） | [gdrive-mcp](.cursor/skills/gdrive-mcp/SKILL.md) |
| テンプレ整合性検証 | `python scripts/validate_agent_docs.py` |

### 監査（Agentic セキュリティ）

| タスク | Skill |
|------|-------|
| ASI 準拠チェック（ASI01–10） | [agent-owasp-compliance](.cursor/skills/agent-owasp-compliance/SKILL.md) |
| ツール・ポリシー統制（ASI01/02） | [agent-governance](.cursor/skills/agent-governance/SKILL.md) |
| MCP 設定監査 | [mcp-security-audit](.cursor/skills/mcp-security-audit/SKILL.md) |
| OWASP 公式 ASI 定義 | [owasp-agentic](.cursor/skills/owasp-agentic/SKILL.md) |
| LLM アプリ注入防御 | [llm-security](.cursor/skills/llm-security/SKILL.md) |

## 利用者向けドキュメント

| 文書 | 用途 |
|------|------|
| [doc/初心者ガイド.md](doc/初心者ガイド.md) | テンプレの使い方・Cursor への頼み方 |
| [doc/reference/README.md](doc/reference/README.md) | セットアップ・早見表の索引 |
| [DESIGN.md](DESIGN.md) | UI のデザイン正本 |
| [doc/design/ai-design-brief.md](doc/design/ai-design-brief.md) | デザイン依頼手順・用語カタログ |
| [README.md](README.md) | テンプレ概要 |

---

## クイックリファレンス（仕様・知識）

| やりたいこと | 参照先 |
|---|---|
| Notion DB 列・日付の運用（方針 B） | `doc/reference/notion-db-運用.md` |
| フェーズ計画・定期実行 | `doc/specs/06_ROADMAP.md` · [windows-scheduled-sync.md](doc/reference/setup/windows-scheduled-sync.md) |
| 定期失敗の Slack 通知（Phase 3d） | `doc/specs/06_ROADMAP.md` Phase 3d · `doc/specs/02_要件定義.md`（FR-OPS-001）· `doc/specs/03_システム設計.md` §10 |
| Notion CSV import 実装計画（Phase 2.1） | `doc/specs/plans/notion-csv-import.md` |
| GCal 初回 import 計画（Phase 3b） | `doc/specs/plans/gcal-initial-import.md` |
| 要件確認 | `doc/specs/02_要件定義.md` |
| 機能一覧・実装状況 | `doc/specs/04_機能一覧.md` |
| 設計・フロー確認 | `doc/specs/03_システム設計.md` |
| 過去のエラー確認 | `doc/ai/guidelines/試験実装のエラー.md` |
| 開発ルール・作業ブランチ（`master` 直コミット禁止） | `.cursor/rules/git_workflow.mdc` |
| Notion/GCal 固有の実装規約 | `doc/ai/guidelines/notion-gcal-implementation.md` |
| 運用品質リファクタ（Phase 3a 完了） | `doc/records/plans/ops-quality-refactor.md` · [03_システム設計.md](doc/specs/03_システム設計.md) §9 |
| **運用前提（初回セットアップ済み・再確認不要）** | `doc/reference/setup/初回セットアップ.md` |
| **別 PC 移行・秘密ファイルの運び方** | `doc/reference/setup/PC移行・別PCセットアップ.md` |
| 意思決定の財産化（sessions） | `doc/ai/sessions/README.md` |
| 意思決定・セッション索引 | `doc/ai/decisions/README.md` |
| エージェント知識の入口 | `doc/ai/README.md` |
| AI エージェント向け参照先一覧 | `doc/ai/guidelines/Project_map.md` |
| 記憶ストリーム（ファクト追記） | `.cursor/doc/memory_stream.md`（`.cursor/rules/memory_logger.mdc`） |
| skill / MCP 出典一覧 | [doc/agent/agent-capabilities.md](doc/agent/agent-capabilities.md) |
| チャット用語解説（ジュニア向け） | `.cursor/rules/junior_friendly_explanations.mdc` · `.cursor/skills/junior-friendly-explanations/SKILL.md` |
| コードブロックコメント（ジュニア向け） | `.cursor/rules/code_comments.mdc` · `.cursor/skills/junior-code-comments/SKILL.md`（見本: `clients/gcal_body.py`） |
| テンプレ同期 | `.cursor/rules/template_sync.mdc` — 正本はルート直下 `ai-agent-devenv-template/` のみ |

---

## チャットでのスコープ明示（任意）

別プロジェクトの話題が混ざりやすいとき、依頼の先頭に付ける。

```
[ワークスペースのみ]
このリポジトリ以外・別プロジェクトの話はしないでください。
```

常時の境界定義: [`.cursor/rules/workspace_boundary.mdc`](.cursor/rules/workspace_boundary.mdc)
