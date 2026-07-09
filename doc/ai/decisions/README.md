# 意思決定 横断索引

トピック別に、セッション記録・ADR・ABC 台帳・関連 FR へのリンクを集約する。正本の詳細は各リンク先を参照すること。

## 昇格ルール（要約）

| パターン | 昇格先 |
|---|---|
| 毎タスクの好み・禁止 | `.cursor/rules/` |
| タスク手順 | `.cursor/skills/` |
| 要件・設計 | `doc/specs/02_*.md` `03_*.md` |
| 見送り・ギャップ | `doc/specs/01_ABC見送り・ギャップ台帳.md` |
| 技術選定 | `doc/adr/` |
| 再現性あるエラー | `doc/ai/guidelines/試験実装のエラー.md` |

手順: [design-decision-record](../../../.cursor/skills/design-decision-record/SKILL.md) / [agent-session-record](../../../.cursor/skills/agent-session-record/SKILL.md)

---

## スコープ・運用

| トピック | 決定（1行） | セッション | 仕様・台帳 | FR / GAP |
|---|---|---|---|---|
| 期限接近通知 | Slack。月初 digest + 7 日前。期限管理シート分離 | [sessions/2026-07-01_expiration-alert-plan.md](../sessions/2026-07-01_expiration-alert-plan.md) | [06_ROADMAP#アーカイブ](../../specs/06_ROADMAP.md#アーカイブpoc-計画参照用)、[03_システム設計](../../specs/03_システム設計.md) | FR-NTF-003、FR-SHT-003 |
| 1 PDF 1 商品 | 運用ルールで徹底。複数行展開は見送り | [sessions/2026-07-01_1pdf1product-operational-rule.md](../sessions/2026-07-01_1pdf1product-operational-rule.md) | [01_ABC GAP-A-001](../../specs/01_ABC見送り・ギャップ台帳.md)、[gas-operations](../../reference/setup/gas-operations.md) | GAP-A-001 |
| SmartShelf 命名 | GAS/GCP は SmartShelf。Git リポ名は food-label-pdf-gas 維持 | [sessions/2026-06-28_smartshelf-naming-repo-split.md](../sessions/2026-06-28_smartshelf-naming-repo-split.md) | README、`00_プロジェクト概要` | — |
| 名刺版とのリポ分離 | 別リポ継続。コアのみコピー流用 | 同上 | — | — |

## 技術・ランタイム

| トピック | 決定（1行） | セッション / ADR | FR |
|---|---|---|---|
| GAS 本番ランタイム | Python は補助のみ。本番は `gas/src/` | [adr/0003-gas-runtime.md](../../adr/0003-gas-runtime.md) | — |
| clasp rootDir | `gas/src` を rootDir に統一 | [sessions/2026-06-28_clasp-rootdir.md](../sessions/2026-06-28_clasp-rootdir.md) | — |
| 重複キーと期限日 | 全キー種別で expiration_date 比較 | [sessions/2026-07-01_FR-SHT-002-expiration-dup-key.md](../sessions/2026-07-01_FR-SHT-002-expiration-dup-key.md) | FR-SHT-002 |

## 通知・期限管理

| トピック | 決定（1行） | セッション | 仕様・台帳 | FR |
|---|---|---|---|---|
| 期限接近通知 | Slack。月初 digest + 7 日前。商品DB + 期限管理の二層 | [sessions/2026-07-01_expiration-alert-plan.md](../sessions/2026-07-01_expiration-alert-plan.md) | [06_ROADMAP#アーカイブ](../../specs/06_ROADMAP.md#アーカイブpoc-計画参照用)、[03_システム設計](../../specs/03_システム設計.md) | FR-NTF-003、FR-SHT-003 |

## エージェント運用基盤

| トピック | 決定（1行） | セッション | 昇格先 |
|---|---|---|---|
| rules/skills 再配置 | ガイド正本を `.cursor/` へ。doc は索引 | [sessions/2026-06-28_cursor-rules-skills-migration.md](../sessions/2026-06-28_cursor-rules-skills-migration.md) | `.cursor/rules/` `.cursor/skills/` |
| テンプレ二重同期 | ログ更新時は ai-agent-devenv-template も同期 | [sessions/2026-06-28_log-template-sync-policy.md](../sessions/2026-06-28_log-template-sync-policy.md) | `template_sync.mdc` |
| 利用者表記 | 「人間」「人向け」を使わない | [sessions/2026-06-28_documentation-wording.md](../sessions/2026-06-28_documentation-wording.md) | `documentation_wording.mdc` |
| doc/ai 統合 | 知識層を doc/ai/ に集約 | [sessions/2026-07-01_doc-ai-knowledge-consolidation.md](../sessions/2026-07-01_doc-ai-knowledge-consolidation.md) | `doc/ai/README.md` |
| タスク価値優先 | 要否検証・小さな差分・学びの可視化 | [sessions/2026-07-02_task-value-first-skill.md](../sessions/2026-07-02_task-value-first-skill.md) | `task-value-first` Skill、`engineer_signals_core` Rule |
| Google Drive MCP | dylancaponi/gdrive-mcp-server + readonly。正本は Drive、リポジトリは索引のみ | [sessions/2026-07-02_google-drive-mcp-adoption.md](../sessions/2026-07-02_google-drive-mcp-adoption.md) | `gdrive-mcp` Skill、`gdrive_mcp_core` Rule、`doc/reference/setup/google-drive-mcp.md` |

## 追記ルール

新しい意思決定を記録したら、該当カテゴリの表に 1 行追加する。カテゴリが無ければ見出しを新設する。
