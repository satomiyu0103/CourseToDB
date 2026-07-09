# 定期監査チェックリスト（索引）

最終更新: 2026-06-28

> **正本**: 原則・例外・トレードオフは `.cursor/rules/` と `.cursor/skills/` に書く。  
> 本フォルダは **観察可能な Yes/No** の検証項目のみ。判断が必要なときは各項目の「根拠」リンク先を読む。

## いつ使うか

| タイミング | 使うファイル |
|---|---|
| プロトタイプ完成後の総点検 | [post_prototype_audit.md](post_prototype_audit.md) |
| 不定期リファクタ・設計見直し前後 | [refactor_audit.md](refactor_audit.md) |
| カテゴリ単体の棚卸し | 下表の大分類ファイル |

**通常の実装・PR 時**は Phase 1〜2 の Skills を読む。本フォルダは **常時必読にしない**（[`.cursor/rules/agent_implement_entry.mdc`](../../../.cursor/rules/agent_implement_entry.mdc) 参照）。

## 大分類一覧

| 大分類 | ファイル | 根拠ガイド |
|---|---|---|
| セキュリティ | [security.md](security.md) | （項目ごとに specs / `.cursor/rules/` を参照） |
| 実装・品質 | [implementation.md](implementation.md) | `.cursor/rules/git_workflow.mdc` 等 |
| 運用・バッチ | [operations.md](operations.md) | `.cursor/rules/safe_operations_core.mdc` |
| リファクタリング | [refactoring.md](refactoring.md) | `.cursor/rules/god_class_watch.mdc` |
| ドキュメント・完了 | [documentation.md](documentation.md) | `.cursor/skills/phase3-doc-updates/SKILL.md` |

## AI への渡し方（例）

```text
.cursor/skills/audit-refactor-full/SKILL.md に従い、
doc/ai/guidelines/checklists/refactor_audit.md を上から照合してください。
```

## 実施記録（任意）

監査実施時は、該当チェックリスト末尾の「実施記録」に日付・実施者・未達の要約を追記する。
