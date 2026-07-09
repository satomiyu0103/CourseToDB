---
name: design-decision-record
description: 比較検討・設計のみ・コード未変更時の意思決定記録。agent-session-record と併用。
---

# 設計・意思決定記録

## いつ書くか

| タイミング | 必須 |
|---|---|
| 比較検討・設計のみ・コード未変更 | ○（本 SKILL が主） |
| 実装タスク完了 | ×（[agent-session-record](../agent-session-record/SKILL.md) が主） |
| 1 行修正・typo のみ | × |
| 利用者が「記録不要」と指示 | × |

利用者が「記録して」「決定を残して」と指示したときも作動する。

## 2 層構成

| 層 | パス | 内容 |
|---|---|---|
| 索引 | `doc/specs/00_開発日誌.md` | `## YYYYMMDD` で 3〜10 行 + 詳細リンク |
| 本文 | `doc/ai/sessions/YYYY-MM-DD_トピック.md` | 背景・意思決定・昇格先 |

## 本文の必須項目

1. 日付・スコープ（FR コードまたは「設計検討」）
2. 背景・課題
3. **意思決定（採用 / 非採用）** — 表形式
4. **昇格先** — rules / specs / ABC / ADR / エラー正本 / なし
5. 実装サマリ（コード変更なしの場合は「コード変更なし」と明記）
6. 検証（ドキュメント整合確認等）

ファイル名: `YYYY-MM-DD_{短いトピック}.md`

## 昇格ルール

| パターン | 昇格先 |
|---|---|
| 毎タスク守ってほしい好み・禁止 | `.cursor/rules/*.mdc` |
| タスク種別の手順 | `.cursor/skills/` |
| 要件・設計への影響 | `doc/specs/02_*.md` `03_*.md` |
| 見送り・ギャップ | `doc/specs/01_ABC見送り・ギャップ台帳.md` |
| 技術選定 | `doc/adr/NNNN-*.md` |
| 再現性あるエラー | `doc/ai/guidelines/試験実装のエラー.md` |
| 1 回きりの経緯 | `doc/ai/sessions/` のみ |

昇格した場合は `doc/ai/decisions/README.md` の該当トピック行を追記または更新する。

## 関連 SKILL

| 状況 | 参照 |
|---|---|
| 実装あり | [agent-session-record](../agent-session-record/SKILL.md) |
| エラー記録 | [known-error-entry](../known-error-entry/SKILL.md) |
| CHANGELOG | [changelog-entry](../changelog-entry/SKILL.md) |
| 横断索引 | [doc/ai/decisions/README.md](../../../doc/ai/decisions/README.md) |

## テンプレート同期

- セッション本文・`00_開発日誌.md` 索引追記後、`ai-agent-devenv-template/` が存在すれば [template_sync.mdc](../../rules/template_sync.mdc) に従い同期する
- **本 SKILL の手順変更時**もテンプレを同期する
