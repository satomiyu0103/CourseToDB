---
name: agent-session-record
description: チャット実装完了時のセッション記録（2 層: 開発日誌索引 + sessions 本文）。
---

# エージェント実装記録

## 目的

チャット上の実装・設計の **経緯と理由** を `doc/ai/sessions/` に残し、プロジェクトの財産とする。
`memory_stream` は短期ファクト用であり、採用/非採用の比較や検証手順は本 SKILL で sessions に書く。
Beck 型 A シグナル「学びの可視化」は本文の **`## 学び`** が正本（[task-value-first](../task-value-first/SKILL.md) A-3 と接続）。

詳細: [doc/ai/sessions/README.md](../../../doc/ai/sessions/README.md)

## いつ書くか

| タイミング | 必須 |
|---|---|
| 実装タスク完了（Phase 3 同時） | ○ |
| 設計のみ・コード未変更 | ×（[design-decision-record](../design-decision-record/SKILL.md) を参照） |
| 1 行修正・typo のみ | ×（CHANGELOG のみ） |

## 2 層構成

| 層 | パス | 内容 |
|---|---|---|
| 索引 | `doc/specs/00_開発日誌.md` | `## YYYYMMDD` で 3〜10 行 + 詳細リンク |
| 本文 | `doc/ai/sessions/YYYY-MM-DD_トピック.md` | 背景・意思決定・変更一覧・検証 |

## 本文の必須項目

1. 日付・スコープ（FR コード）
2. 背景・課題
3. 意思決定（採用 / 非採用）
4. 実装サマリ
5. **学び** — 本タスクから得た再利用可能な知見 1 段落（なぜその設計か、次に効く原則）。昇格候補があれば 1 行（rules / skills / 試験実装のエラー 等）
6. 変更ファイル一覧
7. 検証コマンド

ファイル名: `YYYY-MM-DD_{短いトピック}.md`

## 昇格

意思決定が繰り返し参照される場合は [design-decision-record](../design-decision-record/SKILL.md) の昇格ルールに従い `doc/ai/decisions/README.md` へ索引を追記する。

## テンプレート同期

- セッション本文・`00_開発日誌.md` の索引追記後、`ai-agent-devenv-template/` が存在すれば [`.cursor/rules/template_sync.mdc`](../../rules/template_sync.mdc) に従い **同一ターン内・同一パスでコピー**する（省略不可）
- 利用者が「終了」等でセッションを閉じるときは [memory_logger.mdc](../../rules/memory_logger.mdc) の **締め** で commit / push まで行う
- **本 SKILL の手順・必須項目・パス構造**を変更した場合、または `doc/ai/README.md` / `00_開発日誌.md` の **記入ルール** を変更した場合も、完了報告前にテンプレを同期する
