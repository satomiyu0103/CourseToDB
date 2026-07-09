---
name: task-value-first
description: 実装タスク着手前の価値検証と Tidy First 型の差分計画。Phase 1/2、大きな変更前、「本質だけ」「最小で」等の依頼時に使用。
---

# Task Value First

## 目的

実装タスクで「最短完了」ではなく、要否検証・本質の特定・小さな差分・学びの可視化を行う。
背景: [doc/ai/guidelines/task-value-first.md](../../../doc/ai/guidelines/task-value-first.md)

## いつ読むか

| タイミング | 必須 |
|---|---|
| Phase 1（implementation-phase1）完了後、Phase 2 着手前 | ○ |
| 差分 300 行超の見込み | ○ |
| 利用者が「最小で」「本質だけ」と指示 | ○ |
| 1 行 typo・パス明示の限定変更 | × |

## A-0: 要否検証

1. `doc/specs/02_要件定義.md` と対象 FR の AC を読む
2. **やらなくてよい理由**（重複実装・既存で足りる・スコープ外）があるか検討する
3. 見送る場合は [design-decision-record](../design-decision-record/SKILL.md) で記録し、実装を止める
4. やる場合、**価値の 90% を生む 10%** を 1 文で書く

## A-1: 差分計画（着手前）

1. **構造変更**（難しい変更）と**振る舞い変更**（易しい変更）に分ける
2. コミット列を先に列挙する（[git_workflow](../../rules/git_workflow.mdc) §PR 粒度・atomic に準拠）
3. 各コミットが単独でレビュー可能か確認する
4. 着手方針をチャットに 1 段落で共有する

コミット列の例:

```text
1. refactor: 抽出先モジュールを追加（振る舞い不変）
2. feat: FR-XXX-NNN 本質の変更のみ
3. test: FR-XXX-NNN の検証追加
```

## A-2: 実施中

- 最低 **2 案**を比較し、採用理由を 1 行残す（チャットまたはセッション）
- 公式タスクを遅らせるスコープ外改善はしない（[implementation-phase2](../implementation-phase2/SKILL.md) §要件外）
- [engineer_signals_core](../../rules/engineer_signals_core.mdc) の C シグナル禁止に従う
- 既知エラーは [試験実装のエラー.md](../../../doc/ai/guidelines/試験実装のエラー.md) を先に確認する

## A-3: 完了時（Phase 3 接続）

1. [agent-session-record](../agent-session-record/SKILL.md) の **`## 学び`** を必ず 1 段落書く
2. 繰り返し参照されそうな内容は [design-decision-record](../design-decision-record/SKILL.md) §昇格ルールに従い rules / skills 等へ昇格を検討する

## チェックリスト

着手前・実施中・完了時: [checklist.md](checklist.md)

## 省略可

1 行 typo・利用者がパスを明示した限定変更 → 本 SKILL は読まない。

## テンプレート同期

本 SKILL の手順・必須項目を変更した場合、`ai-agent-devenv-template/` が存在すれば [template_sync](../../rules/template_sync.mdc) に従い同一ターン内に同期する。
