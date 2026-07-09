---
name: implementation-phase1
description: 実装着手前に読む doc の選定と確認手順。src/doc/tests/scripts/config 変更タスクの Phase 1 で使用。
---

# Phase 1 — 着手前

## 手順

1. タスク内容から下表の追加参照を選び、**存在確認後に読む**
2. 読めないファイルがあれば中断しユーザーに通知
3. 選んだ doc の要点を 1〜3 行で要約してから実装に入る

## 追加参照マトリクス

| 作業内容 | 参照 |
|---|---|
| 設計・モデル変更 | `doc/specs/03_システム設計.md` |
| ディレクトリ・ファイル追加 | `doc/specs/05_ディレクトリ構成.md` |
| UI/テンプレート変更（Django 等） | `doc/reference/setup/django.md` |
| セキュリティ確認（セットアップ） | `.cursor/skills/audit-security/SKILL.md` |
| テスト・動作確認 | `doc/Information.md` |
| バッチ・パイプライン・通知・上限 | `.cursor/skills/safe-operations-detail/SKILL.md` |
| プロトタイプ後・リファクタ監査 | `doc/ai/guidelines/checklists/README.md`（通常実装では読まない） |

## 省略可

1 行修正・typo・ユーザーがパスを明示した限定変更 → `project_identity.mdc` の必読 3 件も省略可。
