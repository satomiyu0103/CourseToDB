---
name: audit-security
description: セキュリティチェックリストの実行手順。セットアップ時・プロトタイプ後監査で使用。
---

# セキュリティ監査

0. Agentic リスクは先に [agent-owasp-compliance](../agent-owasp-compliance/SKILL.md) で ASI01–10 をスキャンする
1. `doc/ai/guidelines/checklists/security.md` を上から照合
2. 未達を「大分類 / 項目 / 根拠 § / 該当ファイル」付きで列挙
3. 判断が要る項目は根拠 doc を引用して結論を書く
4. 任意: チェックリスト末尾「実施記録」に日付・未達要約
