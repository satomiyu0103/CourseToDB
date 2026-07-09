# リファクタ前後 — 監査

最終更新: 2026-05-30

> 不定期の設計見直し・大規模整理前後に使う。実装中の神クラス検知はガイドを都度参照。

## 実施手順

1. **リファクタ前**: [refactoring.md](refactoring.md) で予兆とタイミングを記録する。
2. **リファクタ中**: [implementation.md](implementation.md) の PR・テスト・禁止事項を守る。
3. **リファクタ後**: [documentation.md](documentation.md) で specs・CHANGELOG・既知エラーを更新する。
4. バッチ・GAS パイプラインに触れた場合は [operations.md](operations.md) を再実行する。

## 監査順（大分類）

| 順 | タイミング | ファイル |
|:---:|---|---|
| 1 | 前 | [refactoring.md](refactoring.md) |
| 2 | 中 | [implementation.md](implementation.md) |
| 3 | 後 | [documentation.md](documentation.md) |
| 4 | 後（該当時） | [operations.md](operations.md) |

## 総合ゲート

- [ ] リファクタ前の「推奨タイミング」と実施理由が記録にある
- [ ] 振る舞い変更がある場合、unit / integration が通っている（[testing_rules.mdc](../../../.cursor/rules/testing_rules.mdc)）
- [ ] FR コード付き CHANGELOG 行がある（[changelog-entry/SKILL.md](../../../.cursor/skills/changelog-entry/SKILL.md)）

---

## 実施記録

| 日付 | スコープ | 前（予兆数） | 後（未達 doc） |
|---|---|:---:|---|
| | | | |
