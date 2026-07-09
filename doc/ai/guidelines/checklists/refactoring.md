# チェックリスト — リファクタリング

最終更新: 2026-06-28

> 根拠ガイド: [`.cursor/rules/god_class_watch.mdc`](../../../.cursor/rules/god_class_watch.mdc) / [`.cursor/skills/refactoring-report/SKILL.md`](../../../.cursor/skills/refactoring-report/SKILL.md)  
> **2 項目以上** 該当で神クラス予兆として報告。リファクタ実施はユーザー指示まで行わない。

---

## A. 定量

- [ ] 1 クラスのメソッド数が **10 個超**  
  - 根拠: [god_class_watch.mdc](../../../.cursor/rules/god_class_watch.mdc)
- [ ] 認証・権限判定を引数に取るメソッドが **5 個以上**  
- [ ] 1 関数・メソッドが **30 行超**（ビジネスロジックの混入）  
- [ ] 同じ権限判定パターンが **3 箇所以上** に散在  

## B. 定性

- [ ] 「データ保持」「権限判定」「状態遷移」のうち **2 種類以上** が 1 クラスにある  
  - 根拠: [god_class_watch.mdc](../../../.cursor/rules/god_class_watch.mdc)
- [ ] テストを書くとき、対象クラスだけでは完結せず複数クラスが必要  
- [ ] 仕様変更のたびに **同じファイルを何度も** 開いている  

## C. タイミング判断（監査メモ）

該当するものにチェック（複数可）:

- [ ] 同じ種類の変更が **2 回以上** 発生した → **今すぐ検討**  
- [ ] 新機能追加前に「どこに書けばいいか分からない」  
- [ ] 上記 A/B で 2 項目以上該当 → **次の機能追加前までに**  
- [ ] PoC・捨て前提コード → リファクタ **対象外**  

  - 根拠: [refactoring-report/SKILL.md](../../../.cursor/skills/refactoring-report/SKILL.md)

---

## 報告テンプレ（AI 用）

```
[神クラス予兆] `クラス名` に以下の予兆:
- （該当項目）

推奨タイミング: 今すぐ / 次の機能追加前 / 将来的に
根拠: .cursor/skills/refactoring-report/SKILL.md
```

---

## 実施記録

| 日付 | 対象モジュール | 該当数 | 推奨タイミング |
|---|---|:---:|---|
| 2026-07-06 | `NotionClient`（src 全体リファクタ着手前） | A:2 / B:2 | 今すぐ（分割実施） |
| 2026-07-06 | `NotionClient` 分割後・命名監査 | 0 | — |

### 2026-07-06 命名監査結果

- 旧名（`differ`, `Tombstone`, `event_codec`, `dedup`, `csv_resolver`）: src に残存なし
- 禁止単体語彙: 問題なし（`data/` パス・requests `data=`・csv ファイル `handle` のみ）
- 変更しない境界: パッケージ名 `notion_gcal_sync`、CLI サブコマンド、環境変数、SDK JSON キー
- `gcal_apply.py`: リネームせずモジュール docstring で「型と Stats のみ」と明記

### 2026-07-06 未達一覧（着手前）

| 大分類 | 項目 | 該当ファイル |
|---|---|---|
| A 定量 | メソッド 10 超 | `clients/notion_client.py`（24 メソッド） |
| A 定量 | 関数 30 行超 | `notion_client._fetch_events_split`, `_classify_delete_reason`; `gcal_client.fetch_events_for_import` |
| B 定性 | 仕様変更のたび同ファイル | `notion_client.py`（sync / import / delete / repair） |
| B 定性 | テストが複数責務にまたがる | `test_notion_tombstone`, `test_notion_create` |
| 横断 | `JST` 重複 | `csv_reader`, `duplicate_key`, `gcal_import/time_range`, `notifier`（`notion_datetime` が正本） |
| 横断 | `resolve_project_root` 重複 | `infra/logger.py`, `sync/engine.py` |
| 横断 | モジュール docstring 欠落 | `cli.py`, `gcal_body.py`, `event_json.py` 他 |
