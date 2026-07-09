# src 全体リファクタ（NotionClient 責務分割）

日付: 2026-07-06  
スコープ: **NFR-OPS-002**

## 背景・課題

- `NotionClient` が 492 行・メソッド 24 超で神クラス閾値を超過（sync / import / delete / repair が1クラスに集中）
- `JST`・`resolve_project_root`・日付 ISO 変換が複数モジュールに重複
- モジュール docstring 不足（`cli.py`, `event_json.py` 等）

## 意思決定

| 案 | 採用 |
|---|---|
| `NotionClient` を `clients/notion/` へ責務分割し facade を維持 | 採用 |
| `SyncEngine` 分割 | 非採用（過去判断を維持） |
| `gcal_apply.py` → `gcal_types.py` リネーム | 非採用（import 破壊を避け docstring で明記） |
| `import_key` / `page_import_key` を `notion_mapping` へ移動 | 採用（`duplicate_key` は re-export） |

## 実装サマリ

- 新設 `clients/notion/`: `transport`, `sync_reader`, `delete_confirm`, `page_ops`, `csv_import_writer`, `gcal_import_writer`
- `notion_client.py` は公開 API 不変の facade（委譲のみ）
- `domain/notion_datetime.format_notion_datetime_iso` 集約、`JST` 正本化
- `logger` / `SyncEngine` は `config.resolve_project_root` を使用
- `GCalClient.fetch_events_for_import` を `_list_events_page` に分割

## 学び

大きな API クライアントは **transport（SDK・ページネーション）と feature writer（import/delete）** に切ると、runner 側の import は facade 固定のまま内部だけ差し替えられる。テストは `_transport.retrieve_page` のように分割後の境界を patch する。

## 変更ファイル一覧

- `src/notion_gcal_sync/clients/notion/*`（新設 7 モジュール）
- `src/notion_gcal_sync/clients/notion_client.py`（facade 化）
- `src/notion_gcal_sync/clients/notion_mapping.py`（import キー移動）
- `src/notion_gcal_sync/domain/notion_datetime.py`, `models.py`, 他 docstring
- `tests/unit/test_notion_tombstone.py`, `test_notion_create.py`
- `doc/specs/05_ディレクトリ構成.md`, `03_システム設計.md`, `07_CHANGELOG.md`

## 検証

```powershell
uv run pytest -q
```

結果: 84 passed
