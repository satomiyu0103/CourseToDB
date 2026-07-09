# FR-SYNC-014 GCal 初回 import 実装

- 日付: 2026-07-02
- スコープ: FR-SYNC-014
- ブランチ: `feat/FR-SYNC-014-gcal-initial-import`

## 背景・課題

GCal にのみ存在する単発予定を Notion DB へ初回取り込みたい。既存ツールは Notion→GCal 一方向のみ。ID 紐付けは `sync_state.json` のみだと二重作成リスクがある。

## 意思決定

| 項目 | 採用 | 非採用 |
|:---|:---|:---|
| モード | 初回 import のみ（`import-gcal`） | 継続逆同期・真の双方向 |
| 期間 | 今日 00:00 JST 以降（既定） | 過去イベント一括 |
| マッチング | `gcal_id` 列 + `mapped_ids` | タイトル+日時マッチ |
| 状態更新 | `mapped_ids` と **`last_events` 同時更新** | `mapped_ids` のみ |
| TZ | 既存と同様 `timezone(timedelta(hours=9))` | `zoneinfo`（Windows で tzdata 未同梱） |

## 実装サマリ

- CLI: `notion-gcal-sync import-gcal [--dry-run] [--time-min] [--time-max]`
- パッケージ: `gcal_import/`（runner, filters, models, time_range）
- GCal: `fetch_events_for_import` + `gcal_mapping.py`
- Notion: `create_page_from_gcal`, `fetch_existing_gcal_ids`
- 設定: `NOTION_PROP_GCAL_ID`（既定 `gcal_id`）

## 学び

後から `同期check` を ON にしても GCal 二重作成を防ぐには、`mapped_ids` だけでなく `last_events` にも import 行を載せる必要がある。`EventChangeFinder` は `last_events` 不在を created とみなすため。

## 変更ファイル

- `src/notion_gcal_sync/gcal_import/*`（新規）
- `src/notion_gcal_sync/clients/gcal_mapping.py`（新規）
- `src/notion_gcal_sync/clients/gcal_client.py`, `notion_client.py`, `notion_mapping.py`
- `src/notion_gcal_sync/infra/config.py`, `cli.py`
- `tests/unit/test_gcal_*.py`, `tests/fixtures/gcal_event_*.json`
- doc: 04, 07, 01 GAP, 計画書, README, notion-db-運用

## 検証

```powershell
uv run pytest
uv run notion-gcal-sync import-gcal --dry-run
```

実機（Notion create・`sync --dry-run` で created=0）は利用者確認待ち。
