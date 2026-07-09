# FR-SYNC-015 定期 GCal import

日付: 2026-07-02  
スコープ: **FR-SYNC-015**（定期ランチャーへの `import-gcal` 組み込み）

## 背景・課題

- FR-SYNC-014 で `import-gcal` は手動 CLI として実装済みだったが、定期 `sync`（FR-SYNC-009）には含まれていなかった。
- GCal のみに追加された予定を Notion に載せるには、毎日の手動実行が必要だった。

## 意思決定

| 採用 | 非採用 |
|:---|:---|
| `07` / `05` bat で `import-gcal` → `sync` を同一タスクで順実行 | 別タスク・別時刻登録 |
| import 失敗時も sync を続行（best-effort） | import 失敗で sync をスキップ |
| 既存 `GcalImportRunner` をそのまま再利用 | 新 CLI サブコマンド |
| `notifier.py` に `gcal_import_failed` 検知を追加 | 通知専用コマンドの新設 |

継続逆同期（GCal 変更の継続反映）は引き続き見送り（GAP-A-001）。

## 実装サマリ

- `07_GCal同期_定期_無人.bat` / `05_gcal_sync_live.bat`: `import-gcal` → `sync`、いずれか失敗で終了コード 1 + Slack
- `infra/notifier.py`: `gcal_import_failed` マーカー、import 用 phase ラベル・Slack タイトル
- `gcal_import/runner.py`: `gcal_import_started` ログ

## 学び

定期運用へ既存 CLI を載せるときは、**bat の直列実行 + 終了コードの合成**で足りることが多い。取り込みロジックを sync エンジンに混ぜず、冪等な `import-gcal` をそのまま呼ぶと FR-SYNC-014 のテスト資産を維持できる。Slack 通知は import 失敗直後の `unexpected_error` より `gcal_import_failed` を優先抽出すると、利用者向けの段階説明が正確になる。

## 変更ファイル一覧

- `scripts/launcher/07_GCal同期_定期_無人.bat`
- `scripts/launcher/05_gcal_sync_live.bat`
- `src/notion_gcal_sync/infra/notifier.py`
- `src/notion_gcal_sync/gcal_import/runner.py`
- `tests/unit/test_notifier.py`
- `doc/specs/plans/gcal-scheduled-import.md`（新規）
- `doc/specs/04_機能一覧.md`, `02_要件定義.md`, `06_ROADMAP.md`, `01_ABC見送り・ギャップ台帳.md`, `07_CHANGELOG.md`
- `doc/reference/notion-db-運用.md`, `doc/reference/setup/windows-scheduled-sync.md`
- `scripts/launcher/README.md`
- `doc/specs/plans/gcal-initial-import.md`（Out of scope 相互リンク）

## 検証

```powershell
uv run pytest
# 実機（利用者）: 07 手動実行 → logs/file.log に gcal_import_finished → sync_finished
```
