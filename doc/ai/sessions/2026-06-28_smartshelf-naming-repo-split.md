# Agent Session: SmartShelf 名称明記とリポジトリ分離方針

日付: 2026-06-28

## スコープ

- 設計検討（FR コード未割当）
- ドキュメント更新（`NFR-OPS-002` 運用・保守）

## 背景・課題

- 名刺 PDF 版と食品 PDF 版を同一リポジトリにまとめるか、別プロジェクトとするかの判断が必要だった
- GAS / GCP のプロジェクト名は **SmartShelf**、Git リポジトリ名は `food-label-pdf-gas` で不一致

## 意思決定

| 案 | 採用 |
|---|---|
| 名刺版を拡張して食品を追加 | 非採用（ドメイン差・本番リスク・設計コスト） |
| `food-label-pdf-gas` を独立リポとして継続、名刺版からコアのみコピー流用 | **採用** |
| Git リポジトリ名を SmartShelf にリネーム | 非採用（リモート・clone 先の更新コスト） |
| README 等に運用名とリポ名の対応を明記 | **採用** |
| 運用名表記 SmartShelef → SmartShelf | **採用**（ユーザー訂正） |

## 実装サマリ

- README・`00_プロジェクト概要.md`・`gas/README.md` に「SmartShelf（GAS/GCP）」と「food-label-pdf-gas（Git）」の対応表を追加
- `07_CHANGELOG.md` に変更行を追記

## 変更ファイル

- `README.md`
- `doc/specs/00_プロジェクト概要.md`
- `gas/README.md`
- `doc/specs/07_CHANGELOG.md`
- `.cursor/doc/memory_stream.md`
- `doc/specs/00_開発日誌.md`
- `doc/records/agent_sessions/2026-06-28_smartshelf-naming-repo-split.md`

## 検証

- `grep SmartShelef` で旧表記が残っていないことを確認
