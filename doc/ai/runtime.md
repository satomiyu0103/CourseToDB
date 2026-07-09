# 実行層（.cursor）と知識層（doc/ai）の分担

## なぜ二層か

| 層 | パス | 役割 |
|---|---|---|
| **実行層** | `.cursor/` | Cursor が自動読み込みする制約・手順・短期記憶 |
| **知識層** | `doc/ai/` | 検索・参照・監査向けの構造化知識 |
| **仕様層** | `doc/specs/` | プロダクトの要件・設計正本 |

`.cursor/rules/` と `.cursor/skills/` は Cursor の仕様上、`.cursor/` 配下に置く必要がある。長文ログや横断索引は `doc/ai/` に置き、利用者とエージェントの双方が `doc/` から辿れるようにする。

## 実行層の索引

| パス | 内容 |
|---|---|
| [`.cursor/rules/`](../../.cursor/rules/) | 常時または glob 適用の制約 |
| [`.cursor/skills/`](../../.cursor/skills/) | タスク手順（Phase 1〜3・監査等） |
| `~/.agents/skills/` / `.agents/skills/` | 外部スキル（[skills CLI](../../reference/setup/skills-cli.md) で `npx skills add` したもの） |
| [`.cursor/doc/memory_stream.md`](../../.cursor/doc/memory_stream.md) | チャット由来の短期ファクト（タスク完了時に自動追記） |

## 知識層との関係

| 書く内容 | 置き場 |
|---|---|
| 詳細な比較検討・採用/非採用 | `doc/ai/sessions/` |
| 時系列索引 | `doc/specs/00_開発日誌.md` |
| トピック横断で探す | `doc/ai/decisions/README.md` |
| キーワード・好み・禁止（短い） | `.cursor/doc/memory_stream.md` |
| 毎タスク守る制約 | `.cursor/rules/` へ昇格 |

詳細は sessions、再利用キーワードは memory_stream。50KB 超過時は memory_stream の古いブロックを `過去の共通原則` に圧縮する（[memory_logger.mdc](../../.cursor/rules/memory_logger.mdc)）。
