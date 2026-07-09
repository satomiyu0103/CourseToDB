# sessions/ — 意思決定の財産化

チャット上の実装・設計の **経緯と理由** を、プロジェクト横断の資産として残す正本です。

## 目的

| 記録層 | パス | 残すもの |
|---|---|---|
| 短期ファクト | `.cursor/doc/memory_stream.md` | 再利用キーワード（エラー対処・好み） |
| **セッション本文（本ディレクトリ）** | `doc/ai/sessions/` | 背景・比較検討・採用/非採用・検証 |
| 索引 | `doc/specs/00_開発日誌.md` | 日付ごとの 3〜10 行サマリ + 本文リンク |
| 横断索引 | `doc/ai/decisions/README.md` | トピック別の決定 1 行 + リンク |

`memory_stream` だけでは **なぜその判断に至ったか** が失われます。セッション記録は、後から参加する開発者やエージェントが文脈を復元するための財産です。

## 変更ファイル一覧と外部パス

セッション本文の「変更ファイル一覧」に `gas/`・`商品PDF_引継ぎ資料.md` 等が出てくる場合がある。これは **当時の業務リポジトリ内パス** の歴史記録である（本文の意思決定は改変しない）。

現行テンプレ内の正本は次を優先する。

| 当時のパス（例） | 現行正本 |
|---|---|
| `gas/README.md` | [doc/reference/setup/gas-operations.md](../../reference/setup/gas-operations.md) |
| `doc/reference/setup/product-schema-design.md` | 同上（昇格済み） |
| `doc/reference/migration/meishi-to-product.md` | [doc/reference/migration/meishi-to-product.md](../../reference/migration/meishi-to-product.md) |
| `doc/specs/商品PDF_引継ぎ資料.md` | [doc/templates/TPL_引継ぎ資料.md](../../templates/TPL_引継ぎ資料.md) |

横断検索は [decisions/README.md](../decisions/README.md) と本ディレクトリの grep を併用する。

## 本リポジトリ（ai-agent-devenv-template）の位置づけ

このリポジトリの `sessions/` は **単一プロジェクトのログだけではない**。

- 各開発リポジトリで記録したセッションを、[template_sync](../../../.cursor/rules/template_sync.mdc) 経由で **ここに集約** する
- PoC や過去プロジェクト由来の `2026-*.md` も、横断的な意思決定資産として **削除せず維持** する
- テンプレートを fork した新規リポジトリは、必要に応じてこのアーカイブを参照しつつ、自プロジェクト分を追記する

## いつ書くか

| 状況 | 手順 |
|---|---|
| 実装タスク完了 | [agent-session-record](../../../.cursor/skills/agent-session-record/SKILL.md) |
| 設計のみ・コード未変更 | [design-decision-record](../../../.cursor/skills/design-decision-record/SKILL.md) |
| 1 行 typo のみ | 不要（CHANGELOG のみ） |

## ファイル命名

```
YYYY-MM-DD_{短いトピック}.md
```

例: `2026-07-02_auth-provider-selection.md`

## フォーマット見本

架空の記録例: [_example_session.md](_example_session.md)  
実記録の参照例: 本ディレクトリの `2026-*.md`

## 昇格

繰り返し参照される決定は [design-decision-record](../../../.cursor/skills/design-decision-record/SKILL.md) の昇格ルールに従い、`.cursor/rules/`・`doc/specs/`・`doc/adr/` 等へ昇格し、`decisions/README.md` に索引を残します。

## テンプレ同期

各プロジェクトで `sessions/` 本文・`00_開発日誌.md` を追記したら、`ai-agent-devenv-template/` が存在すれば [template_sync](../../../.cursor/rules/template_sync.mdc) に従い **同一ターン内** に本リポジトリへコピーする。
