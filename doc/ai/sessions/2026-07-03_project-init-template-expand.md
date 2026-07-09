# Agent Session: プロジェクト初期化（テンプレ展開）

日付: 2026-07-03

## 目的

`ai-agent-devenv-template` を参考に、`chat_log_to_drive` リポジトリへ Agent 運用基盤を展開する。

## 背景

- 既存: `README.md`・`gas/`・`tampermonkey/`・最小限の `doc/specs/`（00 概要・01 要件）
- 不足: `.cursor/` rules/skills、`doc/ai/`、`doc/reference/`、テンプレ形式の specs 02〜07

## 実施内容

- テンプレから `.cursor/`・`.vscode/`・`scripts/`・`src/`・`tests/`・`doc/ai/`・`doc/reference/` をルートへコピー
- `AGENTS.md`・`TEMPLATE_SETUP.md` を配置
- 旧 `01_要件定義.md` を `02_要件定義.md` へ移行（FR カテゴリ EXT/SAV/CFG/IDX）
- `project_identity.mdc`・`Project_map.md`・`05_ディレクトリ構成.md` を本プロジェクト向けに記入
- FR 採番: EXT=抽出, SAV=保存, CFG=設定, IDX=インデックス（Phase 2）

## 採用 / 非採用

| 項目 | 採用 | 非採用 |
|---|---|---|
| アプリ配置 | `gas/` + `tampermonkey/`（既存維持） | `src/` への GAS 移動 |
| 要件 doc 番号 | テンプレ準拠（02=要件） | 旧 01_要件定義 のまま |

## 次のステップ

1. Tampermonkey + GAS Phase 1 実装（FR-EXT-001, FR-SAV-001, FR-SAV-002）
2. clasp デプロイ・Web App URL 設定
3. 受け入れ基準の動作確認
