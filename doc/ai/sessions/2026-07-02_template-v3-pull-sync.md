# 2026-07-02 — ai-agent-devenv-template v3 をプロジェクトへ反映

- **スコープ**: `NFR-OPS-003`
- **トリガー**: 利用者がテンプレ正本を更新後、pull → 本リポジトリへ同期を依頼

## 背景・課題

`ai-agent-devenv-template` が `397d02e` まで更新されていた。
求人情報自動DB化プロジェクトの Agent 運用基盤を v3 構成（`task-value-first`・`gdrive-mcp`・三層知識・`doc/reference` 拡充）へ追従する必要があった。

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| 同期方向 | テンプレ → プロジェクト（同期対象パス） | テンプレの sessions 全文上書き |
| プロジェクト固有 | `project_identity`・業務 sessions・CHANGELOG 業務行を維持 | プレースホルダで上書き |
| GAS 前提 | `gas/**` glob・`agent_core` の `gas/` 記載を残す | テンプレ汎用版のまま |

## 実装サマリ

1. `ai-agent-devenv-template` で `git pull`（`29a4b6a` → `397d02e`）
2. `.cursor/rules/`（`project_identity` 除く）・`.cursor/skills/`・`doc/ai/` 基盤・`doc/reference/`・`doc/adr/` をコピー
3. `AGENTS.md`・`Project_map.md`・`decisions/README.md`・`00_開発日誌.md` をマージ
4. `django-ui-changes`・`Django_UIUX_ガイド`・`uv.md` を削除（テンプレに合わせる）

## 学び

テンプレ pull 後の反映は **一括コピー + 少数ファイルの手動マージ** が効率的。
`project_identity`・業務ログ・他プロジェクト sessions は除外し、`gas/` 等のスタック固有パスだけプロジェクト側で復元する。

## 変更ファイル（主要）

- `.cursor/rules/` `.cursor/skills/`（新規: `engineer_signals_core`・`gdrive_mcp_core`・`task-value-first` 等）
- `doc/ai/README.md` `runtime.md` `Project_map.md` `sessions/README.md`
- `doc/reference/setup/` `doc/adr/`
- `AGENTS.md` `07_CHANGELOG.md` `00_開発日誌.md`

## 検証

- `git status` で意図した差分のみ
- `project_identity.mdc` に求人情報自動DB化の記述が残っていること
