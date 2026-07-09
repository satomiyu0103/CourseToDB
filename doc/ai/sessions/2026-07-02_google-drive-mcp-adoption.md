# セッション記録: Google Drive MCP 運用基盤

- 日付: 2026-07-02
- スコープ: NFR-OPS-002（エージェント運用基盤）

## 背景・課題

- 住所・連絡先などの個人情報を `.cursor/doc` やリポジトリに置くと git 混入リスクがある
- Google Drive に正本を置き、エージェントが MCP 経由で参照する運用が必要だった
- 連携の手軽さと機能を重視し、初期は readonly とする方針が決まった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| MCP サーバー | [dylancaponi/gdrive-mcp-server](https://github.com/dylancaponi/gdrive-mcp-server) | 公式 `@modelcontextprotocol/server-gdrive`（deprecated） |
| スコープ | `drive.readonly`（既定） | 初回から `GDRIVE_ENABLE_UPLOAD` |
| 公式リモート MCP | 保留（機能拡充時に再検討） | 現時点での採用 |
| 正本の置き場所 | Google Drive | リポジトリ `.cursor/doc` |
| MCP 設定 | `~/.cursor/mcp.json`（グローバル） | プロジェクト `.cursor/mcp.json` |
| 文書構成 | Rule + Skill + guideline + setup（task-value-first 同型） | setup のみ |

## 実装サマリ

- `.cursor/rules/gdrive_mcp_core.mdc` — 個人情報境界・readonly 既定
- `.cursor/skills/gdrive-mcp/SKILL.md` + `checklist.md` — 手順正本・昇格・移行トリガ
- `doc/ai/guidelines/google-drive-mcp.md` — 背景・採用理由・データフロー
- `doc/reference/setup/google-drive-mcp.md` — Windows 向けセットアップ手順
- `AGENTS.md`・`Project_map.md`・`reference/README.md`・`decisions/README.md`・`skills-cli.md` を更新
- `memory_logger.mdc`・`agent_core.mdc` に Drive 個人情報の追記禁止を接続

## 学び

- 外部正本（Drive）とリポジトリ（索引のみ）を Rule で明示すると、MCP 利用時の PII 漏洩を防ぎやすい
- MCP 設定はグローバル `mcp.json` に置く方針を setup に書くと、テンプレ clone 先への秘密混入を避けられる
- 書き込み昇格・公式リモート移行のトリガーを Skill に置いておくと、将来の方針転換が迷わない

## 変更ファイル（主要）

- `.cursor/rules/gdrive_mcp_core.mdc`
- `.cursor/skills/gdrive-mcp/SKILL.md`
- `.cursor/skills/gdrive-mcp/checklist.md`
- `doc/ai/guidelines/google-drive-mcp.md`
- `doc/reference/setup/google-drive-mcp.md`
- `AGENTS.md`
- `doc/ai/guidelines/Project_map.md`
- `doc/reference/README.md`
- `doc/ai/decisions/README.md`
- `doc/reference/setup/skills-cli.md`
- `.cursor/rules/memory_logger.mdc`
- `.cursor/rules/agent_core.mdc`
- `doc/specs/07_CHANGELOG.md`

## 検証

- 新規 doc の相互リンクが参照先パスと一致することを目視確認
- 利用者マシンでの OAuth / `mcp.json` 設定は手順書のみ（本セッションでは未実施）
