# セッション記録: ai-agent-devenv-template v3 同期

- 日付: 2026-07-02
- スコープ: `NFR-OPS-002`（運用・保守）

## 背景・課題

- `ai-agent-devenv-template` が `a5ece0a` から `b117a2a` へ大規模更新（rules/skills 分割、`doc/ai/` 統合、`agent_workflows.mdc` 廃止）
- 本プロジェクトは 2026-05-26 のテンプレ取り込み以降、v3 構成が未反映だった

## 意思決定

| 項目 | 採用 | 非採用 |
|---|---|---|
| 知識層 | `doc/ai/`（guidelines・sessions・decisions） | `doc/ai_guidelines/` を正本のまま維持 |
| 実行層 | `.cursor/rules/` + `.cursor/skills/` をテンプレからコピー | 旧 `agent_workflows.mdc` を残す |
| プロジェクト固有規約 | `doc/ai/guidelines/notion-gcal-implementation.md` に §8–11 を分離 | テンプレの GAS 向け ADR をそのままコピー |
| 旧パス | `doc/ai_guidelines/` に stub 残置 | 即時削除 |
| セッション記録 | `doc/records/agent_sessions/` → `doc/ai/sessions/` へ git mv | 本文の改変 |

## 実装サマリ

- `.cursor/rules/`（17 ファイル）・`.cursor/skills/`（17 本）・`.cursor/doc/memory_stream.md` をテンプレから導入
- `project_identity.mdc` を Notion/GCal 向けにカスタマイズ（FR カテゴリ SYNC/OPS）
- `doc/ai/` 新設。`試験実装のエラー.md` と全 agent_sessions を移動
- `AGENTS.md`・`05_ディレクトリ構成.md`・`00_開発日誌.md`・`records/README.md`・`README.md` 更新

## 変更ファイル一覧

- `.cursor/rules/`（新規・`agent_workflows.mdc` 削除）
- `.cursor/skills/`（新規）
- `.cursor/doc/memory_stream.md`（新規）
- `doc/ai/**`（新規・移動）
- `doc/ai_guidelines/*.md`（stub 化）
- `AGENTS.md`・`README.md`
- `doc/specs/00_開発日誌.md`・`05_ディレクトリ構成.md`・`07_CHANGELOG.md`
- `doc/records/README.md`

## 検証

- `Test-Path ai-agent-devenv-template\.git` → `True`
- テンプレ pull 済み（`b117a2a`）
- ドキュメントのみの変更のため自動テストは未実行

## 学び

- テンプレ v3 は「1 巨大 workflow」から「rules + skills の組み合わせ」へ移行。プロジェクト固有の Notion/GCal 規約は汎用 rules と分離して `notion-gcal-implementation.md` に置くと、今後のテンプレ同期が楽になる
