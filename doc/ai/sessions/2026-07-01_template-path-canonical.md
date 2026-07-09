# テンプレ同期正本パスの固定

日付: 2026-07-01  
スコープ: `NFR-OPS-003`

## 背景・課題

`ai-agent-devenv-template` はプロジェクト内に clone 済みだったが、`.gitignore` によりエージェントの Glob が未検出し、テンプレ同期が黙ってスキップされていた。利用者は正本を `job_db_automation/ai-agent-devenv-template/` とし、親ディレクトリの別 clone は本リポジトリから参照しない方針を明示した。

## 意思決定

| 採用 | 非採用 |
|------|--------|
| 同期正本 = ワークスペースルート `ai-agent-devenv-template/` のみ | 親 `RPA_scripts/ai-agent-devenv-template` を本 repo から参照 |
| `Test-Path ai-agent-devenv-template\.git` で存在確認 | Glob 0 件でスキップ |
| `workspace_boundary` に外 clone 禁止を明記 | `.gitignore` からテンプレ行を削除（誤 add リスク） |

## 変更ファイル

- `.cursor/rules/template_sync.mdc` / `workspace_boundary.mdc` / `memory_logger.mdc`
- `.cursor/rules/project_identity.mdc`（本プロジェクト固有表）
- `AGENTS.md` / `.gitignore` / `doc/specs/07_CHANGELOG.md`
- 上記運用基盤を `ai-agent-devenv-template/` へ反映

## 検証

```powershell
Test-Path ai-agent-devenv-template\.git
```

`True` を確認。
