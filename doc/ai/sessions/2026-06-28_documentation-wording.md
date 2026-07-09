# セッション記録: documentation_wording（利用者表記）

- 日付: 2026-06-28
- スコープ: `NFR-OPS-002`

## 背景・課題

- ドキュメント・ルール内で開発者・利用者を「人間」「人向け」と表記していた
- 今後の新規・改訂でも同表記を避けるガイドが必要

## 意思決定

- **採用**: `.cursor/rules/documentation_wording.mdc`（alwaysApply）で置き換え表と使い分けを定義
- **採用**: Phase 3 Skill に doc 表記準拠行を追加
- **非採用**: 「他人の」等、第三者を指す一般用法の一括置換

## 実装サマリ

- 「人間向け」「人向け」→「利用者向け」、「人が」→「利用者が／開発者が」等に置換
- 親リポと `ai-agent-devenv-template/` を同期（`documentation_wording.mdc` 含む）

## 変更ファイル一覧

- `.cursor/rules/documentation_wording.mdc`（新規）
- `.cursor/rules/agent_core.mdc` / `project_identity.mdc` / `safe_operations_core.mdc`
- `.cursor/skills/phase3-doc-updates/SKILL.md`
- `doc/ai_guidelines/Project_map.md` / `checklists/security.md` / `試験実装のエラー.md`
- `doc/reference/README.md` / `doc/specs/05_ディレクトリ構成.md` / `README.md`
- `ai-agent-devenv-template/` 配下の対応ファイル一式

## 検証

- `人間` `人向け` の残存は `documentation_wording.mdc` 内の「避ける語」定義のみ
