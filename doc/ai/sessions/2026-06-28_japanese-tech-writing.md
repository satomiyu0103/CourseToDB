# セッション記録: japanese-tech-writing（日本語技術文書規範）

- 日付: 2026-06-28
- スコープ: `NFR-OPS-002`

## 背景・課題

- k16shikano 氏の `japanese-tech-writing` Skill をリポジトリへ取り込んだ
- いつ・どう読むかの運用ルールがなく、エージェントが任意適用になりうる
- 原本確認用の出典明記と、ライセンス用 gist クローンの整理が必要だった

## 意思決定

- **採用**: `.cursor/rules/japanese_tech_writing.mdc` で適用タイミングと省略条件を定義（`doc/**` 等 glob）
- **採用**: Skill 同梱の `ATTRIBUTION.md` で原本 URL・Unlicense を記録（SKILL 本文は改変しない）
- **採用**: `documentation_wording.mdc` と併用（表記と文体を分離）
- **非採用**: Skill 本文への出典追記（エージェント読了用本文を汚さない）

## 実装サマリ

- Gist から `japanese-tech-writing/SKILL.md` を配置（ルート・テンプレ）
- `japanese_tech_writing.mdc` を新設し Phase 3・`agent_core`・`AGENTS.md`・`Project_map` を更新
- 誤クローン `67625f2a7d96e3bbdfae8d571a936063/` を削除

## 変更ファイル一覧

- `.cursor/skills/japanese-tech-writing/SKILL.md`（新規）
- `.cursor/skills/japanese-tech-writing/ATTRIBUTION.md`（新規）
- `.cursor/rules/japanese_tech_writing.mdc`（新規）
- `.cursor/rules/agent_core.mdc`
- `.cursor/skills/phase3-doc-updates/SKILL.md`
- `AGENTS.md` / `doc/ai_guidelines/Project_map.md`
- `doc/specs/07_CHANGELOG.md` / `doc/specs/00_開発日誌.md`
- `.cursor/doc/memory_stream.md`
- `ai-agent-devenv-template/` 配下の対応ファイル一式

## 検証

- `japanese_tech_writing.mdc` の glob が `doc/**` を含むこと
- `ATTRIBUTION.md` に SKILL 原本・ライセンス方針 URL が記載されていること
