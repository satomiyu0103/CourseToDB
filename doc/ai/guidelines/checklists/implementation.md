# チェックリスト — 実装・品質

最終更新: 2026-06-28

> 根拠ガイド: [`.cursor/rules/git_workflow.mdc`](../../../.cursor/rules/git_workflow.mdc) / [`.cursor/rules/testing_rules.mdc`](../../../.cursor/rules/testing_rules.mdc) / [`.cursor/rules/external_api.mdc`](../../../.cursor/rules/external_api.mdc)

---

## A. 変更前・ブランチ

- [ ] 要件定義（[02_要件定義.md](../../specs/02_要件定義.md)）と実装内容の差分を確認した  
  - 根拠: [git_workflow.mdc](../../../.cursor/rules/git_workflow.mdc)
- [ ] **モード B（本番実装）**: `master` / `main` 上で直接編集せず、作業ブランチを切ってから実装した  
  - 根拠: [git_workflow.mdc](../../../.cursor/rules/git_workflow.mdc) §ブランチ戦略
- [ ] ブランチ名に FR コードまたは変更トピックが分かる名前を付けた（例: `feat/FR-XXX-短い説明`）

## B. PR・コミット

- [ ] 1 PR = 1 責務（目安 15〜300 行）  
  - 根拠: [git_workflow.mdc](../../../.cursor/rules/git_workflow.mdc)
- [ ] コミットメッセージの件名・本文が **日本語** で意図が伝わる  
  - 根拠: [code_comments.mdc](../../../.cursor/rules/code_comments.mdc)

## C. テスト（MVP 最低ライン）

- [ ] コアロジックに unit test がある  
  - 根拠: [testing_rules.mdc](../../../.cursor/rules/testing_rules.mdc)
- [ ] 結合フローに integration test（外部 API はモック）がある  
  - 根拠: [testing_rules.mdc](../../../.cursor/rules/testing_rules.mdc)
- [ ] `tests/unit/` で外部 API を直接叩いていない  
  - 根拠: [testing_rules.mdc](../../../.cursor/rules/testing_rules.mdc)

## D. 外部 API

- [ ] タイムアウトを設定している（無制限禁止）  
  - 根拠: [external_api.mdc](../../../.cursor/rules/external_api.mdc)
- [ ] リトライは指数バックオフで上限回数がある  
  - 根拠: [external_api.mdc](../../../.cursor/rules/external_api.mdc)・[safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)

## E. ディレクトリ README（該当時）

以下の **2 つ以上** に該当したら、サブシステム直下に README を置き [Project_map.md](../Project_map.md) を更新した  

- [ ] 責務を一文で説明できない  
- [ ] ディレクトリ内ファイル数が 10 超  
- [ ] 独立した依存グラフがある  
- [ ] 入口ファイルが外部から判断できない  
- [ ] agent ルール / Project_map に「このフォルダを読め」と書きたくなる  

  - 根拠: [src_readme_policy.mdc](../../../.cursor/rules/src_readme_policy.mdc)

## F. 禁止事項の確認

- [ ] 仕様未確定のまま実装していない  
- [ ] 秘密情報をコード・ログ・テストデータに書いていない  
- [ ] 既存公開 API を事前合意なく破壊していない  

  - 根拠: [agent_core.mdc](../../../.cursor/rules/agent_core.mdc)・[git_workflow.mdc](../../../.cursor/rules/git_workflow.mdc)

---

## 実施記録

| 日付 | 実施者 | 未達サマリ |
|---|---|---|
| | | |
