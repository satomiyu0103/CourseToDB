# チェックリスト — 運用・バッチ

最終更新: 2026-06-28

> 根拠ガイド: [`.cursor/rules/safe_operations_core.mdc`](../../../.cursor/rules/safe_operations_core.mdc) / [`.cursor/skills/safe-operations-detail/SKILL.md`](../../../.cursor/skills/safe-operations-detail/SKILL.md)  
> 本プロジェクト固有のチャネル・モジュール名は [03_システム設計.md](../../specs/03_システム設計.md)・`src/{{APP_DIR}}/README.md`（適用後）。

---

## A. 実行ログ

- [ ] バッチ開始・終了に **run_id** と件数サマリ（processed / skipped / failed）がある  
  - 根拠: [safe_operations_core.mdc](../../../.cursor/rules/safe_operations_core.mdc)
- [ ] 各処理単位で **stage** + 対象 ID がログに出る  
  - 根拠: [safe_operations_core.mdc](../../../.cursor/rules/safe_operations_core.mdc)
- [ ] 失敗時の ERROR に **段階名** が含まれる（どこで落ちたか一目で分かる）  
  - ダメな例: `Error: undefined` のみ
- [ ] 秘密情報・個人情報全文・巨大ペイロード全文がログに出ない  
  - 根拠: [safe_operations_core.mdc](../../../.cursor/rules/safe_operations_core.mdc)

## B. 例外処理と通知

- [ ] 予期可能な失敗（重複・枠不足等）と致命的失敗の応答がコード上で分かれている  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)
- [ ] 障害通知チャネルが **1 つ** に設計上決まっている（未決なら ROADMAP に明記）  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)・[03_システム設計.md](../../specs/03_システム設計.md)
- [ ] 通知本文に API キー・トークン・個人情報全文が含まれない  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)
- [ ] 連続失敗時の重複抑制がある、または「未実装」と [06_ROADMAP.md](../../specs/06_ROADMAP.md) に明記  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)

## C. 処理量・上限・再実行

- [ ] 外部 API 呼び出し前に残枠 or 上限チェックがある（無制限は specs に明示）  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)
- [ ] 上限超過時の挙動（スキップ vs 中止）が要件・設計に書いてある  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)・[02_要件定義.md](../../specs/02_要件定義.md)
- [ ] バッチ終了時に件数サマリと枠消費（または所要時間）がログ or 台帳に残る  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)
- [ ] 同じ入力の再実行で二重登録・二重課金にならない（自然キー・処理済みフラグ）  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)・[試験実装のエラー.md](../試験実装のエラー.md)

## D. 新規バッチ追加時（最小セット）

- [ ] run_id + 段階ログ、上限方針、通知、FR/NFR を [04_機能一覧.md](../../specs/04_機能一覧.md) に追記済み  
  - 根拠: [safe-operations-detail/SKILL.md](../../../.cursor/skills/safe-operations-detail/SKILL.md)

---

## 実施記録

| 日付 | 実施者 | 未達サマリ |
|---|---|---|
| | | |
