# clasp デプロイ・GAS 同期

- 日付: 2026-06-28
- スコープ: NFR-OPS-002（clasp 運用）

## 背景・課題

Phase 1 PoC コードはリポジトリにあったが、`.clasp.json` の `scriptId` が空で GAS 実機へ未反映。利用者が clasp セットアップと push を実施した。

## 意思決定

| 案 | 採否 |
|---|---|
| `pnpm install clasp`（npm の別パッケージ） | 非採用。`@google/clasp` を devDependencies で使用 |
| OAuth 再ログイン | 非採用（`.clasprc.json` 既存・再ログインで汎用エラー） |
| 既存 GAS プロジェクトに scriptId を手動設定 | 採用 |

## 実装サマリ

- `.clasp.json` に scriptId を設定
- `pnpm clasp:push` で `gas/` を SmartShelf GAS へ同期完了
- 誤追加した `clasp@1.0.0` 依存を除去（`package.json`）

## 変更ファイル一覧

- `.clasp.json`
- `package.json`（誤依存除去）

## 検証

- `pnpm clasp:push` 成功（利用者確認）
- 次ステップ: Script Properties に `GEMINI_API_KEY`、スプレッドシートで「商品DB シートを作成」、PoC 関数の実機確認
