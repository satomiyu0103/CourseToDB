# Agent Session: プロジェクト初期化

日付: 2026-06-27

## 目的

`ai-agent-devenv-template` と `商品PDF_引継ぎ資料.md` を基に、商品 PDF → Sheets プロジェクトのフォルダ・ファイルを作成する。

## 実施内容

- テンプレートから doc / config / scripts / .cursor をプロジェクトルートへ展開
- `gas/` ディレクトリと GAS モジュール雛形を作成
- `schema.js` に商品DB 列定義（23列）を実装
- 仕様書（02〜07）を商品 PDF 向けに記入
- 引継ぎ資料を `doc/specs/` へ移動

## 次のステップ

1. `pnpm install` → clasp ログイン → `scriptId` 設定 → `clasp push`
2. 名刺 PDF 版 `meishi-pdf-to-sheets` から各 GAS モジュールを流用・商品向けに改修
3. PoC: ダミー行追記 → Gemini 1 件抽出 → 手動 1 件処理
