# clasp rootDir 変更（GAS エディタフラット化）

- 日付: 2026-06-28
- スコープ: NFR-OPS-002（clasp 運用）

## 背景・課題

GAS コードエディタ上で `src/config.gs` のようにサブフォルダ付き表示となり見づらい。ローカルは `gas/src/` を維持しつつ、GAS 上は `config.gs` としてルート直下に表示したい。

## 意思決定

| 案 | 採否 |
|---|---|
| `rootDir` を `gas/src` に変更し `appsscript.json`・`ControlPanel.html` を `gas/src/` へ移動 | 採用 |
| `gas/src/` を廃止して `gas/` 直下にフラット化 | 非採用（doc 参照パス変更が広範囲） |
| `skipSubdirectories: true` | 非採用（サブフォルダ無視のみでフラット化されない） |

## 実装サマリ

- `.clasp.json` の `rootDir` を `gas` → `gas/src` に変更
- `appsscript.json`・`ControlPanel.html` を `gas/src/` へ移動
- `pnpm exec clasp push --force` で 12 ファイル同期（旧 `src/` 配下のリモートファイル削除）

## 変更ファイル一覧

- `.clasp.json`
- `gas/src/appsscript.json`（移動）
- `gas/src/ControlPanel.html`（移動）
- `doc/specs/05_ディレクトリ構成.md`
- `gas/README.md`
- `doc/specs/07_CHANGELOG.md`

## 検証

- `pnpm exec clasp push --force` — 12 ファイル同期 OK
