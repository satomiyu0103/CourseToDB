# gdrive-mcp チェックリスト

正本: [SKILL.md](SKILL.md) / 制約: [gdrive_mcp_core.mdc](../../rules/gdrive_mcp_core.mdc)

## 着手前

- [ ] MCP サーバーが Cursor に接続されている
- [ ] `GDRIVE_ENABLE_UPLOAD` が未設定または `false`（readonly）
- [ ] 検索キーワード・対象フォルダを利用者と合意した
- [ ] `gdrive_mcp_core` の禁止事項を確認した

## 利用中

- [ ] `search` → `read` / `download` の順で取得した
- [ ] 大きいファイルは `download` を使った
- [ ] ツール実行前に内容・範囲を確認した（間接プロンプトインジェクション対策）
- [ ] チャットに個人情報を不必要に全文貼っていない

## 完了時

- [ ] リポジトリの diff に個人情報本文が含まれていない
- [ ] `memory_stream` / sessions の追記が索引のみである
- [ ] OAuth 鍵・トークン・`mcp.json` を git に add していない
