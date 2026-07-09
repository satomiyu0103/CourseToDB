---
name: agent-capabilities-index
description: skill / MCP の出典台帳（agent-capabilities.md）を再生成・更新する。skill 追加・npx skills add・MCP 変更後、または現状確認依頼時に使用。
---

# Agent Capabilities Index（能力一覧の更新）

## 目的

ai-agent-devenv-template の skill / MCP を一覧化し、出典・由来をエージェントが即参照できるようにする。

## 出力先（正本）

| 種別 | パス |
|---|---|
| manifest（手編集） | [scripts/agent-capabilities-manifest.json](../../../scripts/agent-capabilities-manifest.json) |
| 生成 Markdown | [doc/agent/agent-capabilities.md](../../../doc/agent/agent-capabilities.md) |

## 手順

1. 新規 skill の場合、manifest `entries` に `origin` / `provenance` を追記する
2. スクリプトを実行する:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/update-agent-capabilities-index.ps1
```

3. 「未登録」セクションが残っていれば manifest を追記して再実行
4. [template_sync](../../../.cursor/rules/template_sync.mdc) 対象の場合は業務プロジェクトと同期する

## 自動化

- ルール: [agent_capabilities_index](../../../.cursor/rules/agent_capabilities_index.mdc)

## 禁止

- `mcp.json` の認証情報を manifest に書く
