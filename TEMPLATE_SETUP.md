# AI Agent 開発環境テンプレート — セットアップガイド

このリポジトリは **AI エージェント（Cursor / Copilot）を活用した開発プロジェクト**の雛形です。
言語・フレームワークに依存しない **Agent 運用基盤**（`.cursor/` + `doc/`）が本体で、アプリコードは `src/` にスタック選定後に構築します。

---

## 含まれるファイル

```
.cursor/          … rules・skills・memory_stream
AGENTS.md
DESIGN.md         … デザイン正本（色・タイポ・コンポーネント）
README.md
config/.env.example
src/README.md     … アプリ配置先（空）
tests/README.md
scripts/README.md
doc/
├── ai/           … guidelines・sessions（財産化）・decisions
├── design/       … ai-design-brief（依頼手順）・design-refs（参考サイト）
├── reference/    … getting-started・setup（スタック別）・cheatsheets
├── specs/        … 要件ライフサイクル 00〜07
├── templates/
└── adr/README.md
```

**意図的にルートに含めないもの**: `pyproject.toml`、Python パッケージ雛形、スタック固定の setup スクリプト。

---

## A. 新規プロジェクトへの適用

### 1. テンプレートからリポジトリを作成

GitHub の「Use this template」で新規リポジトリを作成する。

### 2. プレースホルダーを置換する

| プレースホルダー | 置換内容 | 例 |
|---|---|---|
| `{{PROJECT_DESCRIPTION}}` | プロジェクト概要（1行） | `在庫管理システム` |
| `{{PROJECT_CODE}}` | リポジトリ識別子 | `inventory_app` |
| `{{APP_DIR}}` | `src/` 直下のメインディレクトリ名 | `inventory_app` |
| `{{MAIN_MODULE}}` | コアモジュール名 | `services.py` |
| `{{LAST_UPDATED}}` | 最終更新日 | `2026年7月2日` |
| `{{CATEGORY_CODE_1}}` 等 | FR カテゴリ | `AUTH` / `認証` |

### 3. 仕様 doc を埋める

- `doc/specs/02_要件定義.md` — MoSCoW・FR/NFR
- `doc/specs/04_機能一覧.md` — カテゴリ定義・機能台帳
- `doc/ai/guidelines/Project_map.md` — 参照マップ

### 4. スタックを選び `src/` を構築する

[doc/reference/setup/](doc/reference/setup/) から手順を選ぶ。

| スタック | 手順 |
|---|---|
| Python / uv | [python-uv.md](doc/reference/setup/python-uv.md) |
| Django | [django.md](doc/reference/setup/django.md) |
| GAS + clasp | [gas-clasp-pnpm.md](doc/reference/setup/gas-clasp-pnpm.md) |

構築後は `doc/specs/05_ディレクトリ構成.md` を実態に合わせて更新する。

### 4b. Web UI のデザイン（画面がある場合）

1. [doc/design/ai-design-brief.md](doc/design/ai-design-brief.md) でブリーフを埋める（Web アプリ節）
2. ルートの [DESIGN.md](DESIGN.md) に色・タイポ・コンポーネントを転記
3. Cursor で「`DESIGN.md` を正本として [画面名] を実装」と依頼

参考サイト収集は Documents ワークスペースの `playwright-cli` skill を参照（任意）。

### 5. 意思決定の記録を始める

実装・設計タスク完了ごとに:

1. `doc/ai/sessions/YYYY-MM-DD_トピック.md` に本文（[agent-session-record](.cursor/skills/agent-session-record/SKILL.md)）
2. `doc/specs/00_開発日誌.md` に索引
3. 必要なら `doc/ai/decisions/README.md` に横断 1 行

フォーマット見本: [doc/ai/sessions/_example_session.md](doc/ai/sessions/_example_session.md)

---

## B. 既存プロジェクトへの適用

`.cursor/` と `doc/` を既存リポジトリにコピーし、プレースホルダを置換する。
既存の仕様書は `doc/specs/` に転記する。

---

## Agent AI の動作確認

```
このプロジェクトについて教えて
```

`doc/ai/guidelines/Project_map.md` を基に回答が返れば成功。

---

## よくある質問

**Q. セッション記録はテンプレに含まれる？**  
A. はい。`doc/ai/sessions/` は全開発リポジトリの意思決定を集約するアーカイブです。各プロジェクトで追記し、[template_sync](.cursor/rules/template_sync.mdc) で本リポジトリへ同期します。フォーマット見本は [_example_session.md](doc/ai/sessions/_example_session.md)。

**Q. Python 用の `pyproject.toml` は？**  
A. ルートには置きません。Python 採用時は [python-uv.md](doc/reference/setup/python-uv.md) に従って作成します。

**Q. FR カテゴリはいくつ作る？**  
A. 機能の塊ごとに 1 カテゴリが目安。最初は 3〜5 から始める。
