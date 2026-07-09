# Memory Stream (AI Chat Log & Facts)

## 記憶ストリームの運用定義

- 本ファイルはタスク完了時に AI エージェントが自動でファクトを追記する領域です。
- 利用者による手動編集は原則不要です。
- 追記ルール: [`.cursor/rules/memory_logger.mdc`](../rules/memory_logger.mdc)

## 過去の共通原則

（50KB 超過時に memory_logger が古いログを圧縮して追記する領域）

## 蓄積されたファクト

### [2026-07-06] junior-code-comments-tier-c

- 日付: [2026-07-06]
- タスク: Tier C docstring 適用（1行ヘルパー・facade・import 系モジュール）
- エラーと解決: なし
- ユーザー指摘: なし

### [2026-07-06] junior-code-comments-tier-b

- 日付: [2026-07-06]
- タスク: Tier B ブロックコメント適用（domain/・infra/・clients 残り）
- エラーと解決: `config.py` の `resolve_project_root` インデント不整合を修正
- ユーザー指摘: なし

### [2026-07-06] junior-code-comments

- 日付: [2026-07-06]
- タスク: ジュニア向けブロックコメント規約・スキル追加、見本 `gcal_body.py`
- エラーと解決: なし
- ユーザー指摘: 【何を】ラベル形式は読みづらい・可読性の高い複数パターンで例示希望

### [2026-07-06] junior-friendly-explanations

- 日付: [2026-07-06]
- タスク: ジュニア向けチャット用語解説ルール・スキル追加
- エラーと解決: なし
- ユーザー指摘: IT用語のみの説明は理解できない・各用語に平易な解説を付ける運用希望

### [2026-07-06] src-refactor-health

- 日付: [2026-07-06]
- タスク: `NotionClient` を `clients/notion/` へ責務分割・JST/root 集約・docstring 充実
- エラーと解決: PowerShell `&&` 不可 → `;` に変更 / テスト patch 先を `_transport` に更新
- ユーザー指摘: なし

### [2026-07-03] chat-log-to-drive-template-expand

- 日付: [2026-07-03]
- タスク: `ai-agent-devenv-template` を `chat_log_to_drive` ルートへ展開（`.cursor/`・`doc/ai/`・specs 02〜07）
- エラーと解決: なし
- ユーザー指摘: テンプレを参考に展開

### [2026-07-01] template-migration

- 日付: [2026-07-01]
- タスク: `ai-agent-devenv-template` に沿い `.cursor/rules/`・`.cursor/skills/`・`doc/ai/` へ移行
- エラーと解決: なし
- ユーザー指摘: GitHub 連携後、テンプレ正本に沿ってリポジトリを更新する

### [2026-07-01] sheet-display-format

- 日付: [2026-07-01]
- タスク: FR-DB-006 シート表示用テキスト整形（勤務地全国畳み・改行3連続圧縮）
- エラーと解決: PowerShell `&&` 不可 → `;` 区切り / なし
- ユーザー指摘: 全国判定は大都市圏＋九州チェック・10県以上・畳み後 `全国（大都市圏・九州）`

### [2026-07-01] template-path-canonical

- 日付: [2026-07-01]
- タスク: テンプレ同期正本を `job_db_automation/ai-agent-devenv-template/` に固定・外 clone 参照禁止を rules 化
- エラーと解決: Glob 未検出でテンプレ無し誤判定 → `Test-Path` 手順・gitignore 再掲 / なし
- ユーザー指摘: 正本はプロジェクト内 clone のみ。`RPA_scripts/ai-agent-devenv-template` は本リポジトリから参照しない

### [2026-07-02] template-v3-pull-sync

- 日付: [2026-07-02]
- タスク: `ai-agent-devenv-template` `397d02e` を pull しプロジェクトへ反映（rules/skills/doc/ai/reference/adr）
- エラーと解決: PowerShell `&&` 不可 → `;` 区切り / なし
- ユーザー指摘: テンプレもと更新後にプロジェクトへ反映

### [2026-07-02] processed-folder-move-mode

- 日付: [2026-07-02]
- タスク: PROCESSED フォルダ移動運用へ切替（INPUT 変更・PROCESSED 追加・smokeTestConfig 拡張）
- エラーと解決: clasp run は API 実行可能デプロイ未設定 → GAS エディタ実行に手順化 / なし
- ユーザー指摘: なし

### [2026-07-02] japanese-chat-mary-rename

- 日付: [2026-07-02]
- タスク: `japanese-chat-marie` を `japanese-chat-mary` にリネームし `.cursor/skills/` とテンプレへ複製
- エラーと解決: なし
- ユーザー指摘: スキル名を mary に統一・グローバル `~/.cursor/skills/` もリネーム

### [2026-07-02] スプレッドシート出力現状確認

- 日付: [2026-07-02]
- タスク: `Job DB Automation` シート（求人DB）の出力現状を Drive API CSV エクスポートで確認
- エラーと解決: Sheets API 未有効 → Drive export で代替 / `.gsheet` ローカル直読不可
- ユーザー指摘: なし

### [2026-07-03] input-folder-clear-on-success

- 日付: [2026-07-03]
- タスク: FR-JOB-003 `PROCESSED_FOLDER_ID` 必須化・成功後 INPUT 空化・stale `.js` 削除
- エラーと解決: clasp push が Skipping → `--force` で 19 ファイル反映 / 説明欄フラグは INPUT に残るため廃止
- ユーザー指摘: 処理成功時のみ INPUT 空化（重複・日次上限は残す）

### [2026-07-03] japanese-chat-mary-tone-revision

- 日付: [2026-07-03]
- タスク: `japanese-chat-mary` 口調改訂 — ロールプレイ廃止・ですます調一貫を最優先
- エラーと解決: なし
- ユーザー指摘: お嬢様口調不要・ですます調を崩さない
