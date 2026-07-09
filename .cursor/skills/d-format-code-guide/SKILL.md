---
name: d-format-code-guide
description: やりたいことの整理や既存コードの理解依頼時に、D形式（日本語変数＋Python制御構文＋行コメント）の解説ファイルを doc/reference/getting-started/d-format/ に作成する。「教えて」「理解したい」「わかりやすく」「やりたいこと」「作りたい」「D形式」等の依頼で作動。
---

# D形式コード解説

## 目的

利用者が **やりたいこと** を伝えるとき、または **既存コードを理解したい** と依頼したとき、**D形式** の解説ファイルを `doc/reference/getting-started/d-format/` に作成する。

D形式は学習用の標準形であり、`src/` `tests/` の識別子は [naming_conventions.mdc](../../rules/naming_conventions.mdc) の平易な英語のまま変更しない。

手順: [output-template.md](output-template.md)  
索引: [doc/reference/getting-started/d-format/README.md](../../../doc/reference/getting-started/d-format/README.md)

## いつ作動するか

| 種別 | トリガー文言の例 |
|---|---|
| やりたいこと（Intent） | 「やりたいこと」「こうしたい」「作りたい」「新規プロジェクトで」 |
| 理解（Understand） | 「教えて」「理解したい」「わかりやすく」「読み解いて」「D形式」「Dで解説」 |

明示的に「D形式で」と言わなくても、理解系の依頼では本 SKILL を優先する。

**作動しない**: 実装タスクのみ（コード変更が主目的で解説ファイルの作成を求めていない）、1 行 typo、利用者が「チャットだけで」と指示した場合。

## D形式の定義

| 要素 | ルール |
|---|---|
| 変数・処理名 | 日本語（例: `数字`, `十倍した値`） |
| 制御構文 | Python のまま（`for`, `if`, 字下げ） |
| 英語 API | `range`, `print` 等。用語セクションで初出定義 |
| 行コメント | 各行の意図を日本語で補足 |
| 実行 | 学習用としてそのまま動かせる Python |

### やること

- 範囲は `0 から 9 まで` のように境界を明示する（`0～10の間` は使わない）
- 演算子は半角（`*`, `+`）
- 複雑な処理は関数単位で D形式ブロックを分ける
- 本文は [japanese-tech-writing](../japanese-tech-writing/SKILL.md) の一文一行・段落一トピックに従う

### やらないこと

- `src/` `tests/` の識別子を日本語に変更しない
- 自然文だけ（A形式のみ）で終わらない。必ず D形式コードブロックを含める
- Z形式（英語のみ・コメントなし）を標準形にしない

## 2モード

### Intent モード（やりたいことを伝える）

1. 利用者の自然語要望を読み、スコープを 1 トピックに絞る（広すぎる場合は分割を提案）
2. `doc/specs/02_要件定義.md` があれば整合確認（無ければスキップ可）
3. D形式コードを **要望から先に** 書く（実装が無くても可）
4. `doc/reference/getting-started/d-format/{kebab-case-slug}.md` を新規作成
5. [d-format/README.md](../../../doc/reference/getting-started/d-format/README.md) の一覧に 1 行追記
6. チャットでファイルパスを報告する。実装へ進む場合は [implementation-phase1](../implementation-phase1/SKILL.md) を案内

### Understand モード（コードを理解したい）

1. 対象を特定する（ファイルパス・関数・チャットで示された範囲）
2. 実コードを読み、処理の塊ごとに D形式へ **意訳** する（コピペではなく読みやすさ優先）
3. 「対応する実装」表で日本語名と英語識別子を対応づける（必須）
4. 既存の `d-format/{slug}.md` があれば同ファイルに節追加、なければ新規作成
5. [コード解説.md](../../../doc/reference/getting-started/コード解説.md) の該当ファイル節があれば相互リンクを張る

## 出力先・命名

| 項目 | 内容 |
|---|---|
| ディレクトリ | `doc/reference/getting-started/d-format/` |
| ファイル名 | `{kebab-case-slug}.md`（例: `loop-multiply-by-ten.md`, `csv-sum-total.md`） |
| 索引 | 同ディレクトリの `README.md` を必ず更新 |
| 重複 | 同一トピックなら既存ファイルに節追加。別トピックなら新規ファイル |

## 出力テンプレート

[output-template.md](output-template.md) の見出し構造に従う。省略してよいのは「発展」のみ。

## 他ドキュメントとの関係

| ドキュメント | 役割 |
|---|---|
| [コード解説.md](../../../doc/reference/getting-started/コード解説.md) | ファイル単位の責務・設計意図 |
| 本 SKILL の出力 | 処理の流れを追う学習層（D形式） |
| [naming_conventions.mdc](../../rules/naming_conventions.mdc) | `src/` の実装識別子（平易な英語） |
| [agent-session-record](../agent-session-record/SKILL.md) | 実装完了の経緯記録。本 SKILL とは別 |

## 完了チェック

- [ ] `d-format/{slug}.md` を [output-template.md](output-template.md) 構造で作成した
- [ ] D形式コードブロックが含まれ、実行可能な Python である
- [ ] `d-format/README.md` の一覧を更新した
- [ ] Understand モードなら「対応する実装」表を記載した
- [ ] `src/` の識別子を変更していない
- [ ] `.cursor/skills/` を変更した場合、完了報告前に [template_sync.mdc](../../rules/template_sync.mdc) を確認した
