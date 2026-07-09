---
name: junior-code-comments
description: src の if/for/try 等にジュニア向けの読みやすいブロックコメントを付ける手順と例文パターン。コード新規・変更時・実装規約の参照時に使用。
---

# ジュニア向けブロックコメント

## 目的

制御構文（`if` / `for` / `while` / `try`）の**直前**に、  
「なぜここか・何をするか・失敗したらどうするか」を、**読みやすい日本語**で残す。

正本ルール: [`.cursor/rules/code_comments.mdc`](../../rules/code_comments.mdc)

## いつ書くか

| 書く | 書かない |
|---|---|
| 業務ルールに関わる `if` | 自明な `import` |
| 何件を回すか分かりにくい `for` / `while` | getter・1行の単純代入 |
| 失敗時の扱いが複数ある `try/except` | コメントがコードと同文の繰り返し |

## パターン選択（可読性優先）

ラベル `【何を】` を毎回付けない。**短い自然文**か **2行以内の箇条書き**を選ぶ。

| パターン | 向いている構文 | 形 |
|:---:|---|---|
| **P1 一文リード** | 単純な `if`、短い分岐 | ブロック直前に1文 |
| **P2 見出し＋補足** | `if` / `else` の大きな分岐 | 1行目＝状況、2行目＝理由や結果（任意） |
| **P3 ループの目的** | `for` / `while` | 「各○について〜」＋終了条件を1文 |
| **P4 例外の分岐表** | `try/except` | 成功時／各 except 時を箇条書き |
| **P5 インライン** | `continue` / `break` / 小さな `if` | その行の直前に1行だけ |

---

## 例文パターン

### P1 一文リード（単純な if）

```python
# 説明文があるときだけ body に載せる
if event.description:
    body["description"] = event.description
```

```python
# Import 用設定ではステータス削除の列が無いので打ち切る
if not isinstance(self._config, AppConfig):
    return None
```

---

### P2 見出し＋補足（大きな分岐）

```python
# 終日予定は GCal の「日付のみ」形式で送る（時刻ありとは別ルート）
if event.is_all_day:
    # end は翌日にする（GCal が終了日を排他的に扱うため）
    ...
else:
    # 時刻ありは UTC の dateTime で送る
    ...
```

```python
# 同期対象から外れた ID を削除候補に足す
for source_id in tracked - current_ids:
    # すでに候補に入っていれば何もしない
    if source_id in by_id:
        continue
    ...
```

---

### P3 ループの目的（for / while）

```python
# Notion の全ページを1件ずつ Event に変換し、同期リストか削除候補に振り分ける
for page in self._transport.iter_pages(sync_filter=True):
    ...
```

```python
# GCal に次ページがある限りイベントを取り続ける（pageToken が空になったら終了）
while True:
    response = self._list_events_page(...)
    ...
    if not page_token:
        break
```

---

### P4 例外の分岐表（try / except）

```python
# Notion ページを1件取得して削除理由を調べる
# ・成功 → 下の if で archived / ステータス / sync_off を判定
# ・404 → ページ削除済みとして GCal 削除対象にする
# ・その他 → 同期全体を止める（raise）
try:
    page = self._transport.retrieve_page(event.source_id)
except APIResponseError as error:
    if error.status == 404:
        return DeleteMark(event=event, reason="not_found")
    raise
```

```python
# 1ページ分を Event に変換する
# ・成功 → 同期リスト or 削除候補へ
# ・NotionPageParseError → ログしてスキップ（他のページは続行）
try:
    event = page_to_event(page, ...)
except NotionPageParseError as error:
    parse_skipped += 1
    self._logger.warning(...)
    continue
```

---

### P5 インライン（continue / break / 小さな補正）

```python
    if end_date <= start_date:
        # 終了が開始以前なら、開始の翌日にそろえる
        end_date = _utc_date_str(event.start_at + timedelta(days=1))
```

```python
    if has_more and not cursor:
        # 次ページがあるのに cursor が無い → 無限ループ防止で打ち切る
        break
```

---

## 悪い例（避ける）

```python
# 【何を】説明文を設定する
# 【いつ】description があるとき
# 【どうする】body に入れる
if event.description:
```

→ ラベルがノイズ。P1 の1文で足りる。

```python
# ループする
for page in pages:
```

→ 何のループか書いていない。P3 を使う。

```python
try:
    ...
except Exception:
    pass
```

→ 例外を握りつぶすだけのコメントは不可。P4 で各分岐の意図を書くか、握りつぶさない。

---

## 関数 docstring（任意で足す行）

ジュニア向けに、入出力のイメージを1〜2行足してよい。

```python
def event_to_gcal_body(event: Event) -> dict[str, Any]:
    """Event を GCal の insert/update 用 dict に変換する。

    入力: タイトル・開始終了・終日フラグを持つ Event
    出力: GCal API にそのまま渡せる body（HTTP は呼び出し側が担当）
    """
```

---

## 適用ティア（目安）

| ティア | 対象 | 義務 |
|:---:|---|---|
| A | `sync/`, `clients/notion/`, `cli.py` | 非自明な制御構文すべて |
| B | `domain/`, `infra/`, 他 `clients/` | 業務判断・例外・ループのみ |
| C | 1行ヘルパー | docstring のみ |

既存ファイルは**触った関数・ブロックだけ**足す（一括全ファイルはしない）。

## パイロット・Tier A 適用済み

| パス | 状態 |
|---|---|
| `clients/gcal_body.py` | パイロット見本 |
| `clients/notion/*.py` | Tier A 適用済み |
| `sync/engine.py` | Tier A 適用済み |
| `cli.py` | Tier A 適用済み |

## Tier B 適用済み

| パス | 状態 |
|---|---|
| `domain/event_change.py` | Tier B 適用済み |
| `domain/notion_datetime.py` | Tier B 適用済み |
| `domain/delete_guard.py` | Tier B 適用済み |
| `infra/config.py` | Tier B 適用済み |
| `infra/state_store.py` | Tier B 適用済み |
| `infra/rate_limit.py` | Tier B 適用済み |
| `infra/notifier.py` | Tier B 適用済み |
| `infra/event_json.py` | Tier B 適用済み |
| `infra/logger.py` | Tier B 適用済み |
| `clients/notion_mapping.py` | Tier B 適用済み |
| `clients/gcal_client.py` | Tier B 適用済み |
| `clients/gcal_mapping.py` | Tier B 適用済み |
| `clients/gcal_apply.py` | docstring のみ（Tier C） |
| `clients/notion_client.py` | docstring のみ（facade・Tier C） |

## Tier C 適用済み

| パス | 内容 |
|---|---|
| `clients/notion_client.py` | facade 各メソッドに1行 docstring |
| `clients/gcal_apply.py` | モジュール・型 docstring（変更なし） |
| `clients/gcal_body.py` | `_utc_date_str` docstring |
| `clients/notion_mapping.py` | `_plain_text` / `_get_property` docstring |
| `clients/gcal_mapping.py` | `is_gcal_cancelled` / `_parse_gcal_*` docstring |
| `domain/notion_datetime.py` | `is_timed_notion_date` docstring |
| `infra/config.py` | `_env_*` / `_parse_status_*` / `_notion_property_names` docstring |
| `infra/rate_limit.py` | `_retry_error_code` docstring |
| `infra/logger.py` | モジュール docstring / `_launcher_quiet_console` docstring |
| `infra/state_store.py` | `_load_last_events` docstring |
| `infra/notifier.py` | `_parse_failure_line` docstring |
| `csv_import/` | 各モジュール docstring / `_year_month_from_name` docstring |
| `gcal_import/` | 各モジュール docstring |
| `timezone_repair/runner.py` | モジュール / `TimezoneRepairResult` / `_build_date_patch` docstring |
| `main.py` | 既存 docstring（変更なし） |
| `domain/models.py` 等 | dataclass docstring 既存（変更なし） |

## 関連

- チャット用語解説: [junior-friendly-explanations](../junior-friendly-explanations/SKILL.md)
- D形式の別紙解説: [d-format-code-guide](../d-format-code-guide/SKILL.md)
