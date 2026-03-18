# 仕様書

## 概要

Nagiosのキャッシュファイルを読み込み、Excelファイルに出力するExcelマクロ（VBA）ツール。

---

## 機能

### 1. サービス定義の出力（objects.cache）

`objects.cache` に含まれる `define service` ブロックを全て抽出し、以下の項目をExcelの2シート目（「サービス定義」シート）に出力する。

| 列 | 項目 | 説明 |
|----|------|------|
| A | host_name | ホスト名 |
| B | service_description | サービス説明 |
| C | check_period | チェック期間 |
| D | check_command | チェックコマンド |
| E | event_handler | イベントハンドラ |
| F | notification_period | 通知期間 |
| G | check_interval | チェック間隔 |
| H | retry_interval | リトライ間隔 |
| I | max_check_attempts | 最大チェック試行回数 |
| J | active_checks_enabled | アクティブチェック有効フラグ |
| K | passive_checks_enabled | パッシブチェック有効フラグ |
| L | event_handler_enabled | イベントハンドラ有効フラグ |

- 1行目にヘッダーを出力し、オートフィルタを設定する（A〜O列）

### 2. ダウンタイム情報の追記（status.dat）※任意

`status.dat` に含まれる `servicedowntime` ブロックを全て抽出し、
`host_name` と `service_description` が一致する行に以下の項目を追記する。

| 列 | 項目 | 説明 |
|----|------|------|
| M | is_in_effect | ダウンタイム有効フラグ |
| N | start_time | 開始時刻（ローカル時刻、`yyyy/mm/dd hh:mm:ss` 形式） |
| O | end_time | 終了時刻（ローカル時刻、`yyyy/mm/dd hh:mm:ss` 形式） |

- `start_time` / `end_time` はUnixタイムスタンプ（秒）をOSのタイムゾーンに基づいてローカル時刻に変換して表示する

---

## Excelファイル構成

| シート | 名前 | 内容 |
|--------|------|------|
| 1シート目 | 操作 | ツールの機能説明・利用方法・マクロ実行ボタン |
| 2シート目 | サービス定義 | マクロ実行結果のデータ出力先 |

### マクロ実行ボタン

- 「操作」シートに配置された角丸矩形のプッシュボタン（300×80px）
- クリックするとマクロが起動し、ファイル選択ダイアログが表示される

---

## 操作手順

1. Excelファイルを開く
2. 1シート目「操作」の「マクロ実行」ボタンをクリック
3. `objects.cache` ファイルを選択する
   - ダイアログの初期ディレクトリはExcelファイルと同じフォルダ
4. `status.dat` ファイルの選択ダイアログが表示される（キャンセルで省略可）
5. マクロが実行され、2シート目「サービス定義」にデータが生成される

---

## 備考

- `status.dat` の選択はデフォルトで表示されるが、キャンセルすることで省略可能
- `status.dat` を省略した場合、ダウンタイム情報（M〜O列）は出力されない
- マクロを再実行すると、2シート目のデータは上書きされる
- ファイルはUTF-8エンコーディングに対応（日本語サービス名を含む場合も正常に読み込まれる）
