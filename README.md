# nagios-conf-to-xlsx

NagiosのキャッシュファイルをExcelに変換するVBAマクロツール。

## 機能

- `objects.cache` に含まれる `define service` ブロックを全件抽出してExcelシートに出力
- `status.dat` に含まれる `servicedowntime` ブロックを抽出し、対応行にダウンタイム情報を追記（任意）
- 出力列にオートフィルタ付き（A〜O列）
- `start_time` / `end_time` はOSのタイムゾーンに合わせてローカル時刻に変換して表示

## 出力項目

| 列 | 項目 | 出典 |
|----|------|------|
| A | host_name | objects.cache |
| B | service_description | objects.cache |
| C | check_period | objects.cache |
| D | check_command | objects.cache |
| E | event_handler | objects.cache |
| F | notification_period | objects.cache |
| G | check_interval | objects.cache |
| H | retry_interval | objects.cache |
| I | max_check_attempts | objects.cache |
| J | active_checks_enabled | objects.cache |
| K | passive_checks_enabled | objects.cache |
| L | event_handler_enabled | objects.cache |
| M | is_in_effect | status.dat（任意） |
| N | start_time | status.dat（任意） |
| O | end_time | status.dat（任意） |

## 使い方

### 1. 事前準備（初回のみ）

Excelの「トラストセンター」で以下を有効にする。

- **マクロを有効にする**
- **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する**

### 2. マクロの実行

1. `nagios-conf-to-xlsx.xlsm` を開く
2. 1シート目「操作」の **「マクロ実行」** ボタンをクリック
3. `objects.cache` ファイルを選択
4. `status.dat` ファイルを選択（不要な場合はキャンセル）
5. 2シート目「サービス定義」にデータが出力される

## ファイルの再生成

`nagios-conf-to-xlsx.xlsm` はPowerShellスクリプトから生成できる。

```powershell
powershell -ExecutionPolicy Bypass -File .\create_excel.ps1
```

> **前提条件**: Excel がインストールされていること。残留プロセスがある場合は先に終了させること。

## ファイル構成

```
nagios-conf-to-xlsx/
├── create_excel.ps1          # xlsmを生成するビルドスクリプト
├── nagios-conf-to-xlsx.xlsm  # Excelマクロファイル（成果物）
├── docs/
│   └── spec.md               # 仕様書
└── README.md
```

## 動作環境

- Windows
- Microsoft Excel（マクロ有効）
- PowerShell 5.x 以上（xlsm再生成時）
