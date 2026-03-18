@docs/spec.md

# プロジェクト概要

## 目的

NagiosのキャッシュファイルをExcelに変換するVBAマクロツールを提供する。
`objects.cache` からサービス定義を抽出し、`status.dat` からダウンタイム情報を付加して、
Excelシートに一覧出力する。

## 技術スタック

- **Excel VBA** : マクロ本体（`nagios-conf-to-xlsx.xlsm` に組み込み）
- **PowerShell** : xlsmファイルを生成するビルドスクリプト（`create_excel.ps1`）
- **ADODB.Stream** : VBA内でのUTF-8ファイル読み込み
- **Windows API** : タイムゾーン取得（GetLocalTime / GetSystemTime）

## ディレクトリ構成

```
nagios-conf-to-xlsx/
├── create_excel.ps1          # xlsmを生成するPowerShellスクリプト
├── nagios-conf-to-xlsx.xlsm  # 成果物（Excelマクロファイル）
├── docs/
│   └── spec.md               # 仕様書
└── sample-data/              # サンプルデータ置き場（gitignore対象）
    ├── objects.cache
    └── status.dat
```

## 開発ルール

- VBAコード内に日本語文字列を含めない（PowerShell経由でAddFromStringに渡す際に文字化けするため）
- `create_excel.ps1` を編集した後は必ずBOM付きで保存すること（PowerShell 5.xはBOMなしUTF-8をShift-JISとして読む）
- シートの参照はシート名でなくインデックス（`Sheets(2)`）を使う

## よく使うコマンド

```powershell
# xlsmファイルを再生成する
powershell -ExecutionPolicy Bypass -File .\create_excel.ps1

# 残留Excelプロセスを強制終了する（生成失敗時）
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
```
