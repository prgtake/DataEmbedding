# データ埋込マクロ (Data Embedding Macro) v1.1.0

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Excelで「Wordの差し込み印刷」のような機能を、より柔軟かつ高速に実現するためのVBAマクロです。

## 🚀 はじめに
「Wordの差し込み印刷は便利だけど、Excelでやりたい！」「複雑なレイアウトや計算結果をそのまま個別のファイルにしたい」
そんな不便を解決するために開発されました。

単なるコピー＆ペーストではなく、**「インジェクション方式（数式参照）」**と**「真の引き算型ロジック」**を組み合わせることで、プロパティやリンクを維持したまま、ひな形のレイアウト変更にプログラム修正なしで柔軟に対応できるのが最大の特徴です。

## ✨ 主な機能と特徴

### 1. インジェクション方式（数式参照）
ひな形シートに `=データ!A2` のような数式を組んでおくだけで、データシートの内容を次々と流し込みます。ひな形を直すだけで出力結果が変わるため、VBAの知識がなくても運用できます。

### 2. 真の引き算型ロジック (v1.1.0 新機能)
ブック全体を複製し、不要なシートを削除する方式を採用。
- **ドキュメントプロパティ（カスタムプロパティ）を完全継承**。
- 「内容にリンク」したメタデータも壊さずに出力可能。
- カラーテーマ、スタイル、名前定義なども元のブックの状態を完璧に維持します。

### 3. OneDrive / SharePoint 完全対応
クラウドストレージ特有の「同期エラー」や「ファイルロック」を徹底的に回避します。
- 実行中のみ **OneDrive プロセスの一時停止・再起動** を制御。
- `https://` 形式の URL パスをローカル物理パスへ自動変換。

### 4. 高度なトラブルシューティング
実行中に `output_results/process_log.txt` を生成。
- どのレコードで、どの処理中にエラーが起きたかを秒単位で記録します。

### 5. 万全のデータ保護
- 元ファイルを **読み取り専用 (ReadOnly)** で開き、メモリ上の仮想空間で処理。
- 実行直前の **自動バックアップ & リストア** 機能を搭載。

## 🛠 使用方法

1. `データ埋込マクロ.xlsm` を開きます（マクロを有効にしてください）。
2. 表示されるダイアログで以下の項目を設定します。
   - **対象ブック**: データとひな形が入ったExcelファイル。
   - **データシート / ひな形シート**: それぞれの名称。
   - **データ開始行**: 実際のデータが始まる行番号（通常は2）。
   - **出力ブックのシート数**: 1ファイルに何人分まとめるか。
3. 「実行」をクリックすると、マクロと同じフォルダの `output_results` フォルダに結果が出力されます。

## 📋 動作環境
- OS: Windows 10 / 11
- Excel: Microsoft Excel 2016 以降 (Microsoft 365 含む)

## 📄 ライセンス
[MIT License](LICENSE)
Copyright (c) 2026 Datan (データン)

---
Developed by **Datan (データン)**

---

# Data Embedding Macro v1.1.0

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A VBA macro designed to achieve "Mail Merge" functionality in Excel with more flexibility and speed than standard tools.

## ✨ Key Features

### 1. Injection Method
Set formulas like `=Data!A2` in your template sheet, and the macro will inject data record by record.

### 2. True Subtraction Method (New in v1.1.0)
Uses a workbook-replication approach to ensure metadata integrity.
- **Perfectly preserves Document Properties** (including custom ones).
- Keeps "Link to Content" metadata intact.
- Maintains color themes, styles, and named ranges from the source workbook.

### 3. Full OneDrive / SharePoint Support
Avoids sync errors and file locks common in cloud storage.
- Controls **OneDrive process pausing/restarting** during execution.
- Automatically converts `https://` URL paths to local physical paths.

### 4. Advanced Troubleshooting
Generates `output_results/process_log.txt` during execution.
- Records every step and provides detailed error information for easy debugging.

### 5. Robust Data Protection
- Opens target files in **Read-Only** mode.
- Includes **Automatic Backup & Restore** functionality.

## 📋 Environment
- OS: Windows 10 / 11
- Excel: Microsoft Excel 2016 or later (including Microsoft 365)

## 📄 License
[MIT License](LICENSE)
Copyright (c) 2026 Datan (Datan)

---
Developed by **Datan (Datan)**
