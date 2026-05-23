# データ埋込マクロ (Data Embedding Macro) v1.0.0

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Excelで「Wordの差し込み印刷」のような機能を、より柔軟かつ高速に実現するためのVBAマクロです。

## 🚀 はじめに
「Wordの差し込み印刷は便利だけど、Excelでやりたい！」「複雑なレイアウトや計算結果をそのまま個別のファイルにしたい」
そんな不便を解決するために開発されました。

単なるコピー＆ペーストではなく、**「インジェクション方式（数式参照）」**を採用することで、ひな形のレイアウト変更にプログラム修正なしで柔軟に対応できるのが最大の特徴です。

## ✨ 主な機能と特徴

### 1. インジェクション方式（数式参照）
ひな形シートに `=データ!A2` のような数式を組んでおくだけで、データシートの内容を次々と流し込みます。ひな形を直すだけで出力結果が変わるため、VBAの知識がなくても運用できます。

### 2. OneDrive / SharePoint 完全対応
クラウドストレージ特有の「同期エラー」や「ファイルロック」を徹底的に回避します。
- 実行中のみ **OneDrive プロセスの一時停止・再起動** を制御。
- **AutoSave（自動保存）の動的オフ**。
- `https://` 形式の URL パスをローカル物理パスへ自動変換。

### 3. 超速重複キーチェック
数千件〜数万件の巨大データでも、ID（キー）の重複を **Scripting.Dictionary (連想配列)** を使用して一瞬で検知します。

### 4. 万全のデータ保護
- 元ファイルを **読み取り専用 (ReadOnly)** で開き、メモリ上の仮想空間で処理。
- 実行直前の **自動バックアップ & リストア** 機能を搭載。

## 🛠 使用方法

1. `データ埋込マクロ.xls` を開きます（マクロを有効にしてください）。
2. 表示されるダイアログで以下の項目を設定します。
   - **対象ブック**: データとひな形が入ったExcelファイル。
   - **データシート / ひな形シート**: それぞれの名称。
   - **データ開始行**: 実際のデータが始まる行番号（通常は2）。
   - **出力ブックのシート数**: 1ファイルに何人分まとめるか。
3. 「実行」をクリックすると、マクロと同じフォルダの `output_results` フォルダに結果が出力されます。

## 📋 動作環境
- OS: Windows 10 / 11
- Excel: Microsoft Excel 2016 以降 (Microsoft 365 推奨)

## 📄 ライセンス
[MIT License](LICENSE)
Copyright (c) 2026 Datan (データン)

---
Developed by **Datan (データン)**

---

# Data Embedding Macro v1.0.0

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A VBA macro designed to achieve "Mail Merge" functionality in Excel with more flexibility and speed than standard tools.

## 🚀 Introduction
Have you ever wished for Word's "Mail Merge" capability directly in Excel to handle complex layouts and calculations? This tool was developed to solve that exact inconvenience.

By using the **"Injection Method (Formula-based Reference)"**, you can change the output layout just by editing the template sheet, without any modifications to the VBA code.

## ✨ Key Features

### 1. Injection Method
Set formulas like `=Data!A2` in your template sheet, and the macro will inject data record by record. Since the layout depends on standard Excel formulas, anyone can manage it without VBA knowledge.

### 2. Full OneDrive / SharePoint Support
Thoroughly avoids sync errors and file locks common in cloud storage.
- Controls **OneDrive process pausing/restarting** during execution.
- Dynamically disables **AutoSave**.
- Automatically converts `https://` URL paths to local physical paths.

### 3. High-Speed Duplicate Key Check
Instantly detects duplicate IDs (Keys) even in massive datasets using the **Scripting.Dictionary** object.

### 4. Robust Data Protection
- Opens target files in **Read-Only** mode, processing in a virtual memory space.
- Includes **Automatic Backup & Restore** functionality right before execution.

## 🛠 How to Use

1. Open `データ埋込マクロ.xls` (Enable macros).
2. Configure the following in the dialog:
   - **Target Workbook**: File containing your data and template.
   - **Data/Template Sheet**: Respective names of the sheets.
   - **Start Row**: Row number where data begins (usually 2).
   - **Sheets per Book**: Number of records per output file.
3. Click **Execute**. Results will be output to the `output_results` folder in the same directory.

## 📋 Environment
- OS: Windows 10 / 11
- Excel: Microsoft Excel 2016 or later (Microsoft 365 recommended)

## 📄 License
[MIT License](LICENSE)
Copyright (c) 2026 Datan (Datan)

---
Developed by **Datan (Datan)**
