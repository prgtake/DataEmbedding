=====================================================
 データ埋込マクロ (Data Embedding Macro) v1.1.0
=====================================================

【概要】
Excelで「Wordの差し込み印刷」のような機能を、より柔軟かつ高速に実現するためのマクロです。
テンプレート（ひな形）シートに数式を組んでおくだけで、データシートの内容を次々と
流し込み、個別のExcelファイルとして出力します。

【主な特徴】
1. インジェクション方式: ひな形に数式（=データ!A2など）を置くだけでレイアウト自由自在。
2. 真の引き算型ロジック: ブック全体を複製して不要な部分を削る新方式を採用。
   ブックのプロパティや「内容にリンク」したメタデータ、書式設定を完全に継承します。
3. OneDrive完全対応: 同期エラーを回避するための自動プロセス制御・パス変換機能を搭載。
4. 高度なトラブルシューティング: 実行ログ出力（process_log.txt）機能により、
   万が一の際の原因特定が容易になりました。
5. 安心の元データ保護: 読み取り専用オープン＆バックアップ機能で、元データを汚しません。

【使用方法】
1. 「データ埋込マクロ.xlsm」を開きます。
2. ダイアログが表示されるので、以下の項目を指定します。
   - 対象ブック: データとひな形が入ったExcelファイルを選択。
   - データシート/ひな形シート: それぞれのシート名を選択。
   - データ開始行: 実際のデータが始まる行番号（通常は2）。
   - シート数: 1つのブックに何人分まとめるか（通常は1）。
3. 「実行」ボタンを押すと、マクロと同じフォルダの「output_results」に出力されます。

【動作環境】
- Windows 10/11
- Microsoft Excel 2016 以降 (Office 365 含む)
- OneDrive/SharePoint 環境対応

【ライセンス】
MIT License
Copyright (c) 2026 Datan (データン)

本ソフトウェアは無保証です。自己責任においてご利用ください。
詳細は HISTORY.md を参照してください。

=====================================================
 Copyright (c) 2026 Datan (データン)
=====================================================

=====================================================
 Data Embedding Macro v1.1.0
=====================================================

[Overview]
This macro provides a flexible and high-speed alternative to Word's "Mail Merge" directly within Excel. 
By setting formulas in a template sheet, it sequentially injects data from a data sheet and 
outputs them as individual Excel files.

[Key Features]
1. Injection Method: Layout flexibility by using formulas (e.g., =Data!A2) in the template.
2. True Subtraction Method: Newly adopts a workbook-replication approach. 
   Completely preserves workbook properties, linked metadata, and formatting.
3. OneDrive Ready: Automatic process control and path conversion to avoid sync errors.
4. Advanced Troubleshooting: Now includes a process logging (process_log.txt) feature 
   to easily identify root causes if an error occurs.
5. Data Protection: Operates in Read-Only mode with automatic backup/restore.

[How to Use]
1. Open "データ埋込マクロ.xlsm" and enable macros.
2. Fill in the following in the dialog:
   - Target Workbook: The Excel file containing your data and template.
   - Data/Template Sheet: Select the respective sheet names.
   - Start Row: The row number where actual data begins (usually 2).
   - Sheets per Book: Number of records to group in one file (usually 1).
3. Click "Execute". Results will be saved in the "output_results" folder.

[Environment]
- Windows 10/11
- Microsoft Excel 2016 or later (including Office 365)
- Supports OneDrive/SharePoint environments

[License]
MIT License
Copyright (c) 2026 Datan (Datan)

This software is provided "as is", without warranty of any kind. 
Use at your own risk. Refer to HISTORY.md for details.

=====================================================
 Copyright (c) 2026 Datan (Datan)
=====================================================
