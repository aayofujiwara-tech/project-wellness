# 介護施設「無料体験ヒアリングシート」自動データ転記システム

PDF形式の利用者情報を解析し、対応するExcelヒアリングシートへ自動転記するPythonツールです。

## セットアップ

```bash
pip install -r requirements.txt
```

## 使い方

### 1. ファイルを配置する

- `input_pdf/` フォルダにPDFファイルを入れる
- 同じディレクトリにExcelヒアリングシート（`.xlsx`）を配置する

```
project-wellness/
├── auto_pdf_to_excel.py        # メインスクリプト
├── config.json                 # 設定ファイル
├── input_pdf/
│   ├── 利用者情報_三代花子.pdf
│   └── 利用者情報_笠岡太郎.pdf
├── パシ三代さん　無料体験ヒアリングシート.xlsx
└── ルネ笠岡さん　無料体験ヒアリングシート.xlsx
```

### 2. 実行する

```bash
# 通常実行（PDFからExcelへ転記）
python auto_pdf_to_excel.py

# ドライラン（書き込みせず確認のみ）
python auto_pdf_to_excel.py --dry-run

# ログをファイルに保存
python auto_pdf_to_excel.py --log-file result.log

# ディレクトリを指定
python auto_pdf_to_excel.py --pdf-dir input_pdf --excel-dir .
```

## 照合ロジック

1. PDFからテキストを抽出し、氏名・住所等を正規表現で構造化
2. 抽出した苗字をExcelファイル名と部分一致で照合
3. Excel内のラベル（「名前」「住所」等）を検索し、隣接セルにデータを入力
4. 要介護度は該当する選択肢の横に「○」を記入

## テスト用スクリプト

```bash
# サンプルデータ生成（PDF + Excel）
python create_sample_data.py

# PDF解析のみテスト
python test_pdf_parse.py

# Excelラベル検索のみテスト
python test_excel_labels.py
```

## 設定のカスタマイズ

`config.json` でPDF抽出パターンやExcelラベルのマッピングを変更できます。

- `pdf_extraction_patterns`: PDFテキストから各項目を抽出する正規表現
- `excel_label_mapping`: Excelシート内のラベル名と入力先セルのオフセット
- `facility_abbreviations`: 施設名の略称と正式名の対応
- `filename_pattern`: Excelファイル名から氏名を抽出する正規表現
