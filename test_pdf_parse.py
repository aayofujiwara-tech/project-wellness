#!/usr/bin/env python3
"""ステップ1: サンプルPDFを読み取り、構造化データに変換できるかテストする。

使い方:
    python test_pdf_parse.py [PDFファイルパス]
    (引数省略時は input_pdf/ 内の全PDFを処理)
"""

import json
import os
import re
import sys

import fitz  # pymupdf


def load_config(config_path="config.json"):
    """config.json を読み込む。"""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def extract_text_from_pdf(pdf_path):
    """PDFからテキストを抽出する（pymupdf使用）。"""
    all_text = []
    doc = fitz.open(pdf_path)
    for page in doc:
        text = page.get_text()
        if text:
            all_text.append(text)
    doc.close()
    return "\n".join(all_text)


def parse_pdf_text(text, config):
    """抽出テキストから正規表現でフィールドを抽出し、辞書に構造化する。"""
    patterns = config["pdf_extraction_patterns"]
    result = {}

    for field_name, regex_list in patterns.items():
        for pattern in regex_list:
            match = re.search(pattern, text)
            if match:
                value = match.group(1).strip()
                result[field_name] = value
                break

    return result


def extract_data_from_pdf(pdf_path, config):
    """PDFファイル1つから構造化データを抽出する。"""
    text = extract_text_from_pdf(pdf_path)
    data = parse_pdf_text(text, config)
    return text, data


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config = load_config(os.path.join(base_dir, "config.json"))

    # 対象PDFの決定
    if len(sys.argv) > 1:
        pdf_files = [sys.argv[1]]
    else:
        pdf_dir = os.path.join(base_dir, "input_pdf")
        if not os.path.isdir(pdf_dir):
            print(f"エラー: ディレクトリが見つかりません: {pdf_dir}")
            sys.exit(1)
        pdf_files = [
            os.path.join(pdf_dir, f)
            for f in os.listdir(pdf_dir)
            if f.lower().endswith(".pdf")
        ]

    if not pdf_files:
        print("処理対象のPDFファイルが見つかりません。")
        sys.exit(1)

    print("=" * 60)
    print("PDF解析テスト")
    print("=" * 60)

    for pdf_path in sorted(pdf_files):
        print(f"\n--- {os.path.basename(pdf_path)} ---")

        raw_text, data = extract_data_from_pdf(pdf_path, config)

        print(f"\n[抽出テキスト]")
        print(raw_text)

        print(f"\n[構造化データ]")
        if data:
            for key, value in data.items():
                print(f"  {key}: {value}")
        else:
            print("  (データを抽出できませんでした)")

        # 必須フィールドのチェック
        required_fields = ["利用者氏名", "住所", "要介護度"]
        missing = [f for f in required_fields if f not in data]
        if missing:
            print(f"\n  [警告] 不足フィールド: {', '.join(missing)}")
        else:
            print(f"\n  [OK] 必須フィールドすべて抽出成功")

    print("\n" + "=" * 60)
    print("テスト完了")
    print("=" * 60)


if __name__ == "__main__":
    main()
