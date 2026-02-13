#!/usr/bin/env python3
"""ステップ2: Excelファイルの固定セル座標マッピングが正しいか確認する。

使い方:
    python test_excel_labels.py [Excelファイルパス]
    python test_excel_labels.py --form form1
    (引数省略時はカレントディレクトリの全ヒアリングシートを処理)
"""

import argparse
import json
import os
import re
import sys

from openpyxl import load_workbook


def load_config(config_path="config.json"):
    """config.json を読み込む。"""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def extract_name_from_filename(filename, config):
    """ファイル名から氏名（苗字）部分を抽出する。"""
    pattern = config.get("filename_pattern", "")
    match = re.search(pattern, filename)
    if match:
        raw = match.group(1)
        abbrevs = config.get("facility_abbreviations", {})
        for abbr in abbrevs:
            if raw.startswith(abbr):
                return raw[len(abbr):]
        return raw
    return None


def analyze_excel(filepath, config, form_name="form1"):
    """Excelファイルを解析し、固定座標マッピングの各セルの現在値を表示する。"""
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    print(f"\n  シート名: {ws.title}")
    print(f"  使用範囲: {ws.dimensions}")
    print(f"  フォーム: {form_name}")

    cell_mapping = config["fixed_cell_mapping"][form_name]

    print(f"\n  {'フィールド':<14} {'セル座標':<10} {'現在値':<30} {'状態'}")
    print(f"  {'-'*14} {'-'*10} {'-'*30} {'-'*6}")

    for field_name, field_info in cell_mapping.items():
        cell_addr = field_info["cell"]
        current_val = ws[cell_addr].value
        current_val_str = str(current_val) if current_val else "(空)"
        fmt = field_info.get("format", "-")
        status = "値あり" if current_val else "空"
        print(f"  {field_name:<14} {cell_addr:<10} {current_val_str:<30} {status}  [format: {fmt}]")

    wb.close()


def main():
    parser = argparse.ArgumentParser(description="Excel固定セル座標マッピングテスト")
    parser.add_argument("files", nargs="*", help="対象Excelファイル")
    parser.add_argument("--form", default="form1", help="フォーム名 (デフォルト: form1)")
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    config = load_config(os.path.join(base_dir, "config.json"))

    if args.form not in config.get("fixed_cell_mapping", {}):
        print(f"エラー: フォーム「{args.form}」が config.json に存在しません。")
        sys.exit(1)

    # 対象ファイルの決定
    if args.files:
        excel_files = args.files
    else:
        excel_dir = os.path.join(base_dir, "excel_sheets")
        if not os.path.isdir(excel_dir):
            print(f"エラー: ディレクトリが見つかりません: {excel_dir}")
            sys.exit(1)
        excel_files = [
            os.path.join(excel_dir, f)
            for f in os.listdir(excel_dir)
            if f.endswith(".xlsx") and "ヒアリングシート" in f
        ]

    if not excel_files:
        print("対象のExcelファイルが見つかりません。")
        sys.exit(1)

    print("=" * 70)
    print("Excel固定セル座標マッピングテスト")
    print("=" * 70)

    for filepath in sorted(excel_files):
        filename = os.path.basename(filepath)
        name = extract_name_from_filename(filename, config)
        print(f"\n--- {filename} ---")
        print(f"  抽出氏名: {name}")
        analyze_excel(filepath, config, form_name=args.form)

    print("\n" + "=" * 70)
    print("テスト完了")
    print("=" * 70)


if __name__ == "__main__":
    main()
