#!/usr/bin/env python3
"""ステップ2: Excelファイル内のラベル検索が正しく機能するか確認する。

使い方:
    python test_excel_labels.py [Excelファイルパス]
    (引数省略時はカレントディレクトリの全ヒアリングシートを処理)
"""

import json
import os
import re
import sys

from openpyxl import load_workbook


def load_config(config_path="config.json"):
    """config.json を読み込む。"""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def find_label_cell(ws, label_candidates):
    """ワークシート内でラベル候補に一致するセルを検索する。

    Returns:
        (row, col, matched_label) or None
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            cell_text = str(cell.value).strip()
            for label in label_candidates:
                if label in cell_text:
                    return cell.row, cell.column, cell_text
    return None


def find_target_cell(ws, label_info):
    """ラベル設定に基づいて、入力先セルの座標を特定する。

    Returns:
        (label_row, label_col, target_row, target_col, matched_label) or None
    """
    result = find_label_cell(ws, label_info["labels"])
    if result is None:
        return None

    label_row, label_col, matched_label = result
    offset_col = label_info.get("offset_col", 1)
    offset_row = label_info.get("offset_row", 0)

    target_row = label_row + offset_row
    target_col = label_col + offset_col

    return label_row, label_col, target_row, target_col, matched_label


def extract_name_from_filename(filename, config):
    """ファイル名から氏名（苗字）部分を抽出する。"""
    pattern = config.get("filename_pattern", "")
    match = re.search(pattern, filename)
    if match:
        raw = match.group(1)
        # 施設略称を除去
        abbrevs = config.get("facility_abbreviations", {})
        for abbr in abbrevs:
            if raw.startswith(abbr):
                return raw[len(abbr):]
        return raw
    return None


def analyze_excel(filepath, config):
    """Excelファイルを解析し、ラベルと入力先セルの位置を表示する。"""
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    print(f"\n  シート名: {ws.title}")
    print(f"  使用範囲: {ws.dimensions}")

    mapping = config["excel_label_mapping"]

    print(f"\n  {'フィールド':<14} {'ラベル位置':<12} {'入力先位置':<12} {'現在値':<20} {'状態'}")
    print(f"  {'-'*14} {'-'*12} {'-'*12} {'-'*20} {'-'*6}")

    for field_name, label_info in mapping.items():
        result = find_target_cell(ws, label_info)
        if result:
            label_row, label_col, target_row, target_col, matched_label = result
            current_val = ws.cell(row=target_row, column=target_col).value
            current_val_str = str(current_val) if current_val else "(空)"
            label_pos = f"{_col_letter(label_col)}{label_row}"
            target_pos = f"{_col_letter(target_col)}{target_row}"
            print(f"  {field_name:<14} {label_pos:<12} {target_pos:<12} {current_val_str:<20} OK")
        else:
            print(f"  {field_name:<14} {'---':<12} {'---':<12} {'---':<20} 未発見")

    wb.close()


def _col_letter(col_num):
    """列番号をアルファベットに変換する。"""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config = load_config(os.path.join(base_dir, "config.json"))

    # 対象ファイルの決定
    if len(sys.argv) > 1:
        excel_files = [sys.argv[1]]
    else:
        excel_files = [
            os.path.join(base_dir, f)
            for f in os.listdir(base_dir)
            if f.endswith(".xlsx") and "ヒアリングシート" in f
        ]

    if not excel_files:
        print("対象のExcelファイルが見つかりません。")
        sys.exit(1)

    print("=" * 70)
    print("Excelラベル検索テスト")
    print("=" * 70)

    for filepath in sorted(excel_files):
        filename = os.path.basename(filepath)
        name = extract_name_from_filename(filename, config)
        print(f"\n--- {filename} ---")
        print(f"  抽出氏名: {name}")
        analyze_excel(filepath, config)

    print("\n" + "=" * 70)
    print("テスト完了")
    print("=" * 70)


if __name__ == "__main__":
    main()
