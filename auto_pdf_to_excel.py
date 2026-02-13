#!/usr/bin/env python3
"""PDF利用者情報 → Excel ヒアリングシート 自動転記スクリプト。

input_pdf/ 内のPDFファイルを解析し、利用者氏名を基にカレントディレクトリの
Excelヒアリングシートへデータを自動転記する。

使い方:
    python auto_pdf_to_excel.py
    python auto_pdf_to_excel.py --config config.json --pdf-dir input_pdf --excel-dir .
    python auto_pdf_to_excel.py --dry-run   # 書き込みせずに確認のみ
"""

import argparse
import json
import logging
import os
import re
import sys
from datetime import datetime

import fitz  # pymupdf
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# ロギング設定
# ---------------------------------------------------------------------------
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def setup_logging(log_file=None):
    """ロギングを設定する。"""
    handlers = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT, datefmt=DATE_FORMAT, handlers=handlers)


# ---------------------------------------------------------------------------
# 設定読み込み
# ---------------------------------------------------------------------------
def load_config(config_path):
    """config.json を読み込む。"""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# PDF 解析
# ---------------------------------------------------------------------------
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


def parse_pdf_text(text, patterns):
    """抽出テキストから正規表現でフィールドを抽出し、辞書に構造化する。"""
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
    patterns = config["pdf_extraction_patterns"]
    data = parse_pdf_text(text, patterns)
    return data


def get_surname_from_pdf_data(data):
    """抽出データから苗字を取得する。"""
    full_name = data.get("利用者氏名", "")
    if not full_name:
        return None
    # 全角・半角スペースで分割し、最初の部分を苗字とする
    parts = re.split(r"[\s　]+", full_name)
    return parts[0] if parts else None


# ---------------------------------------------------------------------------
# Excel ファイル走査・照合
# ---------------------------------------------------------------------------
def find_excel_files(excel_dir):
    """ディレクトリ内のヒアリングシートExcelファイルを一覧する。"""
    files = []
    for f in os.listdir(excel_dir):
        if f.endswith(".xlsx") and "ヒアリングシート" in f:
            files.append(os.path.join(excel_dir, f))
    return sorted(files)


def extract_name_from_filename(filename, config):
    """ファイル名から氏名（苗字）部分を抽出する。"""
    pattern = config.get("filename_pattern", "")
    match = re.search(pattern, os.path.basename(filename))
    if match:
        raw = match.group(1)
        abbrevs = config.get("facility_abbreviations", {})
        for abbr in abbrevs:
            if raw.startswith(abbr):
                return raw[len(abbr):]
        return raw
    return None


def match_excel_file(surname, excel_files, config):
    """苗字に一致するExcelファイルを部分一致で探す。"""
    matches = []
    for filepath in excel_files:
        file_name_part = extract_name_from_filename(filepath, config)
        if file_name_part and surname in file_name_part:
            matches.append(filepath)
    return matches


# ---------------------------------------------------------------------------
# Excel ラベル検索・書き込み
# ---------------------------------------------------------------------------
def find_label_cell(ws, label_candidates):
    """ワークシート内でラベル候補に一致するセルを検索する。"""
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
    """ラベル設定に基づいて、入力先セルの座標を特定する。"""
    result = find_label_cell(ws, label_info["labels"])
    if result is None:
        return None

    label_row, label_col, matched_label = result
    offset_col = label_info.get("offset_col", 1)
    offset_row = label_info.get("offset_row", 0)

    target_row = label_row + offset_row
    target_col = label_col + offset_col

    return target_row, target_col, matched_label


def handle_care_level(ws, label_info, value):
    """要介護度の特別処理: 該当する選択肢のセルに「○」を入れる。"""
    result = find_label_cell(ws, label_info["labels"])
    if result is None:
        return False

    label_row, label_col, _ = result

    # まず、ラベルの右隣にも値を書き込む
    offset_col = label_info.get("offset_col", 1)
    offset_row = label_info.get("offset_row", 0)
    target_row = label_row + offset_row
    target_col = label_col + offset_col
    ws.cell(row=target_row, column=target_col).value = value

    # ラベル行の下にある選択肢をスキャンし、一致するものに「○」
    for scan_row in range(label_row + 1, label_row + 20):
        for scan_col in range(1, ws.max_column + 1):
            cell = ws.cell(row=scan_row, column=scan_col)
            if cell.value and value in str(cell.value):
                # チェック欄（次の列）に○を入力
                check_cell = ws.cell(row=scan_row, column=scan_col + 1)
                check_cell.value = "○"
                return True

    return True


def write_data_to_excel(filepath, data, config, dry_run=False):
    """抽出データをExcelファイルに書き込む。"""
    mapping = config["excel_label_mapping"]
    wb = load_workbook(filepath)
    ws = wb.active

    written_fields = []
    skipped_fields = []

    for field_name, label_info in mapping.items():
        value = data.get(field_name)
        if value is None:
            skipped_fields.append(field_name)
            continue

        # 要介護度の特別処理
        if label_info.get("special_handling") == "care_level":
            if not dry_run:
                success = handle_care_level(ws, label_info, value)
            else:
                success = find_label_cell(ws, label_info["labels"]) is not None
            if success:
                written_fields.append((field_name, value, "特別処理"))
            else:
                skipped_fields.append(field_name)
            continue

        # 通常のラベル検索 → 書き込み
        result = find_target_cell(ws, label_info)
        if result is None:
            skipped_fields.append(field_name)
            continue

        target_row, target_col, matched_label = result

        if not dry_run:
            ws.cell(row=target_row, column=target_col).value = value

        col_letter = _col_letter(target_col)
        written_fields.append((field_name, value, f"{col_letter}{target_row}"))

    if not dry_run:
        wb.save(filepath)

    wb.close()
    return written_fields, skipped_fields


def _col_letter(col_num):
    """列番号をアルファベットに変換する。"""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------
def process_all(config_path, pdf_dir, excel_dir, dry_run=False, log_file=None):
    """全PDFを処理してExcelに転記する。"""
    setup_logging(log_file)
    logger = logging.getLogger(__name__)

    config = load_config(config_path)
    mode_label = "【ドライラン】" if dry_run else ""

    logger.info(f"{'='*60}")
    logger.info(f"{mode_label}PDF → Excel 自動転記処理 開始")
    logger.info(f"{'='*60}")
    logger.info(f"PDF入力元:   {os.path.abspath(pdf_dir)}")
    logger.info(f"Excel出力先: {os.path.abspath(excel_dir)}")

    # PDF一覧
    pdf_files = [
        os.path.join(pdf_dir, f)
        for f in os.listdir(pdf_dir)
        if f.lower().endswith(".pdf")
    ]
    if not pdf_files:
        logger.warning("処理対象のPDFファイルが見つかりません。")
        return

    # Excel一覧
    excel_files = find_excel_files(excel_dir)
    if not excel_files:
        logger.warning("対象のExcelファイルが見つかりません。")
        return

    logger.info(f"PDF数: {len(pdf_files)}, Excel数: {len(excel_files)}")

    # 結果集計用
    success_list = []
    no_match_pdfs = []
    no_name_pdfs = []
    multi_match_pdfs = []
    unmatched_excels = set(excel_files)

    for pdf_path in sorted(pdf_files):
        pdf_name = os.path.basename(pdf_path)
        logger.info(f"\n--- 処理中: {pdf_name} ---")

        # PDF解析
        data = extract_data_from_pdf(pdf_path, config)
        surname = get_surname_from_pdf_data(data)

        if not surname:
            logger.warning(f"  氏名を抽出できませんでした: {pdf_name}")
            no_name_pdfs.append(pdf_name)
            continue

        logger.info(f"  抽出氏名: {data.get('利用者氏名', '?')} (苗字: {surname})")

        # Excel照合
        matched = match_excel_file(surname, excel_files, config)

        if not matched:
            logger.warning(f"  一致するExcelファイルが見つかりません: 苗字「{surname}」")
            no_match_pdfs.append((pdf_name, surname))
            continue

        if len(matched) > 1:
            logger.warning(f"  複数のExcelが一致しました ({len(matched)}件)。全てに書き込みます。")
            multi_match_pdfs.append((pdf_name, [os.path.basename(f) for f in matched]))

        for excel_path in matched:
            excel_name = os.path.basename(excel_path)
            logger.info(f"  転記先: {excel_name}")

            written, skipped = write_data_to_excel(excel_path, data, config, dry_run=dry_run)

            for field, value, pos in written:
                logger.info(f"    ✓ {field}: {value} → {pos}")

            if skipped:
                logger.info(f"    (スキップ: {', '.join(skipped)})")

            success_list.append((pdf_name, excel_name, len(written)))
            unmatched_excels.discard(excel_path)

    # ---------------------------------------------------------------------------
    # サマリー出力
    # ---------------------------------------------------------------------------
    logger.info(f"\n{'='*60}")
    logger.info(f"{mode_label}処理結果サマリー")
    logger.info(f"{'='*60}")

    logger.info(f"\n[成功] {len(success_list)} 件")
    for pdf_name, excel_name, count in success_list:
        logger.info(f"  {pdf_name} → {excel_name} ({count}項目転記)")

    if no_name_pdfs:
        logger.warning(f"\n[未処理: 氏名抽出失敗] {len(no_name_pdfs)} 件")
        for name in no_name_pdfs:
            logger.warning(f"  - {name}")

    if no_match_pdfs:
        logger.warning(f"\n[未処理: Excel照合失敗] {len(no_match_pdfs)} 件")
        for pdf_name, surname in no_match_pdfs:
            logger.warning(f"  - {pdf_name} (苗字: {surname})")

    if multi_match_pdfs:
        logger.info(f"\n[注意: 複数Excel一致] {len(multi_match_pdfs)} 件")
        for pdf_name, excel_names in multi_match_pdfs:
            logger.info(f"  - {pdf_name} → {', '.join(excel_names)}")

    if unmatched_excels:
        logger.info(f"\n[情報: 対応PDFなしのExcel] {len(unmatched_excels)} 件")
        for path in sorted(unmatched_excels):
            logger.info(f"  - {os.path.basename(path)}")

    logger.info(f"\n{'='*60}")
    logger.info("処理完了")
    logger.info(f"{'='*60}")


def main():
    parser = argparse.ArgumentParser(
        description="PDF利用者情報 → Excel ヒアリングシート 自動転記"
    )
    parser.add_argument(
        "--config", default="config.json",
        help="設定ファイルパス (デフォルト: config.json)"
    )
    parser.add_argument(
        "--pdf-dir", default="input_pdf",
        help="PDF入力ディレクトリ (デフォルト: input_pdf)"
    )
    parser.add_argument(
        "--excel-dir", default=".",
        help="Excelファイルディレクトリ (デフォルト: カレント)"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="ドライラン（書き込みせず確認のみ）"
    )
    parser.add_argument(
        "--log-file", default=None,
        help="ログファイルパス（指定時はファイルにも出力）"
    )
    args = parser.parse_args()

    process_all(
        config_path=args.config,
        pdf_dir=args.pdf_dir,
        excel_dir=args.excel_dir,
        dry_run=args.dry_run,
        log_file=args.log_file,
    )


if __name__ == "__main__":
    main()
