#!/usr/bin/env python3
"""PDF利用者情報 → Excel ヒアリングシート 自動転記スクリプト。

input_pdf/ 内のPDFファイルを解析し、利用者氏名を基に excel_sheets/ 内の
Excelヒアリングシートへデータを自動転記する。

使い方:
    python auto_pdf_to_excel.py
    python auto_pdf_to_excel.py --config config.json --pdf-dir input_pdf --excel-dir excel_sheets
    python auto_pdf_to_excel.py --dry-run   # 書き込みせずに確認のみ
"""

import argparse
import json
import logging
import os
import re
import shutil
import sys
import unicodedata
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
# テキスト正規化
# ---------------------------------------------------------------------------
def normalize_text(text):
    """全角英数字・記号を半角に正規化する。

    NFKC正規化により、全角数字(１２３)→半角(123)、
    全角英字(ＡＢＣ)→半角(ABC) などを統一する。
    """
    return unicodedata.normalize("NFKC", text)


# ---------------------------------------------------------------------------
# バックアップ・ファイル整理
# ---------------------------------------------------------------------------
def backup_excel(filepath, backup_base_dir):
    """Excelファイルをバックアップフォルダへコピーする。

    Args:
        filepath: バックアップ対象のExcelファイルパス。
        backup_base_dir: バックアップの親ディレクトリ。

    Returns:
        バックアップ先のパス。
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = os.path.join(backup_base_dir, timestamp)
    os.makedirs(backup_dir, exist_ok=True)

    dest = os.path.join(backup_dir, os.path.basename(filepath))
    shutil.copy2(filepath, dest)
    return dest


def move_processed_pdf(pdf_path, processed_base_dir):
    """処理済みPDFを processed/YYYYMMDD/ フォルダへ移動する。

    Args:
        pdf_path: 移動対象のPDFファイルパス。
        processed_base_dir: 処理済みフォルダの親ディレクトリ。

    Returns:
        移動先のパス。
    """
    date_str = datetime.now().strftime("%Y%m%d")
    dest_dir = os.path.join(processed_base_dir, date_str)
    os.makedirs(dest_dir, exist_ok=True)

    dest = os.path.join(dest_dir, os.path.basename(pdf_path))
    # 同名ファイルが既にある場合はサフィックスを付与
    if os.path.exists(dest):
        base, ext = os.path.splitext(os.path.basename(pdf_path))
        counter = 1
        while os.path.exists(dest):
            dest = os.path.join(dest_dir, f"{base}_{counter}{ext}")
            counter += 1
    shutil.move(pdf_path, dest)
    return dest


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
    """抽出テキストから正規表現でフィールドを抽出し、辞書に構造化する。

    全角英数字は半角に正規化してからマッチングを行う。
    """
    normalized = normalize_text(text)
    result = {}
    for field_name, regex_list in patterns.items():
        for pattern in regex_list:
            normalized_pattern = normalize_text(pattern)
            match = re.search(normalized_pattern, normalized)
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
# 要介護度の正規化
# ---------------------------------------------------------------------------
def normalize_care_level(value, config):
    """PDF抽出値をconfig内の正規名称に正規化する。

    例: "要介護２" → "要介護2", "介護3" → "要介護3"
    """
    normalized_value = normalize_text(value)
    care_map = config.get("care_level_mapping", {})
    for canonical, aliases in care_map.items():
        normalized_aliases = [normalize_text(a) for a in aliases]
        if normalized_value in normalized_aliases or normalized_value == canonical:
            return canonical
    return normalized_value


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


def handle_care_level(ws, label_info, value, config):
    """要介護度の特別処理: 正規化後にラベル右隣と選択肢セルへ書き込む。"""
    result = find_label_cell(ws, label_info["labels"])
    if result is None:
        return False

    label_row, label_col, _ = result
    canonical = normalize_care_level(value, config)

    # ラベルの右隣に正規名称を書き込む
    offset_col = label_info.get("offset_col", 1)
    offset_row = label_info.get("offset_row", 0)
    target_row = label_row + offset_row
    target_col = label_col + offset_col
    ws.cell(row=target_row, column=target_col).value = canonical

    # ラベル行の下にある選択肢をスキャンし、柔軟に照合
    care_map = config.get("care_level_mapping", {})
    match_aliases = care_map.get(canonical, [canonical])
    # 半角正規化したエイリアスも用意
    all_match_values = set()
    for alias in match_aliases:
        all_match_values.add(alias)
        all_match_values.add(normalize_text(alias))
    all_match_values.add(canonical)

    for scan_row in range(label_row + 1, label_row + 20):
        for scan_col in range(1, ws.max_column + 1):
            cell = ws.cell(row=scan_row, column=scan_col)
            if cell.value is None:
                continue
            cell_text = normalize_text(str(cell.value).strip())
            if cell_text in all_match_values:
                check_cell = ws.cell(row=scan_row, column=scan_col + 1)
                check_cell.value = "○"
                return True

    return True


def write_data_to_excel(filepath, data, config, dry_run=False):
    """抽出データをExcelファイルに書き込む。

    空値(None)の場合は既存データを保護し上書きしない。
    """
    mapping = config["excel_label_mapping"]
    wb = load_workbook(filepath)
    ws = wb.active

    written_fields = []
    skipped_fields = []
    protected_fields = []

    for field_name, label_info in mapping.items():
        value = data.get(field_name)

        # 空欄保護: 抽出結果がNoneまたは空文字の場合は既存値を保持
        if value is None or (isinstance(value, str) and value.strip() == ""):
            # ラベルが存在するかだけチェックして保護ログを出す
            result = find_label_cell(ws, label_info["labels"])
            if result:
                target_row = result[0] + label_info.get("offset_row", 0)
                target_col = result[1] + label_info.get("offset_col", 1)
                existing = ws.cell(row=target_row, column=target_col).value
                if existing:
                    protected_fields.append((field_name, str(existing)))
                else:
                    skipped_fields.append(field_name)
            else:
                skipped_fields.append(field_name)
            continue

        # 要介護度の特別処理
        if label_info.get("special_handling") == "care_level":
            if not dry_run:
                success = handle_care_level(ws, label_info, value, config)
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
    return written_fields, skipped_fields, protected_fields


def _col_letter(col_num):
    """列番号をアルファベットに変換する。"""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ---------------------------------------------------------------------------
# レポート出力
# ---------------------------------------------------------------------------
def generate_report(report_path, success_list, no_name_pdfs, no_match_pdfs,
                    multi_match_pdfs, unmatched_excels, backup_log, move_log,
                    protected_log, dry_run=False):
    """簡易レポートファイルを生成する。"""
    mode_label = "【ドライラン】" if dry_run else ""
    lines = []
    lines.append(f"{mode_label}PDF → Excel 自動転記 実行レポート")
    lines.append(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 60)

    # 成功
    lines.append(f"\n■ 成功: {len(success_list)} 件")
    for pdf_name, excel_name, count in success_list:
        lines.append(f"  {pdf_name} → {excel_name} ({count}項目転記)")

    # 空欄保護
    if protected_log:
        lines.append(f"\n■ 空欄保護（既存値を維持）: {len(protected_log)} 件")
        for pdf_name, field, existing in protected_log:
            lines.append(f"  {pdf_name}: {field} (既存値「{existing}」を保護)")

    # 氏名抽出失敗
    if no_name_pdfs:
        lines.append(f"\n■ 失敗（氏名抽出不可）: {len(no_name_pdfs)} 件")
        for name in no_name_pdfs:
            lines.append(f"  - {name}")

    # Excel照合失敗
    if no_match_pdfs:
        lines.append(f"\n■ 失敗（対応Excel未発見）: {len(no_match_pdfs)} 件")
        for pdf_name, surname in no_match_pdfs:
            lines.append(f"  - {pdf_name} (苗字: {surname})")

    # 複数一致
    if multi_match_pdfs:
        lines.append(f"\n■ 注意（複数Excel一致）: {len(multi_match_pdfs)} 件")
        for pdf_name, excel_names in multi_match_pdfs:
            lines.append(f"  - {pdf_name} → {', '.join(excel_names)}")

    # 対応PDFなしExcel
    if unmatched_excels:
        lines.append(f"\n■ スキップ（対応PDFなしのExcel）: {len(unmatched_excels)} 件")
        for path in sorted(unmatched_excels):
            lines.append(f"  - {os.path.basename(path)}")

    # バックアップログ
    if backup_log:
        lines.append(f"\n■ バックアップ: {len(backup_log)} 件")
        for src, dest in backup_log:
            lines.append(f"  {os.path.basename(src)} → {dest}")

    # PDF移動ログ
    if move_log:
        lines.append(f"\n■ 処理済みPDF移動: {len(move_log)} 件")
        for src, dest in move_log:
            lines.append(f"  {os.path.basename(src)} → {dest}")

    lines.append("\n" + "=" * 60)

    content = "\n".join(lines)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(content)
    return report_path


# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------
def process_all(config_path, pdf_dir, excel_dir, dry_run=False, log_file=None):
    """全PDFを処理してExcelに転記する。"""
    setup_logging(log_file)
    logger = logging.getLogger(__name__)

    config = load_config(config_path)
    mode_label = "【ドライラン】" if dry_run else ""
    now = datetime.now()

    logger.info(f"{'='*60}")
    logger.info(f"{mode_label}PDF → Excel 自動転記処理 開始")
    logger.info(f"{'='*60}")
    logger.info(f"PDF入力元:   {os.path.abspath(pdf_dir)}")
    logger.info(f"Excel出力先: {os.path.abspath(excel_dir)}")

    # ディレクトリが存在しない場合は自動作成
    os.makedirs(pdf_dir, exist_ok=True)
    os.makedirs(excel_dir, exist_ok=True)

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

    # バックアップ・整理用ディレクトリ
    backup_base = os.path.join(excel_dir, "backups")
    processed_base = os.path.join(excel_dir, "processed")

    # 結果集計用
    success_list = []
    no_match_pdfs = []
    no_name_pdfs = []
    multi_match_pdfs = []
    unmatched_excels = set(excel_files)
    backup_log = []
    move_log = []
    protected_log = []
    processed_pdf_paths = []  # 転記成功したPDFパスを記録

    for pdf_path in sorted(pdf_files):
        pdf_name = os.path.basename(pdf_path)
        logger.info(f"\n--- 処理中: {pdf_name} ---")

        # PDF解析
        try:
            data = extract_data_from_pdf(pdf_path, config)
        except Exception as e:
            logger.error(f"  PDF解析エラー: {pdf_name} - {e}")
            no_name_pdfs.append(pdf_name)
            continue

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

        pdf_success = False
        for excel_path in matched:
            excel_name = os.path.basename(excel_path)

            # バックアップ（実行前）
            if not dry_run:
                try:
                    backup_dest = backup_excel(excel_path, backup_base)
                    logger.info(f"  バックアップ: {excel_name} → {backup_dest}")
                    backup_log.append((excel_path, backup_dest))
                except Exception as e:
                    logger.error(f"  バックアップ失敗: {excel_name} - {e}")
                    continue

            logger.info(f"  転記先: {excel_name}")

            try:
                written, skipped, protected = write_data_to_excel(
                    excel_path, data, config, dry_run=dry_run
                )
            except Exception as e:
                logger.error(f"  Excel書き込みエラー: {excel_name} - {e}")
                continue

            for field, value, pos in written:
                logger.info(f"    ✓ {field}: {value} → {pos}")

            for field, existing in protected:
                logger.info(f"    🛡 {field}: 既存値「{existing}」を保護")
                protected_log.append((pdf_name, field, existing))

            if skipped:
                logger.info(f"    (スキップ: {', '.join(skipped)})")

            success_list.append((pdf_name, excel_name, len(written)))
            unmatched_excels.discard(excel_path)
            pdf_success = True

        if pdf_success:
            processed_pdf_paths.append(pdf_path)

    # 処理済みPDFを移動
    if not dry_run:
        for pdf_path in processed_pdf_paths:
            try:
                move_dest = move_processed_pdf(pdf_path, processed_base)
                logger.info(f"  PDF移動: {os.path.basename(pdf_path)} → {move_dest}")
                move_log.append((pdf_path, move_dest))
            except Exception as e:
                logger.error(f"  PDF移動失敗: {os.path.basename(pdf_path)} - {e}")

    # ---------------------------------------------------------------------------
    # サマリー出力
    # ---------------------------------------------------------------------------
    logger.info(f"\n{'='*60}")
    logger.info(f"{mode_label}処理結果サマリー")
    logger.info(f"{'='*60}")

    logger.info(f"\n[成功] {len(success_list)} 件")
    for pdf_name, excel_name, count in success_list:
        logger.info(f"  {pdf_name} → {excel_name} ({count}項目転記)")

    if protected_log:
        logger.info(f"\n[空欄保護] {len(protected_log)} 件")
        for pdf_name, field, existing in protected_log:
            logger.info(f"  {pdf_name}: {field} (既存値「{existing}」を保護)")

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

    if backup_log:
        logger.info(f"\n[バックアップ] {len(backup_log)} 件作成済み")

    if move_log:
        logger.info(f"[PDF移動] {len(move_log)} 件を processed/ へ移動")

    # レポートファイル生成
    report_name = f"report_{now.strftime('%Y%m%d')}.txt"
    report_path = os.path.join(excel_dir, report_name)
    generate_report(
        report_path, success_list, no_name_pdfs, no_match_pdfs,
        multi_match_pdfs, unmatched_excels, backup_log, move_log,
        protected_log, dry_run=dry_run,
    )
    logger.info(f"\n[レポート] {report_path}")

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
        "--excel-dir", default="excel_sheets",
        help="Excelファイルディレクトリ (デフォルト: excel_sheets)"
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
