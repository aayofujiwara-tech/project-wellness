#!/usr/bin/env python3
"""利用者情報 → Excel ヒアリングシート 自動転記スクリプト。

入力ソース:
    --source pdf   : input_pdf/ 内のPDFから利用者情報を抽出して転記（従来方式）
    --source excel : input_excel/ 内の「基本情報シート」Excelから利用者情報を抽出して転記

使い方:
    python auto_pdf_to_excel.py                            # PDF→ヒアリングシート (form1)
    python auto_pdf_to_excel.py --source excel             # 基本情報Excel→ヒアリングシート
    python auto_pdf_to_excel.py --source excel --dry-run   # ドライラン
    python auto_pdf_to_excel.py --form form2               # 下段(form2)に転記
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
    """全角英数字・記号を半角に正規化する。"""
    return unicodedata.normalize("NFKC", text)


# ---------------------------------------------------------------------------
# バックアップ・ファイル整理
# ---------------------------------------------------------------------------
def backup_excel(filepath, backup_base_dir):
    """Excelファイルをバックアップフォルダへコピーする。"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = os.path.join(backup_base_dir, timestamp)
    os.makedirs(backup_dir, exist_ok=True)
    dest = os.path.join(backup_dir, os.path.basename(filepath))
    shutil.copy2(filepath, dest)
    return dest


def move_processed_file(file_path, processed_base_dir):
    """処理済みファイルを processed/YYYYMMDD/ フォルダへ移動する。"""
    date_str = datetime.now().strftime("%Y%m%d")
    dest_dir = os.path.join(processed_base_dir, date_str)
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, os.path.basename(file_path))
    if os.path.exists(dest):
        base, ext = os.path.splitext(os.path.basename(file_path))
        counter = 1
        while os.path.exists(dest):
            dest = os.path.join(dest_dir, f"{base}_{counter}{ext}")
            counter += 1
    shutil.move(file_path, dest)
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
    """抽出テキストから正規表現でフィールドを抽出し、辞書に構造化する。"""
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
    return parse_pdf_text(text, patterns)


# ---------------------------------------------------------------------------
# 基本情報シート（Excel）解析 – パーサー群
# ---------------------------------------------------------------------------
def parse_age_from_template(cell_value):
    """'年齢　　　　　　68　　歳' → '68'"""
    if not cell_value:
        return None
    normalized = normalize_text(str(cell_value))
    match = re.search(r"(\d+)\s*歳", normalized)
    return match.group(1) if match else None


def parse_era_date(cell_value):
    """元号テンプレートから生年月日を抽出する。

    例: '明治・大正・昭和・平成　　　　32年　　４月　　２日'
    → 年月日の数値を抽出し、年数から元号を推定して 'X和YY年M月D日' を返す。
    """
    if not cell_value:
        return None
    normalized = normalize_text(str(cell_value))
    match = re.search(r"(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日", normalized)
    if not match:
        return None
    year = int(match.group(1))
    month = int(match.group(2))
    day = int(match.group(3))

    # 元号推定: テキスト中に強調マークがないため年数と範囲で推定
    # 大正: 1-15, 昭和: 1-64, 平成: 1-31, 令和: 1-
    # 高齢者利用が大半のため、大正・昭和を優先
    if "令和" in normalized and year <= 20:
        era = "令和"
    elif "平成" in normalized and year <= 31:
        # 平成か昭和かは年数だけでは曖昧 → 年齢から逆算を試みる
        era = "平成" if year <= 31 else "昭和"
    else:
        # デフォルト: 年数から推定
        if year >= 1 and year <= 15:
            era = "大正"  # 大正と昭和の区別は困難だが、高齢者なら大正もあり得る
        elif year >= 1 and year <= 64:
            era = "昭和"
        else:
            era = "昭和"

    return f"{era}{year}年{month}月{day}日"


def parse_gender_template(cell_value):
    """性別テンプレート '男　・　女' → 空文字を返す。

    基本情報シートでは性別は未入力（テンプレ文字列のまま）のため空扱い。
    """
    return ""


# セルパーサーのディスパッチテーブル
_CELL_PARSERS = {
    "age_from_template": parse_age_from_template,
    "era_date": parse_era_date,
    "gender_template": parse_gender_template,
}


# ---------------------------------------------------------------------------
# 基本情報シート（Excel）解析 – 抽出
# ---------------------------------------------------------------------------
def _is_dummy_name(value):
    """氏名セルの値がダミーデータかどうか判定する。

    空文字、None、「年  月  日」系パターン、記号のみの場合 True を返す。
    """
    if not value:
        return True
    s = str(value).strip()
    if not s:
        return True
    # 「年  月  日」「年 月 日」「年月日」系
    if re.match(r"^年\s*月\s*日$", s):
        return True
    # 記号・空白のみ（全角スペース含む）
    if re.match(r"^[\s　\-\−\—\–\/_・．.、。○◯●■□△▲▽▼※＊]+$", s):
        return True
    return False


def _clean_name(name):
    """氏名から末尾の「様」を除去する。"""
    if not name:
        return name
    return re.sub(r"[　\s]*様$", "", name).strip()


def extract_data_from_excel(excel_path, config, logger=None):
    """基本情報シートExcelの全シートから利用者データを抽出する。

    config["excel_source_cell_mapping"] で定義されたセル座標から値を読み取り、
    config["skip_sheets"] に含まれるシートはスキップする。
    氏名セルがダミーデータのシートも除外する。

    Returns:
        list[tuple[str, dict]]: (シート名, データ辞書) のリスト。
        有効なシートがなければ空リスト。
    """
    cell_map = config["excel_source_cell_mapping"]
    skip_sheets = config.get("skip_sheets", ["原本"])
    wb = load_workbook(excel_path, data_only=True)
    results = []

    for ws in wb.worksheets:
        sheet_name = ws.title

        # スキップ対象シート
        if sheet_name in skip_sheets:
            if logger:
                logger.info(f"  [SKIP] シート「{sheet_name}」: スキップ対象のためスキップ")
            continue

        # 氏名セルのバリデーション
        name_info = cell_map.get("利用者氏名", {"cell": "E9"})
        name_cell = name_info["cell"] if isinstance(name_info, dict) else name_info
        raw_name = ws[name_cell].value
        if _is_dummy_name(raw_name):
            reason = "空欄" if not raw_name or not str(raw_name).strip() else f"ダミーデータ「{str(raw_name).strip()}」"
            if logger:
                logger.info(f"  [SKIP] シート「{sheet_name}」: {reason}")
            continue

        # データ抽出
        data = {}
        for field_name, field_info in cell_map.items():
            if field_name.startswith("_"):
                continue  # _comment 等をスキップ
            if isinstance(field_info, dict):
                cell_addr = field_info["cell"]
                parse_name = field_info.get("parse")
            else:
                cell_addr = field_info
                parse_name = None

            value = ws[cell_addr].value

            if parse_name and parse_name in _CELL_PARSERS:
                value = _CELL_PARSERS[parse_name](value)
            elif value is not None:
                value = normalize_text(str(value).strip())
            else:
                value = None

            if value is not None and value != "":
                data[field_name] = value

        # 氏名の「様」を除去
        if "利用者氏名" in data:
            data["利用者氏名"] = _clean_name(data["利用者氏名"])

        if logger:
            logger.info(f"  [OK] シート「{sheet_name}」: 氏名={data.get('利用者氏名', '?')}")

        results.append((sheet_name, data))

    wb.close()
    return results


def find_source_excel_files(input_dir, config):
    """ディレクトリ内の基本情報シートExcelファイルを一覧する。

    config["source_file_pattern"] の正規表現でマッチングする。
    """
    pattern = config.get("source_file_pattern", r"基本情報シート.*\.xlsx$")
    files = []
    for f in os.listdir(input_dir):
        if re.search(pattern, f):
            files.append(os.path.join(input_dir, f))
    return sorted(files)


# ---------------------------------------------------------------------------
# 共通: 苗字抽出
# ---------------------------------------------------------------------------
def get_surname(data):
    """抽出データから苗字を取得する。

    氏名末尾の「様」を除去したうえで苗字部分を返す。
    """
    full_name = data.get("利用者氏名", "")
    if not full_name:
        return None
    full_name = _clean_name(full_name)
    if not full_name:
        return None
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
    """抽出値をconfig内の正規名称に正規化する。"""
    normalized_value = normalize_text(value)
    care_map = config.get("care_level_mapping", {})
    for canonical, aliases in care_map.items():
        normalized_aliases = [normalize_text(a) for a in aliases]
        if normalized_value in normalized_aliases or normalized_value == canonical:
            return canonical
    return normalized_value


# ---------------------------------------------------------------------------
# フォーマット関数群
# ---------------------------------------------------------------------------
def format_gender(value):
    """性別を短縮形に変換する。男性→男, 女性→女"""
    normalized = normalize_text(value.strip())
    if "男" in normalized:
        return "男"
    elif "女" in normalized:
        return "女"
    return normalized


def format_age(value):
    """年齢に「歳」を付与する。既に「歳」があればそのまま。"""
    normalized = normalize_text(str(value).strip())
    age_match = re.search(r"(\d+)", normalized)
    if age_match:
        return f"{age_match.group(1)}歳"
    return normalized


def format_furigana(value):
    """ふりがなの括弧を除去する。"""
    cleaned = re.sub(r"[（()）]", "", value)
    return cleaned.strip()


def format_care_level_inline(value, config):
    """要介護度を丸数字に変換してインライン文字列を返す。

    例: "要介護3" → "要支援（ １・２ ）　／　要介護（ 1・2・③・4・5 ）"
    """
    canonical = normalize_care_level(value, config)
    circled = {"1": "①", "2": "②", "3": "③", "4": "④", "5": "⑤"}

    support_match = re.match(r"要支援(\d)", canonical)
    care_match = re.match(r"要介護(\d)", canonical)

    if support_match:
        num = support_match.group(1)
        support_part = "１・２"
        if num == "1":
            support_part = "①・２"
        elif num == "2":
            support_part = "１・②"
        return f"要支援（ {support_part} ）　／　要介護（ 1・2・3・4・5 ）"

    elif care_match:
        num = care_match.group(1)
        nums_list = ["1", "2", "3", "4", "5"]
        formatted_nums = []
        for n in nums_list:
            if n == num:
                formatted_nums.append(circled[n])
            else:
                formatted_nums.append(n)
        care_part = "・".join(formatted_nums)
        return f"要支援（ １・２ ）　／　要介護（ {care_part} ）"

    return canonical


# ---------------------------------------------------------------------------
# Excel 座標指定書き込み (fixed_cell_mapping 方式)
# ---------------------------------------------------------------------------
def write_data_to_excel(filepath, data, config, form_name="form1", dry_run=False):
    """固定座標指定方式でExcelにデータを書き込む。

    config["fixed_cell_mapping"][form_name] のセル座標に直接書き込む。
    空値(None)の場合は既存データを保護し上書きしない。
    """
    cell_mapping = config["fixed_cell_mapping"][form_name]
    target_sheet = config.get("target_sheet")
    wb = load_workbook(filepath)

    if target_sheet and target_sheet in wb.sheetnames:
        ws = wb[target_sheet]
    else:
        ws = wb.active

    format_functions = {
        "format_gender": format_gender,
        "format_age": format_age,
        "format_furigana": format_furigana,
        "format_care_level_inline": lambda v: format_care_level_inline(v, config),
    }

    written_fields = []
    skipped_fields = []
    protected_fields = []

    for field_name, field_info in cell_mapping.items():
        value = data.get(field_name)
        cell_addr = field_info["cell"]

        # 空欄保護: 抽出結果がNoneまたは空文字の場合は既存値を保持
        if value is None or (isinstance(value, str) and value.strip() == ""):
            existing = ws[cell_addr].value
            if existing and str(existing).strip():
                protected_fields.append((field_name, str(existing)))
            else:
                skipped_fields.append(field_name)
            continue

        # フォーマット関数を適用
        fmt = field_info.get("format")
        if fmt and fmt in format_functions:
            value = format_functions[fmt](value)

        if not dry_run:
            ws[cell_addr].value = value

        written_fields.append((field_name, value, cell_addr))

    if not dry_run:
        wb.save(filepath)

    wb.close()
    return written_fields, skipped_fields, protected_fields


# ---------------------------------------------------------------------------
# レポート出力
# ---------------------------------------------------------------------------
def generate_report(report_path, success_list, no_name_list, no_match_list,
                    multi_match_list, unmatched_excels, backup_log, move_log,
                    protected_log, mode="pdf", dry_run=False):
    """簡易レポートファイルを生成する。"""
    mode_label = "【ドライラン】" if dry_run else ""
    source_label = "PDF" if mode == "pdf" else "基本情報Excel"
    lines = []
    lines.append(f"{mode_label}{source_label} → ヒアリングシート 自動転記 実行レポート")
    lines.append(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 60)

    lines.append(f"\n■ 成功: {len(success_list)} 件")
    for src_name, excel_name, count in success_list:
        lines.append(f"  {src_name} → {excel_name} ({count}項目転記)")

    if protected_log:
        lines.append(f"\n■ 空欄保護（既存値を維持）: {len(protected_log)} 件")
        for src_name, field, existing in protected_log:
            lines.append(f"  {src_name}: {field} (既存値「{existing}」を保護)")

    if no_name_list:
        lines.append(f"\n■ 失敗（氏名抽出不可）: {len(no_name_list)} 件")
        for name in no_name_list:
            lines.append(f"  - {name}")

    if no_match_list:
        lines.append(f"\n■ 失敗（対応Excel未発見）: {len(no_match_list)} 件")
        for src_name, surname in no_match_list:
            lines.append(f"  - {src_name} (苗字: {surname})")

    if multi_match_list:
        lines.append(f"\n■ 注意（複数Excel一致）: {len(multi_match_list)} 件")
        for src_name, excel_names in multi_match_list:
            lines.append(f"  - {src_name} → {', '.join(excel_names)}")

    if unmatched_excels:
        lines.append(f"\n■ スキップ（対応入力なしのExcel）: {len(unmatched_excels)} 件")
        for path in sorted(unmatched_excels):
            lines.append(f"  - {os.path.basename(path)}")

    if backup_log:
        lines.append(f"\n■ バックアップ: {len(backup_log)} 件")
        for src, dest in backup_log:
            lines.append(f"  {os.path.basename(src)} → {dest}")

    if move_log:
        lines.append(f"\n■ 処理済みファイル移動: {len(move_log)} 件")
        for src, dest in move_log:
            lines.append(f"  {os.path.basename(src)} → {dest}")

    lines.append("\n" + "=" * 60)

    content = "\n".join(lines)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(content)
    return report_path


# ---------------------------------------------------------------------------
# 共通転記ループ
# ---------------------------------------------------------------------------
def _process_inputs(input_items, excel_files, config, excel_dir, form_name,
                    mode, dry_run, logger):
    """入力ソース(PDF or 基本情報Excel)のリストを処理する共通ループ。

    Args:
        input_items: [(file_path, data_dict), ...] のリスト。
                     data_dictはextract済みの構造化データ。Noneなら抽出失敗。
    Returns:
        (success_list, no_name_list, no_match_list, multi_match_list,
         unmatched_excels, backup_log, move_log, protected_log,
         processed_paths)
    """
    backup_base = os.path.join(excel_dir, "backups")
    processed_base = os.path.join(excel_dir, "processed")

    success_list = []
    no_match_list = []
    no_name_list = []
    multi_match_list = []
    unmatched_excels = set(excel_files)
    backup_log = []
    move_log = []
    protected_log = []
    processed_paths = []

    for file_path, data in input_items:
        file_name = os.path.basename(file_path)

        if data is None:
            # 抽出失敗（呼び出し元で既にログ済み）
            no_name_list.append(file_name)
            continue

        surname = get_surname(data)

        if not surname:
            logger.warning(f"  氏名を抽出できませんでした: {file_name}")
            no_name_list.append(file_name)
            continue

        logger.info(f"  抽出氏名: {data.get('利用者氏名', '?')} (苗字: {surname})")

        matched = match_excel_file(surname, excel_files, config)

        if not matched:
            logger.warning(f"  一致するExcelファイルが見つかりません: 苗字「{surname}」")
            no_match_list.append((file_name, surname))
            continue

        if len(matched) > 1:
            logger.warning(f"  複数のExcelが一致しました ({len(matched)}件)。全てに書き込みます。")
            multi_match_list.append((file_name, [os.path.basename(f) for f in matched]))

        file_success = False
        for excel_path in matched:
            excel_name = os.path.basename(excel_path)

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
                    excel_path, data, config,
                    form_name=form_name, dry_run=dry_run,
                )
            except Exception as e:
                logger.error(f"  Excel書き込みエラー: {excel_name} - {e}")
                continue

            for field, value, pos in written:
                logger.info(f"    ✓ {field}: {value} → {pos}")

            for field, existing in protected:
                logger.info(f"    🛡 {field}: 既存値「{existing}」を保護")
                protected_log.append((file_name, field, existing))

            if skipped:
                logger.info(f"    (スキップ: {', '.join(skipped)})")

            success_list.append((file_name, excel_name, len(written)))
            unmatched_excels.discard(excel_path)
            file_success = True

        if file_success:
            processed_paths.append(file_path)

    # 処理済みファイルを移動（重複排除）
    if not dry_run:
        for path in sorted(set(processed_paths)):
            try:
                move_dest = move_processed_file(path, processed_base)
                logger.info(f"  移動: {os.path.basename(path)} → {move_dest}")
                move_log.append((path, move_dest))
            except Exception as e:
                if not os.path.exists(path):
                    continue
                logger.error(f"  移動失敗: {os.path.basename(path)} - {e}")

    return (success_list, no_name_list, no_match_list, multi_match_list,
            unmatched_excels, backup_log, move_log, protected_log)


# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------
def process_all(config_path, input_dir, excel_dir, form_name="form1",
                mode="pdf", dry_run=False, log_file=None):
    """入力ソースを処理してヒアリングシートに転記する。"""
    setup_logging(log_file)
    logger = logging.getLogger(__name__)

    config = load_config(config_path)

    # フォーム名の存在チェック
    if form_name not in config.get("fixed_cell_mapping", {}):
        logger.error(f"フォーム「{form_name}」が config.json の fixed_cell_mapping に存在しません。")
        sys.exit(1)

    mode_label = "【ドライラン】" if dry_run else ""
    source_label = "PDF" if mode == "pdf" else "基本情報Excel"
    now = datetime.now()

    logger.info(f"{'='*60}")
    logger.info(f"{mode_label}{source_label} → ヒアリングシート 自動転記処理 開始")
    logger.info(f"{'='*60}")
    logger.info(f"入力モード:  {mode}")
    logger.info(f"入力元:      {os.path.abspath(input_dir)}")
    logger.info(f"Excel出力先: {os.path.abspath(excel_dir)}")
    logger.info(f"フォーム:    {form_name}")

    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(excel_dir, exist_ok=True)

    # ヒアリングシート一覧
    excel_files = find_excel_files(excel_dir)
    if not excel_files:
        logger.warning("対象のヒアリングシートExcelファイルが見つかりません。")
        return

    # --- 入力ソースの抽出 ---
    input_items = []  # [(file_path, data_or_None), ...]

    if mode == "pdf":
        pdf_files = [
            os.path.join(input_dir, f)
            for f in os.listdir(input_dir)
            if f.lower().endswith(".pdf")
        ]
        if not pdf_files:
            logger.warning("処理対象のPDFファイルが見つかりません。")
            return

        logger.info(f"PDF数: {len(pdf_files)}, ヒアリングシート数: {len(excel_files)}")

        for pdf_path in sorted(pdf_files):
            pdf_name = os.path.basename(pdf_path)
            logger.info(f"\n--- 処理中: {pdf_name} ---")
            try:
                data = extract_data_from_pdf(pdf_path, config)
            except Exception as e:
                logger.error(f"  PDF解析エラー: {pdf_name} - {e}")
                data = None
            input_items.append((pdf_path, data))

    elif mode == "excel":
        source_files = find_source_excel_files(input_dir, config)
        if not source_files:
            logger.warning("処理対象の基本情報シートが見つかりません。")
            return

        logger.info(f"基本情報シート数: {len(source_files)}, ヒアリングシート数: {len(excel_files)}")

        for src_path in source_files:
            src_name = os.path.basename(src_path)
            logger.info(f"\n--- 処理中: {src_name} ---")
            try:
                sheet_results = extract_data_from_excel(src_path, config, logger)
            except Exception as e:
                logger.error(f"  基本情報シート解析エラー: {src_name} - {e}")
                sheet_results = []
            if not sheet_results:
                logger.warning(f"  有効なデータシートなし: {src_name}")
            for sheet_name, data in sheet_results:
                input_items.append((src_path, data))

    else:
        logger.error(f"不明なモード: {mode}")
        sys.exit(1)

    # --- 共通転記ループ ---
    (success_list, no_name_list, no_match_list, multi_match_list,
     unmatched_excels, backup_log, move_log, protected_log) = _process_inputs(
        input_items, excel_files, config, excel_dir, form_name,
        mode, dry_run, logger,
    )

    # --- サマリー出力 ---
    logger.info(f"\n{'='*60}")
    logger.info(f"{mode_label}処理結果サマリー")
    logger.info(f"{'='*60}")

    logger.info(f"\n[成功] {len(success_list)} 件")
    for src_name, excel_name, count in success_list:
        logger.info(f"  {src_name} → {excel_name} ({count}項目転記)")

    if protected_log:
        logger.info(f"\n[空欄保護] {len(protected_log)} 件")
        for src_name, field, existing in protected_log:
            logger.info(f"  {src_name}: {field} (既存値「{existing}」を保護)")

    if no_name_list:
        logger.warning(f"\n[未処理: 氏名抽出失敗] {len(no_name_list)} 件")
        for name in no_name_list:
            logger.warning(f"  - {name}")

    if no_match_list:
        logger.warning(f"\n[未処理: ヒアリングシート照合失敗] {len(no_match_list)} 件")
        for src_name, surname in no_match_list:
            logger.warning(f"  - {src_name} (苗字: {surname})")

    if multi_match_list:
        logger.info(f"\n[注意: 複数Excel一致] {len(multi_match_list)} 件")
        for src_name, excel_names in multi_match_list:
            logger.info(f"  - {src_name} → {', '.join(excel_names)}")

    if unmatched_excels:
        logger.info(f"\n[情報: 対応入力なしのExcel] {len(unmatched_excels)} 件")
        for path in sorted(unmatched_excels):
            logger.info(f"  - {os.path.basename(path)}")

    if backup_log:
        logger.info(f"\n[バックアップ] {len(backup_log)} 件作成済み")

    if move_log:
        logger.info(f"[ファイル移動] {len(move_log)} 件を processed/ へ移動")

    # レポートファイル生成
    report_name = f"report_{now.strftime('%Y%m%d')}.txt"
    report_path = os.path.join(excel_dir, report_name)
    generate_report(
        report_path, success_list, no_name_list, no_match_list,
        multi_match_list, unmatched_excels, backup_log, move_log,
        protected_log, mode=mode, dry_run=dry_run,
    )
    logger.info(f"\n[レポート] {report_path}")

    logger.info(f"\n{'='*60}")
    logger.info("処理完了")
    logger.info(f"{'='*60}")


def main():
    parser = argparse.ArgumentParser(
        description="利用者情報 → Excel ヒアリングシート 自動転記"
    )
    parser.add_argument("--config", default="config.json", help="設定ファイルパス")
    parser.add_argument("--source", default="pdf", choices=["pdf", "excel"],
                        help="入力ソース: pdf=PDF解析, excel=基本情報シート読取 (デフォルト: pdf)")
    parser.add_argument("--source-dir", default=None,
                        help="入力ソースディレクトリ (デフォルト: pdf→input_pdf, excel→config.source_dir)")
    parser.add_argument("--excel-dir", default="excel_sheets",
                        help="ヒアリングシート格納ディレクトリ")
    parser.add_argument("--form", default="form1",
                        help="書き込み先フォーム名 (デフォルト: form1)")
    parser.add_argument("--dry-run", action="store_true",
                        help="ドライラン（書き込みせず確認のみ）")
    parser.add_argument("--log-file", default=None, help="ログファイルパス")
    args = parser.parse_args()

    # 入力ディレクトリのデフォルト値をソースに応じて設定
    if args.source_dir is not None:
        input_dir = args.source_dir
    elif args.source == "pdf":
        input_dir = "input_pdf"
    else:
        # config の source_dir を参照（デフォルト: input_excel）
        config = load_config(args.config)
        input_dir = config.get("source_dir", "input_excel")

    process_all(
        config_path=args.config,
        input_dir=input_dir,
        excel_dir=args.excel_dir,
        form_name=args.form,
        mode=args.source,
        dry_run=args.dry_run,
        log_file=args.log_file,
    )


if __name__ == "__main__":
    main()
