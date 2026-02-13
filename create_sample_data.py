#!/usr/bin/env python3
"""サンプルのPDFファイルとExcelファイルを生成するスクリプト。
開発・テスト用途。"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# reportlab で日本語フォント登録
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))
FONT_NAME = "HeiseiMin-W3"


def create_sample_excel(filepath, facility_name, person_name):
    """ヒアリングシート風のExcelファイルを作成する。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "ヒアリングシート"

    # スタイル定義
    header_font = Font(name="游ゴシック", size=14, bold=True)
    label_font = Font(name="游ゴシック", size=11, bold=True)
    input_font = Font(name="游ゴシック", size=11)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    label_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # タイトル
    ws.merge_cells("A1:F1")
    ws["A1"] = f"{facility_name}　無料体験ヒアリングシート"
    ws["A1"].font = header_font
    ws["A1"].alignment = Alignment(horizontal="center")

    # ラベルと入力欄のペア
    labels = [
        (3, "A", "名前"),
        (4, "A", "ふりがな"),
        (5, "A", "性別"),
        (6, "A", "年齢"),
        (7, "A", "生年月日"),
        (8, "A", "住所"),
        (9, "A", "電話番号"),
        (10, "A", "要介護認定"),
    ]

    # 列幅設定
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 10

    for row, col, label_text in labels:
        cell = ws[f"{col}{row}"]
        cell.value = label_text
        cell.font = label_font
        cell.border = thin_border
        cell.fill = label_fill
        cell.alignment = Alignment(vertical="center")

        # 入力セル (B列)
        input_cell = ws[f"B{row}"]
        input_cell.border = thin_border
        input_cell.font = input_font
        input_cell.alignment = Alignment(vertical="center")

    # 要介護認定の特別レイアウト: C10〜F10 に選択肢
    care_levels = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"]
    # B10 をメイン入力欄として使い、下にも選択肢を配置
    ws["B10"].value = ""
    for i, level in enumerate(care_levels):
        row_offset = 11 + i
        ws[f"A{row_offset}"] = ""
        ws[f"B{row_offset}"] = level
        ws[f"B{row_offset}"].font = input_font
        ws[f"B{row_offset}"].border = thin_border
        ws[f"C{row_offset}"] = ""  # チェック欄
        ws[f"C{row_offset}"].border = thin_border

    # 追加項目
    extra_start = 11 + len(care_levels) + 1
    extra_labels = [
        (extra_start, "A", "既往歴"),
        (extra_start + 1, "A", "服薬情報"),
        (extra_start + 2, "A", "備考"),
    ]
    for row, col, label_text in extra_labels:
        cell = ws[f"{col}{row}"]
        cell.value = label_text
        cell.font = label_font
        cell.border = thin_border
        cell.fill = label_fill
        input_cell = ws[f"B{row}"]
        input_cell.border = thin_border
        input_cell.font = input_font

    wb.save(filepath)
    print(f"  [Excel] 作成: {filepath}")


def create_sample_pdf(filepath, data):
    """利用者情報のサンプルPDFを作成する。"""
    c = canvas.Canvas(filepath, pagesize=A4)
    width, height = A4

    c.setFont(FONT_NAME, 16)
    c.drawString(40 * mm, height - 30 * mm, "利用者情報")

    c.setFont(FONT_NAME, 11)
    y = height - 50 * mm
    line_height = 10 * mm

    fields = [
        ("氏名", data["氏名"]),
        ("ふりがな", data["ふりがな"]),
        ("性別", data["性別"]),
        ("生年月日", data["生年月日"]),
        ("年齢", data["年齢"]),
        ("住所", data["住所"]),
        ("電話番号", data["電話番号"]),
        ("要介護度", data["要介護度"]),
    ]

    for label, value in fields:
        c.drawString(30 * mm, y, f"{label}：{value}")
        y -= line_height

    c.save()
    print(f"  [PDF]   作成: {filepath}")


def _fill_source_excel_sheet(ws, data):
    """1シート分の基本情報シートデータを書き込む（新仕様: B列始まり）。

    セル配置（個別シート用、B列始まり）:
        B1:T1 タイトル,
        E6:T7 住所, E8:O8 ふりがな, E9:O9 氏名,
        R8:T9 性別テンプレ, E10:N11 生年月日テンプレ, O10:T10 年齢テンプレ,
        F13:T14 現在の病気, F15:T16 既往歴, F18:T19 服薬情報,
        F21:O21 医療機関名, P21:T21 主治医,
        C27:F28 緊急連絡先1氏名, G27:H28 続柄, I27:O28 住所, P27:T28 電話
    """
    header_font = Font(name="游ゴシック", size=14, bold=True)
    label_font = Font(name="游ゴシック", size=11, bold=True)
    input_font = Font(name="游ゴシック", size=11)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    label_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    # 列幅
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
                "L", "M", "N", "O", "P", "Q", "R", "S", "T"]:
        ws.column_dimensions[col].width = 6

    # タイトル
    ws.merge_cells("B1:T1")
    ws["B1"] = "利用者基本情報シート"
    ws["B1"].font = header_font
    ws["B1"].alignment = Alignment(horizontal="center")

    # ラベル＋データ配置
    def _label(cell, text):
        ws[cell] = text
        ws[cell].font = label_font
        ws[cell].border = thin_border
        ws[cell].fill = label_fill

    def _value(cell, text):
        ws[cell] = text
        ws[cell].font = input_font
        ws[cell].border = thin_border

    # 住所 (E6)
    _label("B6", "住所")
    _value("E6", data.get("住所", ""))

    # ふりがな (E8)
    _label("B8", "ふりがな")
    _value("E8", data.get("ふりがな", ""))

    # 氏名 (E9)
    _label("B9", "氏名")
    _value("E9", data.get("氏名", ""))

    # 性別テンプレート (R8) - 常に「男　・　女」テンプレ
    _label("R8", data.get("性別テンプレ", "男　・　女"))

    # 生年月日テンプレ (E10) - 元号テンプレ + 実年月日
    era_year = data.get("和暦年", "")
    era_month = data.get("月", "")
    era_day = data.get("日", "")
    dob_template = f"明治・大正・昭和・平成　　　　{era_year}年　　{era_month}月　　{era_day}日"
    _label("B10", "生年月日")
    _value("E10", dob_template if era_year else "明治・大正・昭和・平成　　　　年　　月　　日")

    # 年齢テンプレ (O10)
    age_num = data.get("年齢数値", "")
    _value("O10", f"年齢　　　　　　{age_num}　　歳" if age_num else "年齢　　　　　　　　歳")

    # 現在の病気 (F13)
    _label("B13", "現在の病気")
    _value("F13", data.get("現在の病気", ""))

    # 既往歴 (F15)
    _label("B15", "既往歴")
    _value("F15", data.get("既往歴", ""))

    # 服薬情報 (F18)
    _label("B18", "服薬情報")
    _value("F18", data.get("服薬情報", ""))

    # 医療機関名 (F21) / 主治医 (P21)
    _label("B21", "医療機関名")
    _value("F21", data.get("医療機関名", ""))
    _value("P21", data.get("主治医", ""))

    # 緊急連絡先 (C27, G27, I27, P27)
    _label("B27", "緊急連絡先")
    _value("C27", data.get("緊急連絡先1_氏名", ""))
    _value("G27", data.get("緊急連絡先1_続柄", ""))
    _value("I27", data.get("緊急連絡先1_住所", ""))
    _value("P27", data.get("緊急連絡先1_電話", ""))


def create_source_excel(filepath, users_data):
    """基本情報シートExcelを作成する（原本 + 個別シート + ダミー）。

    シート構成:
      - 「原本」: A列始まりテンプレート（氏名空欄）
      - 各利用者名のシート: B列始まり実データ
      - 「空シート」: 氏名が空のダミーシート
    """
    wb = Workbook()

    # 1.「原本」シート（テンプレート）
    ws_template = wb.active
    ws_template.title = "原本"
    _fill_source_excel_sheet(ws_template, {})  # 全空

    # 2. 各利用者のデータシート
    for user_data in users_data:
        # シート名は「氏名様」形式
        name = user_data.get("氏名", "不明")
        sheet_name = _clean_sheet_name(name) + "様"
        ws = wb.create_sheet(title=sheet_name)
        _fill_source_excel_sheet(ws, user_data)

    # 3. ダミーシート（氏名空欄）
    ws_empty = wb.create_sheet(title="空シート")
    _fill_source_excel_sheet(ws_empty, {})

    wb.save(filepath)
    print(f"  [基本情報シート] 作成: {filepath} ({len(wb.sheetnames)}シート: {', '.join(wb.sheetnames)})")


def _clean_sheet_name(name):
    """シート名に使える形に整形する。"""
    return name.replace("　", "").replace(" ", "")


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_dir = os.path.join(base_dir, "input_pdf")
    excel_dir = os.path.join(base_dir, "excel_sheets")
    basic_info_dir = os.path.join(base_dir, "input_excel")
    os.makedirs(pdf_dir, exist_ok=True)
    os.makedirs(excel_dir, exist_ok=True)
    os.makedirs(basic_info_dir, exist_ok=True)

    # サンプルデータ（PDF・ヒアリングシート用）
    sample_users = [
        {
            "施設名略称": "パシ",
            "施設名正式": "パシフィック",
            "氏名": "三代　花子",
            "苗字": "三代",
            "ふりがな": "みしろ はなこ",
            "性別": "女性",
            "生年月日": "昭和15年3月10日",
            "年齢": "85歳",
            "住所": "岡山県笠岡市中央町1-2-3",
            "電話番号": "0865-12-3456",
            "要介護度": "要介護2",
        },
        {
            "施設名略称": "ルネ",
            "施設名正式": "ルネサンス",
            "氏名": "笠岡　太郎",
            "苗字": "笠岡",
            "ふりがな": "かさおか たろう",
            "性別": "男性",
            "生年月日": "昭和10年7月20日",
            "年齢": "90歳",
            "住所": "岡山県笠岡市港町4-5-6",
            "電話番号": "0865-78-9012",
            "要介護度": "要介護3",
        },
        {
            "施設名略称": "パシ",
            "施設名正式": "パシフィック",
            "氏名": "山田　次郎",
            "苗字": "山田",
            "ふりがな": "やまだ じろう",
            "性別": "男性",
            "生年月日": "昭和20年1月15日",
            "年齢": "80歳",
            "住所": "岡山県倉敷市本町7-8-9",
            "電話番号": "086-123-4567",
            "要介護度": "要支援1",
        },
    ]

    # 基本情報シートExcel用データ（新仕様: B列始まり）
    source_excel_users = {
        "パシ": [
            {
                "氏名": "杉田　誠一",
                "ふりがな": "すぎた　せいいち",
                "住所": "大阪府大阪市西淀川区野里1-32-14-108",
                "和暦年": "32", "月": "4", "日": "2",
                "年齢数値": "68",
                "現在の病気": "家族性パーキンソン病Yahr5",
                "既往歴": "肺炎治療後、神経症、腰痛症",
                "服薬情報": "ドパコール配合錠L150mg",
                "医療機関名": "田島クリニック",
                "主治医": "谷川　祐二",
                "緊急連絡先1_氏名": "杉田　孝次郎",
                "緊急連絡先1_続柄": "弟",
                "緊急連絡先1_住所": "大阪府大阪市鶴見区放出東3-21-61",
                "緊急連絡先1_電話": "090-6056-2930",
            },
            {
                "氏名": "林　郁子",
                "ふりがな": "はやし　いくこ",
                "住所": "大阪府大阪市西淀川区姫島5-12-3",
                "和暦年": "28", "月": "11", "日": "15",
                "年齢数値": "72",
                "現在の病気": "変形性膝関節症",
                "既往歴": "高血圧、糖尿病",
                "服薬情報": "アムロジピン5mg",
                "医療機関名": "西淀川医院",
                "主治医": "佐々木　健一",
                "緊急連絡先1_氏名": "林　正夫",
                "緊急連絡先1_続柄": "夫",
                "緊急連絡先1_住所": "大阪府大阪市西淀川区姫島5-12-3",
                "緊急連絡先1_電話": "06-1234-5678",
            },
        ],
        "ルネ": [
            {
                "氏名": "笠岡　太郎",
                "ふりがな": "かさおか　たろう",
                "住所": "岡山県笠岡市港町4-5-6",
                "和暦年": "10", "月": "7", "日": "20",
                "年齢数値": "90",
                "現在の病気": "慢性心不全",
                "既往歴": "脳梗塞、前立腺肥大",
                "服薬情報": "ワーファリン2mg、タムスロシン0.2mg",
                "医療機関名": "笠岡中央病院",
                "主治医": "田中　一郎",
                "緊急連絡先1_氏名": "笠岡　良子",
                "緊急連絡先1_続柄": "妻",
                "緊急連絡先1_住所": "岡山県笠岡市港町4-5-6",
                "緊急連絡先1_電話": "0865-78-9012",
            },
        ],
    }

    print("=== サンプルデータ生成 ===\n")

    for user in sample_users:
        # ヒアリングシート Excel 生成
        excel_name = f"{user['施設名略称']}{user['苗字']}さん　無料体験ヒアリングシート.xlsx"
        excel_path = os.path.join(excel_dir, excel_name)
        create_sample_excel(excel_path, user["施設名正式"], user["苗字"])

        # PDF ファイル生成
        pdf_name = f"利用者情報_{user['氏名'].replace('　', '')}.pdf"
        pdf_path = os.path.join(pdf_dir, pdf_name)
        create_sample_pdf(pdf_path, user)

    # 照合失敗テスト用: ヒアリングシートに対応するPDFがないケース
    excel_no_match = os.path.join(excel_dir, "パシ田中さん　無料体験ヒアリングシート.xlsx")
    create_sample_excel(excel_no_match, "パシフィック", "田中")

    # 基本情報シートExcel生成（新仕様: 施設ごとに1ファイル、原本+個別シート+ダミー）
    for facility_abbr, users in source_excel_users.items():
        bi_name = f"基本情報シート_{facility_abbr}_.xlsx"
        bi_path = os.path.join(basic_info_dir, bi_name)
        create_source_excel(bi_path, users)

    # 基本情報シートに対応するヒアリングシートも追加生成
    extra_hearing_sheets = [
        ("パシ", "杉田"), ("パシ", "林"),
    ]
    for abbr, surname in extra_hearing_sheets:
        hs_name = f"{abbr}{surname}さん　無料体験ヒアリングシート.xlsx"
        hs_path = os.path.join(excel_dir, hs_name)
        if not os.path.exists(hs_path):
            create_sample_excel(hs_path, {"パシ": "パシフィック", "ルネ": "ルネサンス"}[abbr], surname)

    print("\n=== 生成完了 ===")
    print(f"PDF格納先: {pdf_dir}")
    print(f"基本情報シート格納先: {basic_info_dir}")
    print(f"ヒアリングシート格納先: {excel_dir}")


if __name__ == "__main__":
    main()
