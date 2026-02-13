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


def _fill_basic_info_sheet(ws, data):
    """1シート分の基本情報データを書き込む（共通ヘルパー）。"""
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
    section_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 4
    ws.column_dimensions["C"].width = 35

    # タイトル
    ws.merge_cells("A1:C1")
    ws["A1"] = "利用者基本情報"
    ws["A1"].font = header_font
    ws["A1"].alignment = Alignment(horizontal="center")

    # 基本情報セクション
    ws["A2"] = "【基本情報】"
    ws["A2"].font = label_font
    ws["A2"].fill = section_fill

    rows = [
        (3, "ふりがな", data.get("ふりがな", "")),
        (4, "氏名", data.get("氏名", "")),
        (5, "生年月日", data.get("生年月日", "")),
        (6, "性別", data.get("性別", "")),
        (7, "住所", data.get("住所", "")),
        (8, "電話番号", data.get("電話番号", "")),
    ]
    for row, label, value in rows:
        ws[f"A{row}"] = label
        ws[f"A{row}"].font = label_font
        ws[f"A{row}"].border = thin_border
        ws[f"A{row}"].fill = label_fill
        ws[f"C{row}"] = value
        ws[f"C{row}"].font = input_font
        ws[f"C{row}"].border = thin_border

    # 介護保険セクション
    ws["A10"] = "【介護保険】"
    ws["A10"].font = label_font
    ws["A10"].fill = section_fill
    ws["A11"] = "要介護度"
    ws["A11"].font = label_font
    ws["A11"].border = thin_border
    ws["A11"].fill = label_fill
    ws["C11"] = data.get("要介護度", "")
    ws["C11"].font = input_font
    ws["C11"].border = thin_border


def create_basic_info_sheet(filepath, data):
    """基本情報シート風のExcelファイルを作成する（単一シート版）。

    セル配置:
        A1: タイトル, A3: ふりがなラベル, C3: ふりがな値
        A4: 氏名ラベル, C4: 氏名値, A5: 生年月日ラベル, C5: 生年月日値
        A7: 住所ラベル, C7: 住所値, A8: 電話番号ラベル, C8: 電話番号値
        A10: 介護保険セクション, A11: 要介護度ラベル, C11: 要介護度値
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "基本情報"
    _fill_basic_info_sheet(ws, data)
    wb.save(filepath)
    print(f"  [基本情報] 作成: {filepath}")


def create_multi_sheet_basic_info(filepath, users_data):
    """マルチシート基本情報Excelを作成する（テスト用）。

    シート構成:
      - 「原本」: テンプレートシート（ダミーデータ「年  月  日」）
      - 各利用者名のシート: 実データ
      - 「空シート」: C4が空欄のダミーシート
    """
    wb = Workbook()

    # 1. 「原本」シート（テンプレート）
    ws_template = wb.active
    ws_template.title = "原本"
    _fill_basic_info_sheet(ws_template, {
        "ふりがな": "",
        "氏名": "年  月  日",
        "生年月日": "年  月  日",
        "性別": "",
        "住所": "",
        "電話番号": "",
        "要介護度": "",
    })

    # 2. 各利用者のデータシート
    for user_data in users_data:
        sheet_name = user_data.get("氏名", "不明").replace("　", " ")
        ws = wb.create_sheet(title=sheet_name)
        _fill_basic_info_sheet(ws, user_data)

    # 3. ダミーシート（C4が空欄）
    ws_empty = wb.create_sheet(title="空シート")
    _fill_basic_info_sheet(ws_empty, {
        "ふりがな": "",
        "氏名": "",
        "生年月日": "",
        "性別": "",
        "住所": "",
        "電話番号": "",
        "要介護度": "",
    })

    wb.save(filepath)
    print(f"  [マルチシート基本情報] 作成: {filepath} ({len(wb.sheetnames)} シート: {', '.join(wb.sheetnames)})")


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_dir = os.path.join(base_dir, "input_pdf")
    excel_dir = os.path.join(base_dir, "excel_sheets")
    basic_info_dir = os.path.join(base_dir, "input_excel")
    os.makedirs(pdf_dir, exist_ok=True)
    os.makedirs(excel_dir, exist_ok=True)
    os.makedirs(basic_info_dir, exist_ok=True)

    # サンプルデータ
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

    print("=== サンプルデータ生成 ===\n")

    for user in sample_users:
        # Excel ファイル生成
        excel_name = f"{user['施設名略称']}{user['苗字']}さん　無料体験ヒアリングシート.xlsx"
        excel_path = os.path.join(excel_dir, excel_name)
        create_sample_excel(excel_path, user["施設名正式"], user["苗字"])

        # PDF ファイル生成
        pdf_name = f"利用者情報_{user['氏名'].replace('　', '')}.pdf"
        pdf_path = os.path.join(pdf_dir, pdf_name)
        create_sample_pdf(pdf_path, user)

        # 基本情報シート生成
        bi_name = f"基本情報シート（{user['施設名略称']}）_{user['苗字']}.xlsx"
        bi_path = os.path.join(basic_info_dir, bi_name)
        create_basic_info_sheet(bi_path, user)

    # 照合失敗テスト用: Excelに対応するPDFがないケース
    excel_no_match = os.path.join(excel_dir, "パシ田中さん　無料体験ヒアリングシート.xlsx")
    create_sample_excel(excel_no_match, "パシフィック", "田中")

    # マルチシート基本情報テスト用（原本+実データ+ダミー）
    # パシフィック施設の2名分を1ファイルにまとめる
    pasi_users = [u for u in sample_users if u["施設名略称"] == "パシ"]
    multi_bi_path = os.path.join(basic_info_dir, "基本情報シート（パシ）_まとめ.xlsx")
    create_multi_sheet_basic_info(multi_bi_path, pasi_users)

    # 「様」付き氏名テスト用
    sama_user = {
        "氏名": "佐藤　花子　様",
        "ふりがな": "さとう はなこ",
        "性別": "女性",
        "生年月日": "昭和18年5月1日",
        "住所": "岡山県笠岡市中央町10-11",
        "電話番号": "0865-99-8765",
        "要介護度": "要介護1",
    }
    sama_bi_path = os.path.join(basic_info_dir, "基本情報シート（パシ）_佐藤.xlsx")
    create_basic_info_sheet(sama_bi_path, sama_user)
    # 対応するヒアリングシートも作成
    sama_excel = os.path.join(excel_dir, "パシ佐藤さん　無料体験ヒアリングシート.xlsx")
    create_sample_excel(sama_excel, "パシフィック", "佐藤")

    print("\n=== 生成完了 ===")
    print(f"PDF格納先: {pdf_dir}")
    print(f"基本情報シート格納先: {basic_info_dir}")
    print(f"ヒアリングシート格納先: {excel_dir}")


if __name__ == "__main__":
    main()
