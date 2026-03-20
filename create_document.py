#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
請求書・領収書 自動作成ツール
株式会社KB4
"""

import os
import sys
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.utils import ImageReader

# --- 定数 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SEAL_IMAGE = os.path.join(SCRIPT_DIR, "電子印鑑.png")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")

# CIDフォント（macOS/Linux/Windows対応）
FONT_GOTHIC = "HeiseiKakuGo-W5"
FONT_MINCHO = "HeiseiMin-W3"

# 会社情報
COMPANY = {
    "name": "株式会社KB4",
    "zip": "〒5830016",
    "address": "大阪府藤井寺市陵南町2－21",
    "tel": "TEL: 09098802887",
    "email": "s.kubuki0826@gmail.com",
    "web": "https://shinkb4",
    "registration": "登録番号: T6120101064620",
}

# ページサイズ
PAGE_W, PAGE_H = A4  # 210mm x 297mm

# 印鑑サイズ（24mm = 2.4cm）
SEAL_SIZE = 24 * mm

# 透過処理済み印鑑キャッシュパス
SEAL_IMAGE_TRANSPARENT = os.path.join(SCRIPT_DIR, "電子印鑑_透過.png")


def register_fonts():
    """CIDフォントを登録"""
    pdfmetrics.registerFont(UnicodeCIDFont(FONT_GOTHIC))
    pdfmetrics.registerFont(UnicodeCIDFont(FONT_MINCHO))


def prepare_seal_image():
    """電子印鑑の白背景を透過処理する"""
    if os.path.exists(SEAL_IMAGE_TRANSPARENT):
        return SEAL_IMAGE_TRANSPARENT
    if not os.path.exists(SEAL_IMAGE):
        return None
    from PIL import Image
    img = Image.open(SEAL_IMAGE).convert("RGBA")
    pixels = img.load()
    w, h = img.size
    for y_px in range(h):
        for x_px in range(w):
            r, g, b, a = pixels[x_px, y_px]
            # 白～薄いグレーの背景を透明にする
            if r > 200 and g > 200 and b > 200:
                pixels[x_px, y_px] = (r, g, b, 0)
            # 薄い色（印影の周囲）を半透明にする
            elif r > 160 and g > 160 and b > 160:
                alpha = int(255 * (1 - min(r, g, b) / 255))
                pixels[x_px, y_px] = (r, g, b, alpha)
    img.save(SEAL_IMAGE_TRANSPARENT, "PNG")
    return SEAL_IMAGE_TRANSPARENT


def draw_seal(c, x, y):
    """電子印鑑を描画（指定座標の右上に配置、背景透過）"""
    seal_path = prepare_seal_image()
    if seal_path:
        img = ImageReader(seal_path)
        c.drawImage(img, x, y - SEAL_SIZE, width=SEAL_SIZE, height=SEAL_SIZE,
                     mask='auto', preserveAspectRatio=True)


def format_number(n):
    """数値をカンマ区切りでフォーマット"""
    return f"{n:,}"


def draw_header(c, doc_type, doc_date, doc_number):
    """ヘッダー部分（日付・番号・タイトル）を描画"""
    # 日付
    y = PAGE_H - 30 * mm
    c.setFont(FONT_GOTHIC, 10)
    c.drawRightString(PAGE_W - 20 * mm, y, doc_date)

    # 書類番号
    y -= 5 * mm
    prefix = "領収書番号" if doc_type == "領収書" else "請求書番号"
    c.drawRightString(PAGE_W - 20 * mm, y, f"{prefix}: {doc_number}")

    # タイトル
    y -= 18 * mm
    c.setFont(FONT_GOTHIC, 22)
    c.drawCentredString(PAGE_W / 2, y, doc_type)

    return y


def draw_recipient(c, y_start, recipient, subject):
    """宛先情報を描画"""
    x = 20 * mm
    y = y_start - 18 * mm

    c.setFont(FONT_GOTHIC, 14)
    # 宛先が長い場合は会社名と「御中」を分けて表示
    full_text = f"{recipient} 御中"
    max_width = PAGE_W / 2 - 15 * mm
    if c.stringWidth(full_text, FONT_GOTHIC, 14) > max_width:
        c.drawString(x, y, recipient)
        line_w = c.stringWidth(recipient, FONT_GOTHIC, 14)
        y -= 6 * mm
        c.drawString(x, y, "御中")
        line_w = max(line_w, c.stringWidth("御中", FONT_GOTHIC, 14))
    else:
        c.drawString(x, y, full_text)
        line_w = c.stringWidth(full_text, FONT_GOTHIC, 14)
    # 下線
    c.setLineWidth(0.8)
    c.line(x, y - 2, x + line_w, y - 2)

    y -= 10 * mm
    c.setFont(FONT_GOTHIC, 9)
    c.drawString(x, y, f"件名: {subject}")

    return y


def draw_company_info(c, y_start, doc_type):
    """会社情報と電子印鑑を描画"""
    x = PAGE_W / 2 + 10 * mm
    y = y_start - 18 * mm

    # 会社名
    c.setFont(FONT_GOTHIC, 12)
    company_name_width = c.stringWidth(COMPANY["name"], FONT_GOTHIC, 12)
    c.drawString(x, y, COMPANY["name"])

    # 電子印鑑を会社名の後ろに配置
    draw_seal(c, x + company_name_width + 2 * mm, y + 3 * mm)

    y -= 6 * mm
    c.setFont(FONT_GOTHIC, 8)
    c.drawString(x, y, COMPANY["zip"])
    y -= 4 * mm
    # 領収書の場合は短い住所、請求書は完全住所
    if doc_type == "領収書":
        c.drawString(x, y, "陵南町2－21")
    else:
        c.drawString(x, y, COMPANY["address"])
        y -= 4 * mm
        c.drawString(x, y, f"web: {COMPANY['web']}")

    y -= 5 * mm
    c.drawString(x, y, COMPANY["tel"])
    y -= 4 * mm
    c.drawString(x, y, COMPANY["email"])
    y -= 4 * mm
    c.drawString(x, y, COMPANY["registration"])

    return y


def draw_amount_summary(c, y, doc_type, total, tax_rate, tax_amount):
    """金額サマリー"""
    x = 20 * mm
    y -= 5 * mm

    if doc_type == "領収書":
        c.setFont(FONT_GOTHIC, 9)
        c.drawString(x, y, "下記のとおり領収いたしました。")
        y -= 10 * mm
        c.setFont(FONT_GOTHIC, 12)
        label = "領収金額"
    else:
        c.setFont(FONT_GOTHIC, 9)
        c.drawString(x, y, "下記のとおりご請求申し上げます。")
        y -= 10 * mm
        c.setFont(FONT_GOTHIC, 12)
        label = "ご請求金額"

    c.drawString(x, y, label)
    c.setFont(FONT_GOTHIC, 14)
    c.drawString(x + 35 * mm, y, f"¥ {format_number(total)} -")
    # 下線
    c.setLineWidth(0.8)
    c.line(x, y - 3, x + 80 * mm, y - 3)

    return y


def draw_items_table(c, y_start, items, tax_rate, tax_label, is_reduced=False):
    """明細テーブルを描画"""
    x_left = 20 * mm
    x_right = PAGE_W - 20 * mm
    table_width = x_right - x_left

    # カラム位置
    col_name_w = table_width * 0.48
    col_mark_w = table_width * 0.04
    col_qty_w = table_width * 0.12
    col_price_w = table_width * 0.16
    col_amount_w = table_width * 0.20

    col_x = [x_left,
             x_left + col_name_w,
             x_left + col_name_w + col_mark_w,
             x_left + col_name_w + col_mark_w + col_qty_w,
             x_left + col_name_w + col_mark_w + col_qty_w + col_price_w,
             x_right]

    y = y_start - 10 * mm
    row_h = 12 * mm  # 2行分の高さ
    header_h = 7 * mm

    # ヘッダー
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.rect(x_left, y - header_h, table_width, header_h, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont(FONT_GOTHIC, 8)

    header_y = y - header_h + 2 * mm
    c.drawCentredString((col_x[0] + col_x[1]) / 2, header_y, "品番・品名")
    c.drawCentredString((col_x[2] + col_x[3]) / 2, header_y, "数量")
    c.drawCentredString((col_x[3] + col_x[4]) / 2, header_y, "単価")
    c.drawCentredString((col_x[4] + col_x[5]) / 2, header_y, "金額")

    y -= header_h
    c.setFillColorRGB(0, 0, 0)

    # 明細行
    subtotal = 0
    for item in items:
        name = item.get("name", "")
        qty = item.get("qty", 0)
        price = item.get("price", 0)
        amount = qty * price
        subtotal += amount
        reduced = item.get("reduced", False)

        y -= row_h
        # 罫線
        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.setLineWidth(0.3)
        c.line(x_left, y, x_right, y)

        c.setFont(FONT_GOTHIC, 8)
        # 品名を2行で表示（折り返し）
        max_name_w = col_x[1] - col_x[0] - 4 * mm
        line1 = name
        line2 = ""
        if c.stringWidth(name, FONT_GOTHIC, 8) > max_name_w:
            # 1行目に収まる文字数を探す
            for i in range(len(name)):
                if c.stringWidth(name[:i+1], FONT_GOTHIC, 8) > max_name_w:
                    line1 = name[:i]
                    line2 = name[i:]
                    break
            # 2行目も溢れる場合は切り詰め
            if c.stringWidth(line2, FONT_GOTHIC, 8) > max_name_w:
                while c.stringWidth(line2, FONT_GOTHIC, 8) > max_name_w and len(line2) > 1:
                    line2 = line2[:-1]
                line2 = line2 + "..."

        item_y_top = y + 6 * mm  # 1行目（上寄り）
        item_y_btm = y + 2 * mm  # 2行目
        c.drawString(col_x[0] + 2 * mm, item_y_top, line1)
        if line2:
            c.drawString(col_x[0] + 2 * mm, item_y_btm, line2)

        # 数量・単価・金額は行の中央に配置
        item_y_mid = y + 4 * mm
        if reduced:
            c.drawCentredString((col_x[1] + col_x[2]) / 2, item_y_mid, "※")
        c.drawRightString(col_x[3] - 2 * mm, item_y_mid, str(qty))
        c.drawRightString(col_x[4] - 2 * mm, item_y_mid, format_number(price))
        c.drawRightString(col_x[5] - 2 * mm, item_y_mid, format_number(amount))

    # 空行を追加（最低8行表示）
    empty_rows = max(0, 8 - len(items))
    for _ in range(empty_rows):
        y -= row_h
        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.setLineWidth(0.3)
        c.line(x_left, y, x_right, y)

    # 最下線
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.5)
    c.line(x_left, y, x_right, y)

    # 軽減税率注記
    if is_reduced:
        y -= 5 * mm
        c.setFont(FONT_GOTHIC, 7)
        c.drawString(x_left, y, "※は軽減税率対象です。")

    # 税額計算
    tax_amount = int(subtotal * tax_rate / (1 + tax_rate)) if is_reduced else int(subtotal * tax_rate)
    total = subtotal if is_reduced else subtotal + int(subtotal * tax_rate)

    # 小計・税・合計
    summary_x = col_x[2]
    summary_right = x_right

    y -= 3 * mm
    summary_row_h = 7 * mm

    # 小計
    c.setStrokeColorRGB(0.5, 0.5, 0.5)
    c.setLineWidth(0.3)
    y -= summary_row_h
    c.line(summary_x, y, summary_right, y)
    c.setFont(FONT_GOTHIC, 9)
    c.drawString(summary_x + 2 * mm, y + 2 * mm, "小計")
    c.drawRightString(summary_right - 2 * mm, y + 2 * mm, format_number(subtotal))

    # 消費税
    y -= summary_row_h
    c.line(summary_x, y, summary_right, y)
    c.drawString(summary_x + 2 * mm, y + 2 * mm, tax_label)
    tax_display = f"({format_number(tax_amount)})" if is_reduced else format_number(int(subtotal * tax_rate))
    c.drawRightString(summary_right - 2 * mm, y + 2 * mm, tax_display)

    # 合計
    y -= summary_row_h
    c.setLineWidth(0.8)
    c.setStrokeColorRGB(0, 0, 0)
    c.line(summary_x, y, summary_right, y)
    c.setFont(FONT_GOTHIC, 10)
    c.drawString(summary_x + 2 * mm, y + 2 * mm, "合計")
    c.drawRightString(summary_right - 2 * mm, y + 2 * mm, format_number(total))

    return y, subtotal, tax_amount, total


def create_receipt(data):
    """領収書PDFを作成"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename = f"{data['recipient']}　領収書　御中.pdf"
    filepath = os.path.join(OUTPUT_DIR, filename)

    c = canvas.Canvas(filepath, pagesize=A4)
    register_fonts()

    # ヘッダー
    y = draw_header(c, "領収書", data["date"], data["number"])

    # 宛先
    y_left = draw_recipient(c, y, data["recipient"], data["subject"])

    # 会社情報（右側）
    draw_company_info(c, y, "領収書")

    # 金額サマリー
    is_reduced = data.get("reduced_tax", False)
    tax_rate = 0.08 if is_reduced else 0.10

    # 先に合計を計算
    subtotal = sum(item["qty"] * item["price"] for item in data["items"])
    if is_reduced:
        tax_amount = int(subtotal * tax_rate / (1 + tax_rate))
        total = subtotal
    else:
        tax_amount = int(subtotal * tax_rate)
        total = subtotal + tax_amount

    y_amount = draw_amount_summary(c, y_left, "領収書", total, tax_rate, tax_amount)

    # 明細テーブル
    tax_label = f"消費税 (軽減{int(tax_rate*100)}% 内税)" if is_reduced else f"消費税({int(tax_rate*100)}%)"
    y_table, _, _, _ = draw_items_table(c, y_amount, data["items"], tax_rate, tax_label, is_reduced)

    # 注文番号
    if data.get("order_number"):
        y_table -= 10 * mm
        c.setFont(FONT_GOTHIC, 9)
        c.drawString(20 * mm, y_table, f"注文番号　{data['order_number']}")

    c.save()
    return filepath


def create_invoice(data):
    """請求書PDFを作成"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename = f"{data['recipient']}請求書.pdf"
    filepath = os.path.join(OUTPUT_DIR, filename)

    c = canvas.Canvas(filepath, pagesize=A4)
    register_fonts()

    # ヘッダー
    y = draw_header(c, "請求書", data["date"], data["number"])

    # 宛先
    y_left = draw_recipient(c, y, data["recipient"], data["subject"])

    # 会社情報（右側）
    draw_company_info(c, y, "請求書")

    # 金額サマリー
    tax_rate = data.get("tax_rate", 0.10)
    subtotal = sum(item["qty"] * item["price"] for item in data["items"])
    tax_amount = int(subtotal * tax_rate)
    total = subtotal + tax_amount

    y_amount = draw_amount_summary(c, y_left, "請求書", total, tax_rate, tax_amount)

    # 支払い期限
    if data.get("due_date"):
        y_amount -= 6 * mm
        c.setFont(FONT_GOTHIC, 8)
        c.drawString(20 * mm, y_amount, f"お支払い期限: {data['due_date']}")

    # 明細テーブル
    tax_label = f"消費税({int(tax_rate*100)}%)"
    y_table, _, _, _ = draw_items_table(c, y_amount, data["items"], tax_rate, tax_label)

    # 納品先
    if data.get("delivery_address"):
        y_table -= 10 * mm
        c.setFont(FONT_GOTHIC, 8)
        c.drawString(20 * mm, y_table, "納品先:")
        y_table -= 4 * mm
        c.drawString(20 * mm, y_table, data["delivery_address"])

    # 振込先
    if data.get("bank_info"):
        y_table -= 8 * mm
        c.setFont(FONT_GOTHIC, 8)
        c.drawString(20 * mm, y_table, "お振込先:")
        y_table -= 4 * mm
        c.drawString(20 * mm, y_table, data["bank_info"])

    c.save()
    return filepath


def interactive_mode():
    """対話モードで書類を作成"""
    print("=" * 50)
    print("  請求書・領収書 自動作成ツール")
    print("  株式会社KB4")
    print("=" * 50)
    print()
    print("作成する書類を選択してください:")
    print("  1. 請求書")
    print("  2. 領収書")
    print("  q. 終了")
    print()

    choice = input("選択 (1/2/q): ").strip()
    if choice == "q":
        print("終了します。")
        return

    if choice not in ("1", "2"):
        print("無効な選択です。")
        return

    doc_type = "請求書" if choice == "1" else "領収書"
    today = datetime.date.today()

    print(f"\n--- {doc_type}情報を入力 ---\n")

    # 宛先
    recipient = input("宛先（会社名）: ").strip()
    if not recipient:
        print("宛先は必須です。")
        return

    # 件名
    subject = input("件名: ").strip() or "商品代"

    # 日付
    date_str = input(f"日付 (Enter で今日 {today.strftime('%Y年%m月%d日')}): ").strip()
    if not date_str:
        date_str = today.strftime("%Y年%m月%d日")

    # 書類番号
    default_num = today.strftime("%Y%m%d") + "-001"
    doc_number = input(f"書類番号 (Enter で {default_num}): ").strip() or default_num

    # 明細入力
    print("\n--- 明細を入力（空行で終了） ---")
    items = []
    idx = 1
    while True:
        name = input(f"\n  品名 [{idx}] (空でEnterで終了): ").strip()
        if not name:
            if not items:
                print("  最低1つの明細が必要です。")
                continue
            break
        try:
            qty = int(input(f"  数量 [{idx}]: ").strip())
            price = int(input(f"  単価 [{idx}]: ").strip())
        except ValueError:
            print("  数値を入力してください。")
            continue

        item = {"name": name, "qty": qty, "price": price}

        if doc_type == "領収書":
            reduced = input(f"  軽減税率対象？ (y/N): ").strip().lower()
            item["reduced"] = reduced == "y"

        items.append(item)
        amount = qty * price
        print(f"  → {name}: {qty} × {format_number(price)} = ¥{format_number(amount)}")
        idx += 1

    data = {
        "recipient": recipient,
        "subject": subject,
        "date": date_str,
        "number": doc_number,
        "items": items,
    }

    if doc_type == "領収書":
        # 軽減税率チェック
        has_reduced = any(item.get("reduced") for item in items)
        data["reduced_tax"] = has_reduced

        order_num = input("\n注文番号 (任意): ").strip()
        if order_num:
            data["order_number"] = order_num

        filepath = create_receipt(data)
    else:
        # 請求書固有情報
        tax_input = input("\n消費税率 (Enter で10%): ").strip()
        if tax_input:
            data["tax_rate"] = int(tax_input) / 100
        else:
            data["tax_rate"] = 0.10

        due = input("支払い期限 (任意): ").strip()
        if due:
            data["due_date"] = due

        delivery = input("納品先 (任意): ").strip()
        if delivery:
            data["delivery_address"] = delivery

        bank = input("振込先 (任意、Enter でデフォルト): ").strip()
        if not bank:
            bank = "住信SBIネット銀行 法人第一支店（106）普通 1833480 株式会社KB4"
        data["bank_info"] = bank

        filepath = create_invoice(data)

    print(f"\n{'='*50}")
    print(f"  {doc_type}を作成しました！")
    print(f"  ファイル: {filepath}")
    print(f"{'='*50}")

    # macOSでファイルを開く
    open_file = input("\nファイルを開きますか？ (Y/n): ").strip().lower()
    if open_file != "n":
        os.system(f'open "{filepath}"')


if __name__ == "__main__":
    interactive_mode()
