#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
請求書・領収書 自動作成 Webアプリ
株式会社KB4
"""

import os
import datetime
import hashlib
import secrets
import urllib.parse
from flask import Flask, render_template, request, send_file, redirect, url_for, session
from create_document import create_receipt, create_invoice, OUTPUT_DIR

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

# 管理者パスワード（環境変数で設定可能）
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "kb4admin")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def get_output_files():
    """作成済みPDFファイル一覧を取得"""
    if not os.path.exists(OUTPUT_DIR):
        return []
    files = []
    for name in sorted(os.listdir(OUTPUT_DIR), reverse=True):
        if name.endswith(".pdf"):
            path = os.path.join(OUTPUT_DIR, name)
            mtime = os.path.getmtime(path)
            date_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y/%m/%d %H:%M")
            display = name.replace(".pdf", "")
            files.append({"name": name, "display": display, "date": date_str})
    return files


@app.route("/")
def index():
    today = datetime.date.today()
    default_number = today.strftime("%Y%m%d") + "-001"
    return render_template("index.html",
                           today=today.isoformat(),
                           default_number=default_number)


@app.route("/create", methods=["POST"])
def create():
    doc_type = request.form.get("doc_type", "invoice")
    recipient = request.form.get("recipient", "").strip()
    subject = request.form.get("subject", "商品代").strip()
    date_raw = request.form.get("date", "")
    number = request.form.get("number", "").strip()

    # 日付フォーマット変換（yyyy-mm-dd → yyyy年mm月dd日）
    if date_raw:
        try:
            d = datetime.datetime.strptime(date_raw, "%Y-%m-%d")
            date_str = d.strftime("%Y年%m月%d日")
        except ValueError:
            date_str = date_raw
    else:
        date_str = datetime.date.today().strftime("%Y年%m月%d日")

    if not number:
        number = datetime.date.today().strftime("%Y%m%d") + "-001"

    # 明細を取得
    names = request.form.getlist("item_name[]")
    qtys = request.form.getlist("item_qty[]")
    prices = request.form.getlist("item_price[]")
    reduced_indices = request.form.getlist("item_reduced[]")

    items = []
    for i in range(len(names)):
        if not names[i].strip():
            continue
        item = {
            "name": names[i].strip(),
            "qty": int(qtys[i]) if i < len(qtys) else 1,
            "price": int(prices[i]) if i < len(prices) else 0,
        }
        if doc_type == "receipt":
            item["reduced"] = str(i) in reduced_indices
        items.append(item)

    if not items:
        return redirect(url_for("index"))

    if doc_type == "receipt":
        has_reduced = any(item.get("reduced") for item in items)
        data = {
            "recipient": recipient,
            "subject": subject,
            "date": date_str,
            "number": number,
            "items": items,
            "reduced_tax": has_reduced,
        }
        order_number = request.form.get("order_number", "").strip()
        if order_number:
            data["order_number"] = order_number
        filepath = create_receipt(data)
    else:
        tax_rate_pct = int(request.form.get("tax_rate", "10"))
        data = {
            "recipient": recipient,
            "subject": subject,
            "date": date_str,
            "number": number,
            "items": items,
            "tax_rate": tax_rate_pct / 100,
        }
        due_date = request.form.get("due_date", "").strip()
        if due_date:
            try:
                d = datetime.datetime.strptime(due_date, "%Y-%m-%d")
                data["due_date"] = d.strftime("%Y年%m月%d日")
            except ValueError:
                data["due_date"] = due_date

        delivery = request.form.get("delivery_address", "").strip()
        if delivery:
            data["delivery_address"] = delivery

        bank = request.form.get("bank_info", "").strip()
        if bank:
            data["bank_info"] = bank

        filepath = create_invoice(data)

    filename = os.path.basename(filepath)
    doc_label = "領収書" if doc_type == "receipt" else "請求書"
    return redirect(url_for("done", message=f"{doc_label}を作成しました: {recipient}", file=filename))


@app.route("/done")
def done():
    today = datetime.date.today()
    default_number = today.strftime("%Y%m%d") + "-001"
    message = request.args.get("message", "")
    filename = request.args.get("file", "")
    download_url = url_for("download", filename=filename) if filename else None
    return render_template("index.html",
                           today=today.isoformat(),
                           default_number=default_number,
                           message=message,
                           download_url=download_url)


@app.route("/download/<path:filename>")
def download(filename):
    # ファイル名にパストラバーサルが含まれていないか確認
    if ".." in filename or filename.startswith("/"):
        return "不正なリクエストです", 400
    filepath = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    return "ファイルが見つかりません", 404


# --- 管理者ページ ---

@app.route("/admin", methods=["GET", "POST"])
def admin():
    if request.method == "POST" and not session.get("admin"):
        password = request.form.get("password", "")
        if password == ADMIN_PASSWORD:
            session["admin"] = True
            return redirect(url_for("admin"))
        else:
            return render_template("admin.html", logged_in=False, error="パスワードが違います")

    if session.get("admin"):
        return render_template("admin.html", logged_in=True, files=get_output_files())
    return render_template("admin.html", logged_in=False)


@app.route("/admin/delete", methods=["POST"])
def admin_delete():
    if not session.get("admin"):
        return redirect(url_for("admin"))
    filename = request.form.get("filename", "")
    if ".." in filename or filename.startswith("/"):
        return "不正なリクエストです", 400
    filepath = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(filepath):
        os.remove(filepath)
    return redirect(url_for("admin"))


@app.route("/admin/logout")
def admin_logout():
    session.pop("admin", None)
    return redirect(url_for("admin"))


os.makedirs(OUTPUT_DIR, exist_ok=True)

if __name__ == "__main__":
    print("\n  請求書・領収書 作成ツール")
    print("  http://localhost:5002\n")
    app.run(host="0.0.0.0", port=5002, debug=True)
