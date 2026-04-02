#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
請求書・領収書 自動作成 Webアプリ
株式会社KB4
"""

import os
import io
import datetime
from flask import Flask, render_template, request, send_file, redirect, url_for, session
from create_document import create_receipt, create_invoice, OUTPUT_DIR

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "704dbadcceb1112e9688e4f03b0148924bb2adfc7e824b55eb1a9c4aa5b8530a")

# 管理者パスワード（環境変数で設定可能）
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "kb4admin")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ---- データベース設定 ----
DATABASE_URL = os.environ.get("DATABASE_URL")


def get_db():
    """PostgreSQL接続を返す。DATABASE_URLがなければNone"""
    if not DATABASE_URL:
        return None
    try:
        import psycopg2
        # Render の DATABASE_URL は postgres:// を使うことがある → postgresql:// に変換
        url = DATABASE_URL
        if url.startswith("postgres://"):
            url = url.replace("postgres://", "postgresql://", 1)
        return psycopg2.connect(url)
    except Exception as e:
        print(f"DB接続エラー: {e}")
        return None


def init_db():
    """起動時にテーブルを作成（存在しなければ）"""
    conn = get_db()
    if not conn:
        return
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS invoices (
                        id SERIAL PRIMARY KEY,
                        created_at TIMESTAMP DEFAULT NOW(),
                        doc_type VARCHAR(20),
                        recipient TEXT,
                        subject TEXT,
                        date_str TEXT,
                        number TEXT,
                        total_amount INTEGER,
                        filename TEXT,
                        pdf_data BYTEA
                    )
                """)
    finally:
        conn.close()


def save_invoice_to_db(doc_type, recipient, subject, date_str, number, total_amount, filename, pdf_bytes):
    """請求書/領収書データをDBに保存"""
    conn = get_db()
    if not conn:
        return
    try:
        import psycopg2.extras
        with conn:
            with conn.cursor() as cur:
                cur.execute("""
                    INSERT INTO invoices (doc_type, recipient, subject, date_str, number, total_amount, filename, pdf_data)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """, (doc_type, recipient, subject, date_str, number, total_amount, filename, psycopg2.Binary(pdf_bytes)))
    finally:
        conn.close()


def get_invoices_from_db():
    """DB から請求書一覧を取得"""
    conn = get_db()
    if not conn:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT id, created_at, doc_type, recipient, subject, date_str, number, total_amount, filename
                FROM invoices
                ORDER BY created_at DESC
            """)
            rows = cur.fetchall()
        return [
            {
                "id": r[0],
                "date": r[1].strftime("%Y/%m/%d %H:%M") if r[1] else "",
                "doc_type": r[2],
                "doc_label": "領収書" if r[2] == "receipt" else "請求書",
                "recipient": r[3],
                "subject": r[4],
                "date_str": r[5],
                "number": r[6],
                "total_amount": r[7],
                "filename": r[8],
            }
            for r in rows
        ]
    finally:
        conn.close()


def get_pdf_from_db(invoice_id):
    """DB から PDF バイナリを取得"""
    conn = get_db()
    if not conn:
        return None, None
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT filename, pdf_data FROM invoices WHERE id = %s", (invoice_id,))
            row = cur.fetchone()
        if row:
            return row[0], bytes(row[1])
        return None, None
    finally:
        conn.close()


def delete_invoice_from_db(invoice_id):
    """DB からレコードを削除"""
    conn = get_db()
    if not conn:
        return
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute("DELETE FROM invoices WHERE id = %s", (invoice_id,))
    finally:
        conn.close()


# ---- ファイルシステム fallback ----

def get_output_files():
    """作成済みPDFファイル一覧をファイルシステムから取得（DB未使用時の fallback）"""
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


# ---- ルート ----

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
        total_amount = sum(item["qty"] * item["price"] for item in items)
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
        subtotal = sum(item["qty"] * item["price"] for item in items)
        total_amount = int(subtotal * (1 + tax_rate_pct / 100))

    filename = os.path.basename(filepath)

    # DBが使える場合はPDFバイナリをDBに保存
    if DATABASE_URL:
        try:
            with open(filepath, "rb") as f:
                pdf_bytes = f.read()
            save_invoice_to_db(doc_type, recipient, subject, date_str, number, total_amount, filename, pdf_bytes)
        except Exception as e:
            print(f"DB保存エラー: {e}")

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

    # まずファイルシステムから探す
    filepath = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)

    # ファイルがなければDBから探す
    if DATABASE_URL:
        conn = get_db()
        if conn:
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT pdf_data FROM invoices WHERE filename = %s ORDER BY created_at DESC LIMIT 1", (filename,))
                    row = cur.fetchone()
                if row:
                    pdf_bytes = bytes(row[0])
                    return send_file(
                        io.BytesIO(pdf_bytes),
                        as_attachment=True,
                        download_name=filename,
                        mimetype="application/pdf"
                    )
            finally:
                conn.close()

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
        use_db = DATABASE_URL is not None
        if use_db:
            invoices = get_invoices_from_db()
            if invoices is None:
                invoices = []
            return render_template("admin.html", logged_in=True, invoices=invoices, use_db=True)
        else:
            files = get_output_files()
            return render_template("admin.html", logged_in=True, files=files, use_db=False)

    return render_template("admin.html", logged_in=False)


@app.route("/admin/download/<int:invoice_id>")
def admin_download(invoice_id):
    if not session.get("admin"):
        return redirect(url_for("admin"))
    filename, pdf_bytes = get_pdf_from_db(invoice_id)
    if pdf_bytes:
        return send_file(
            io.BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype="application/pdf"
        )
    return "ファイルが見つかりません", 404


@app.route("/admin/delete", methods=["POST"])
def admin_delete():
    if not session.get("admin"):
        return redirect(url_for("admin"))

    # DB使用時
    invoice_id = request.form.get("invoice_id", "")
    if invoice_id and DATABASE_URL:
        try:
            delete_invoice_from_db(int(invoice_id))
        except Exception as e:
            print(f"DB削除エラー: {e}")
        return redirect(url_for("admin"))

    # ファイルシステム使用時（fallback）
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
init_db()

if __name__ == "__main__":
    print("\n  請求書・領収書 作成ツール")
    print("  http://localhost:5002\n")
    app.run(host="0.0.0.0", port=5002, debug=True)
