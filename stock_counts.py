"""Physical count documents. Only approval writes to the inventory ledger."""

from datetime import datetime
from decimal import Decimal
from io import BytesIO

import access
import excel_import

SQLITE_SCHEMA = """
CREATE TABLE IF NOT EXISTS stock_counts (
    id INTEGER PRIMARY KEY, day TEXT NOT NULL, note TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'draft' CHECK(status IN ('draft','locked','rejected','approved')),
    revision INTEGER NOT NULL DEFAULT 0, owner TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS stock_count_lines (
    count_id INTEGER NOT NULL REFERENCES stock_counts(id), variant_id INTEGER NOT NULL REFERENCES variants(id),
    code TEXT NOT NULL, name TEXT NOT NULL, category TEXT NOT NULL, color TEXT NOT NULL,
    size TEXT NOT NULL, size_kind TEXT NOT NULL, unit TEXT NOT NULL,
    expected_qty INTEGER NOT NULL, counted_qty INTEGER, unit_price INTEGER,
    PRIMARY KEY(count_id, variant_id)
);
CREATE TABLE IF NOT EXISTS stock_count_documents (
    count_id INTEGER NOT NULL REFERENCES stock_counts(id), document_id INTEGER NOT NULL REFERENCES documents(id),
    PRIMARY KEY(count_id, document_id)
);
"""

SHEET = "Sayım"
HEADERS = ("Ürün kodu", "Ürün", "Renk", "Beden / numara", "Birim",
           "Sistem bakiyesi", "Sayılan miktar", "Birim maliyet TL")


def allow(user, permission):
    if user:
        access.require(user, "view.sayimlar")
        access.require(user, permission)
        if user["department_id"] is not None and not user["is_admin"]:
            raise access.Forbidden("Depo sayımı şirket geneli yetki gerektirir.")


def admin(user):
    return user is None or bool(user["is_admin"])


def counts(db, model):
    return model.rows(db, """SELECT c.*,COALESCE(l.total,0) total,COALESCE(l.entered,0) entered
        FROM stock_counts c LEFT JOIN
        (SELECT count_id,COUNT(*) total,COUNT(counted_qty) entered FROM stock_count_lines GROUP BY count_id) l
        ON l.count_id=c.id ORDER BY c.day DESC,c.id DESC""")


def detail(db, count, user, model):
    count = dict(count)
    count["lines"] = model.rows(db, """SELECT * FROM stock_count_lines WHERE count_id=?
        ORDER BY name,color,size,variant_id""", (count["id"],))
    if user and not access.permitted(user, "field.cost"):
        for line in count["lines"]:
            line.pop("unit_price")
    count["documents"] = model.rows(db, """SELECT d.id,d.day FROM stock_count_documents c
        JOIN documents d ON d.id=c.document_id WHERE c.count_id=? ORDER BY d.id""", (count["id"],))
    return count


def balances(db, day):
    return {row["variant_id"]: row["qty"] for row in db.execute("""
        SELECT m.variant_id,SUM(m.qty) qty FROM movements m JOIN documents d ON d.id=m.document_id
        WHERE d.day<=? GROUP BY m.variant_id""", (day,))}


def add_line(db, count, variant_id, model, expected):
    variant = model.entity(db, "variants", variant_id)
    product = model.definition_on(db, "products", variant["product_id"], count["day"])
    group = (model.definition_on(db, "product_groups", product["group_id"], count["day"])
             if product["group_id"] else {})
    kind = group.get("size_kind") or model.infer_size_kind(product["category"])
    db.execute("""INSERT INTO stock_count_lines(count_id,variant_id,code,name,category,color,size,size_kind,unit,expected_qty)
                  VALUES (?,?,?,?,?,?,?,?,?,?)""",
               (count["id"], variant["id"], product["code"], product["name"], product["category"],
                variant["color"], variant["size"], kind, product["unit"], expected.get(variant["id"], 0)))


def create(db, data, user, model):
    allow(user, "count.edit")
    day = model.day_field(data, "day")
    access.check_period(db, user, day, model)
    if day > datetime.now(model.ISTANBUL).date().isoformat():
        raise ValueError("Sayım tarihi bugün veya öncesi olmalı.")
    expected = balances(db, day)
    cursor = db.execute("INSERT INTO stock_counts(day,note,owner) VALUES (?,?,?)",
                        (day, model.field(data, "note", False), user["account"] if user else "Demo"))
    count = {"id": cursor.lastrowid, "day": day}
    variants = list(db.execute("""SELECT v.id,p.active FROM variants v JOIN products p ON p.id=v.product_id
        ORDER BY v.id"""))
    for variant in variants:
        if variant["active"] or expected.get(variant["id"], 0):
            add_line(db, count, variant["id"], model, expected)
    return count["id"]


def writable(db, count, user, model, revision):
    access.check_period(db, user, count["day"], model)
    if count["status"] == "approved":
        raise ValueError("Onaylanmış sayım saklanır. Yeni bir sayımla düzeltme yapın.")
    if count["status"] != "draft" and not admin(user):
        raise access.Forbidden("Bu sayım kilitli. Yalnızca yönetici işlem yapabilir.")
    if str(revision) != str(count["revision"]):
        raise ValueError("Sayım başka bir işlemle güncellenmiş. Sayfayı yenileyip tekrar deneyin.")


def bump(db, count):
    db.execute("UPDATE stock_counts SET revision=revision+1 WHERE id=?", (count["id"],))


def line_exists(db, count, variant_id):
    return db.execute("SELECT 1 FROM stock_count_lines WHERE count_id=? AND variant_id=?",
                      (count["id"], variant_id)).fetchone() is not None


def save_line(db, count, data, user, model):
    variant_id = model.entity(db, "variants", data.get("variant_id"))["id"]
    operation = data.get("operation")
    exists = line_exists(db, count, variant_id)
    if operation == "add":
        if exists:
            raise ValueError("Ürün zaten sayım listesinde.")
        add_line(db, count, variant_id, model, balances(db, count["day"]))
    elif not exists:
        raise ValueError("Ürün sayım listesinde bulunamadı.")
    elif operation == "delete":
        db.execute("DELETE FROM stock_count_lines WHERE count_id=? AND variant_id=?", (count["id"], variant_id))
    elif operation == "save":
        qty = data.get("counted_qty", "")
        quantity = None if qty == "" else model.number(qty, 1000, True)
        price = data.get("unit_price")
        if price is not None and user:
            access.require(user, "field.cost")
        if price is not None:
            db.execute("""UPDATE stock_count_lines SET counted_qty=?,unit_price=?
                          WHERE count_id=? AND variant_id=?""",
                       (quantity, None if price == "" else model.number(price, 10000, True), count["id"], variant_id))
        else:
            db.execute("UPDATE stock_count_lines SET counted_qty=? WHERE count_id=? AND variant_id=?",
                       (quantity, count["id"], variant_id))
    else:
        raise ValueError("Geçersiz sayım satırı işlemi.")


def approval_lines(db, count, model):
    checked = model.rows(db, """SELECT * FROM stock_count_lines WHERE count_id=? AND counted_qty IS NOT NULL
                               ORDER BY variant_id""", (count["id"],))
    if not checked:
        raise ValueError("Onay için en az bir ürünün sayılan miktarını girin.")
    expected = balances(db, count["day"])
    posting_day = count["day"]
    lines = []
    for row in checked:
        variant_id = row["variant_id"]
        label = " · ".join(filter(None, (row["code"], row["name"], row["color"], row["size"])))
        if expected.get(variant_id, 0) != row["expected_qty"]:
            raise ValueError(f"{label}: sayım tarihi bakiyesi değişti. Sistem bakiyelerini yenileyip tekrar kontrol edin.")
        delta = row["counted_qty"] - row["expected_qty"]
        if not delta:
            continue
        current, value, latest = db.execute("""SELECT COALESCE(SUM(m.qty),0),COALESCE(SUM(m.value),0),MAX(d.day)
            FROM movements m JOIN documents d ON d.id=m.document_id WHERE m.variant_id=?""", (variant_id,)).fetchone()
        if latest and latest > posting_day:
            posting_day = latest
        if current + delta < 0:
            raise ValueError(f"{label}: sayım farkı güncel stoğu eksiye düşürüyor.")
        line = {"variant_id": variant_id, "qty": str(Decimal(current + delta) / 1000)}
        if delta > 0:
            price = row["unit_price"]
            if price is None and current > 0:
                price = model.rounded_ratio(value, 100000, current)
            if price is None:
                last = db.execute("""SELECT m.unit_price FROM movements m JOIN documents d ON d.id=m.document_id
                    WHERE m.variant_id=? AND m.unit_price IS NOT NULL ORDER BY d.day DESC,m.id DESC LIMIT 1""",
                    (variant_id,)).fetchone()
                price = last[0] if last else None
            if price is None:
                raise ValueError(f"{label}: değerleme için birim maliyet girin.")
            line["unit_price"] = str(Decimal(price) / 10000)
        lines.append(line)
    return posting_day, lines


def transition(db, count, operation, user, model):
    if operation == "refresh":
        allow(user, "count.edit")
        expected = balances(db, count["day"])
        for row in list(db.execute("SELECT variant_id FROM stock_count_lines WHERE count_id=?", (count["id"],))):
            db.execute("UPDATE stock_count_lines SET expected_qty=? WHERE count_id=? AND variant_id=?",
                       (expected.get(row[0], 0), count["id"], row[0]))
        return
    if operation == "lock" and count["status"] == "draft":
        allow(user, "count.edit")
        status = "locked"
    elif operation == "reopen" and count["status"] in ("locked", "rejected") and admin(user):
        status = "draft"
    elif operation == "reject" and count["status"] in ("draft", "locked"):
        allow(user, "count.approve")
        status = "rejected"
    elif operation == "approve" and count["status"] in ("draft", "locked"):
        allow(user, "count.approve")
        day, lines = approval_lines(db, count, model)
        for start in range(0, len(lines), 100):
            doc_id = model.add_document(db, {
                "request_id": f"physical-count:{count['id']}:{start // 100}",
                "kind": "adjustment", "day": day, "reference": f"Sayım #{count['id']}",
                "note": f"Onaylı fiziksel sayım. Sayım tarihi {count['day']}.",
                "lines": lines[start:start + 100]
            })
            db.execute("INSERT INTO stock_count_documents(count_id,document_id) VALUES (?,?)", (count["id"], doc_id))
        status = "approved"
    else:
        raise ValueError("Bu sayım durumunda işlem yapılamaz.")
    db.execute("UPDATE stock_counts SET status=? WHERE id=?", (status, count["id"]))


def workbook(count, can_price):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    book = Workbook()
    sheet = book.active
    sheet.title = SHEET
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="164C40")
    for line in count["lines"]:
        sheet.append((line["code"], line["name"], line["color"], line["size"], line["unit"],
                      line["expected_qty"] / 1000,
                      None if line["counted_qty"] is None else line["counted_qty"] / 1000,
                      None if not can_price or line.get("unit_price") is None else line["unit_price"] / 10000))
        for cell in sheet[sheet.max_row][:5]:
            cell.data_type = "s"
            cell.number_format = "@"
    for column, width in zip("ABCDEFGH", (16, 34, 15, 18, 14, 19, 19, 19)):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    output = BytesIO()
    book.save(output)
    return output.getvalue()


def spreadsheet_rows(payload):
    book = excel_import.open_workbook(payload)
    try:
        if SHEET not in book.sheetnames:
            raise ValueError("Sayım şablonundaki «Sayım» sayfası bulunamadı.")
        sheet = book[SHEET]
        if (sheet.max_row or 0) > excel_import.MAX_ROWS + 1 or (sheet.max_column or 0) > len(HEADERS):
            raise ValueError("Şablon sütunlarını koruyun ve en fazla 5.000 veri satırı yükleyin.")
        rows = sheet.iter_rows(max_col=len(HEADERS))
        if tuple(excel_import.cell_text(cell) for cell in next(rows, ())) != HEADERS:
            raise ValueError("Sayım Excel sütun başlıklarını ve sırasını koruyun.")
        result = []
        for number, row in enumerate(rows, 2):
            if number > excel_import.MAX_ROWS + 1:
                raise ValueError("Sayım dosyası en fazla 5.000 veri satırı içerebilir.")
            try:
                values = [excel_import.cell_text(cell) for cell in row]
            except ValueError as error:
                raise ValueError(f"Sayım, satır {number}: {error}") from error
            if values[6] or values[7]:
                result.append((number, values))
        if not result:
            raise ValueError("Excel dosyasında sayılan miktar girilmiş satır bulunamadı.")
        return result
    finally:
        book.close()


def import_lines(db, count, records, user, model):
    labels = {tuple(model.normalized(row[key]) for key in ("code", "color", "size")): row["variant_id"]
              for row in db.execute("SELECT * FROM stock_count_lines WHERE count_id=?", (count["id"],))}
    used = set()
    for number, (code, _, color, size, _, _, qty, price) in records:
        try:
            variant = labels.get(tuple(model.normalized(value) for value in (code, color, size)))
            if variant is None or variant in used:
                raise ValueError("Ürün kodu, renk ve beden eşleşmedi veya dosyada tekrar ediyor.")
            if not qty:
                raise ValueError("Birim maliyetle birlikte sayılan miktarı girin.")
            used.add(variant)
            payload = {"variant_id": variant, "operation": "save", "counted_qty": qty}
            if price:
                payload["unit_price"] = price
            save_line(db, count, payload, user, model)
        except (ValueError, access.Forbidden) as error:
            raise type(error)(f"Sayım, satır {number}: {error}") from error
