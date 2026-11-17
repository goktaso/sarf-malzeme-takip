import argparse
import base64
import binascii
import json
import re
import sqlite3
import sys
import unicodedata
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

import excel_import
import export_files
import sqlserver
import access
import stock_counts

ROOT = Path(__file__).resolve().parent
DB = ROOT / "data" / "stok.sqlite3"
SQL_MODE = False
STATIC = {
    "/": ("index.html", "text/html"),
    "/app.js": ("app.js", "text/javascript"),
    "/table-filters.js": ("table-filters.js", "text/javascript"),
    "/style.css": ("style.css", "text/css"),
}
IMPORT_TEMPLATE = ROOT / "outputs" / "excel-import" / "toplu-aktarim-sablonlari.xlsx"
XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

VARIANT_COLUMNS = """
    id INTEGER PRIMARY KEY, product_id INTEGER NOT NULL REFERENCES products(id),
    size TEXT NOT NULL, color TEXT NOT NULL DEFAULT '', minimum INTEGER NOT NULL,
    UNIQUE(product_id, size, color)
"""

SCHEMA = f"""
CREATE TABLE IF NOT EXISTS departments (
    id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE,
    active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1))
);
CREATE TABLE IF NOT EXISTS people (
    id INTEGER PRIMARY KEY, code TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
    department_id INTEGER NOT NULL REFERENCES departments(id),
    active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1))
);
CREATE TABLE IF NOT EXISTS product_groups (
    id INTEGER PRIMARY KEY, name TEXT NOT NULL, prefix TEXT NOT NULL UNIQUE,
    module TEXT NOT NULL CHECK(module IN ('sarf','kkd')),
    next_number INTEGER NOT NULL DEFAULT 1,
    size_kind TEXT NOT NULL DEFAULT 'size' CHECK(size_kind IN ('size','shoe'))
);
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY, code TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
    module TEXT NOT NULL CHECK(module IN ('sarf','kkd')), category TEXT NOT NULL,
    unit TEXT NOT NULL, scope TEXT NOT NULL CHECK(scope IN ('person','department','both')),
    group_id INTEGER REFERENCES product_groups(id),
    active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1))
);
CREATE TABLE IF NOT EXISTS variants (
    {VARIANT_COLUMNS}
);
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY, request_id TEXT NOT NULL UNIQUE, kind TEXT NOT NULL,
    day TEXT NOT NULL, purchase_day TEXT, supplier TEXT, reference TEXT, note TEXT,
    department_id INTEGER REFERENCES departments(id), person_id INTEGER REFERENCES people(id),
    source_id INTEGER REFERENCES documents(id)
);
CREATE TABLE IF NOT EXISTS movements (
    id INTEGER PRIMARY KEY, document_id INTEGER NOT NULL REFERENCES documents(id),
    variant_id INTEGER NOT NULL REFERENCES variants(id),
    qty INTEGER NOT NULL, value INTEGER NOT NULL, unit_price INTEGER,
    UNIQUE(document_id, variant_id)
);
CREATE INDEX IF NOT EXISTS movement_variant ON movements(variant_id);
CREATE TABLE IF NOT EXISTS definition_versions (
    id INTEGER PRIMARY KEY,
    entity TEXT NOT NULL CHECK(entity IN ('products','departments','people','product_groups')),
    entity_id INTEGER NOT NULL,
    effective_day TEXT NOT NULL,
    recorded_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    state TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS definition_as_of ON definition_versions(entity,entity_id,effective_day DESC,id DESC);
"""
VERSIONED_TABLES = ("products", "departments", "people", "product_groups")
BASELINE = "0001-01-01"
ISTANBUL = timezone(timedelta(hours=3))


@contextmanager
def connect():
    if SQL_MODE:
        db = sqlserver.connect()
        try:
            yield db
            db.commit()
        except BaseException:
            db.rollback()
            raise
        finally:
            db.close()
        return
    db = sqlite3.connect(DB, timeout=10)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys = ON")
    try:
        with db:
            yield db
    finally:
        db.close()


def initialize():
    if SQL_MODE:
        with connect() as db:
            db.raw.execute((ROOT / "sqlserver_schema.sql").read_text(encoding="utf-8"))
            if db.execute("SELECT COL_LENGTH('dbo.product_groups','size_kind')").fetchone()[0] is None:
                db.execute("""ALTER TABLE product_groups ADD size_kind varchar(8) NOT NULL
                              DEFAULT 'size' CHECK(size_kind IN ('size','shoe'))""")
                backfill_group_size_kind(db)
        return
    DB.parent.mkdir(parents=True, exist_ok=True)
    with connect() as db:
        db.execute("PRAGMA journal_mode = WAL")
        db.executescript(SCHEMA + stock_counts.SQLITE_SCHEMA)
        db.execute("PRAGMA foreign_keys = OFF")
        db.execute("BEGIN IMMEDIATE")
        if "size_kind" not in {row["name"] for row in db.execute("PRAGMA table_info(product_groups)")}:
            db.execute("""ALTER TABLE product_groups ADD size_kind TEXT NOT NULL
                          DEFAULT 'size' CHECK(size_kind IN ('size','shoe'))""")
            backfill_group_size_kind(db)
        if "group_id" not in {row["name"] for row in db.execute("PRAGMA table_info(products)")}:
            db.execute("ALTER TABLE products ADD COLUMN group_id INTEGER REFERENCES product_groups(id)")
        for table in ("products", "departments", "people"):
            if "active" not in {row["name"] for row in db.execute(f"PRAGMA table_info({table})")}:
                db.execute(f"ALTER TABLE {table} ADD COLUMN active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1))")
        if "color" not in {row["name"] for row in db.execute("PRAGMA table_info(variants)")}:
            db.execute(f"CREATE TABLE variants_new ({VARIANT_COLUMNS})")
            db.execute("""INSERT INTO variants_new(id,product_id,size,minimum)
                          SELECT id,product_id,size,minimum FROM variants""")
            db.execute("DROP TABLE variants")
            db.execute("ALTER TABLE variants_new RENAME TO variants")
        if db.execute("PRAGMA foreign_key_check").fetchone():
            raise RuntimeError("Veritabanı kayıt bağlantıları doğrulanamadı.")
        for table in VERSIONED_TABLES:
            missing = rows(db, f"""SELECT id FROM {table} t WHERE NOT EXISTS (
                SELECT 1 FROM definition_versions h WHERE h.entity=? AND h.entity_id=t.id)""", (table,))
            for row in missing:
                record_definition(db, table, row["id"], BASELINE)


def field(data, key, required=True):
    value = data.get(key, "")
    if not isinstance(value, str):
        raise ValueError("Geçersiz metin alanı.")
    value = value.strip()
    if len(value) > 250 or (required and not value):
        raise ValueError("Zorunlu alanları doldurun. Metin sınırı 250 karakterdir.")
    return value


def number(value, scale, allow_zero=False):
    try:
        original = Decimal(str(value).replace(",", "."))
        scaled = original * scale
        if not scaled.is_finite() or scaled != scaled.to_integral_value():
            raise ValueError()
        if not (int(not allow_zero) <= scaled <= 1_000_000_000):
            raise ValueError()
        return int(scaled)
    except (InvalidOperation, ValueError):
        raise ValueError("Miktar ve maliyet alanlarında geçerli bir sayı kullanın.") from None


def day_field(data, key):
    try:
        return date.fromisoformat(field(data, key)).isoformat()
    except ValueError:
        raise ValueError("Geçerli bir tarih girin.") from None


def rows(db, query, args=()):
    return [dict(row) for row in db.execute(query, args)]


def entity(db, table, raw_id):
    if not isinstance(raw_id, (str, int)) or not re.fullmatch(r"[0-9]{1,18}", str(raw_id)):
        raise ValueError("Seçili kayıt bulunamadı.")
    row = db.execute(f"SELECT * FROM {table} WHERE id = ?", (int(raw_id),)).fetchone()
    if row is None:
        raise ValueError("Seçili kayıt bulunamadı.")
    return row


def definition_state(db, table, record_id):
    state = dict(entity(db, table, record_id))
    if table == "products":
        state["variants"] = rows(db, """SELECT id,color,size,minimum FROM variants
                                       WHERE product_id=? ORDER BY color,size,id""", (record_id,))
    if table == "product_groups":
        state.pop("next_number")
    return state


def effective_day(data):
    day = day_field(data, "effective_day") if data.get("effective_day") else datetime.now(ISTANBUL).date().isoformat()
    if day == BASELINE or day > datetime.now(ISTANBUL).date().isoformat():
        raise ValueError("Geçerlilik tarihi bugün veya öncesi olmalı.")
    return day


def record_definition(db, table, record_id, day):
    state = json.dumps(definition_state(db, table, record_id), ensure_ascii=False, sort_keys=True)
    last = db.execute("""SELECT effective_day,state FROM definition_versions
                        WHERE entity=? AND entity_id=? ORDER BY effective_day DESC,id DESC LIMIT 1""",
                      (table, record_id)).fetchone()
    if last:
        if day < last["effective_day"]:
            raise ValueError(f"Geçerlilik tarihi son değişiklikten ({last['effective_day']}) önce olamaz.")
        if last["state"] == state:
            return
    db.execute("INSERT INTO definition_versions(entity,entity_id,effective_day,state) VALUES (?,?,?,?)",
               (table, record_id, day, state))


def definition_on(db, table, record_id, day):
    current = entity(db, table, record_id)
    version = db.execute("""SELECT state FROM definition_versions
                           WHERE entity=? AND entity_id=? AND effective_day<=?
                           ORDER BY effective_day DESC,id DESC LIMIT 1""",
                         (table, current["id"], day)).fetchone()
    return json.loads(version["state"]) if version else dict(current)


def definition_history(db, table, record_id):
    if table not in VERSIONED_TABLES:
        raise ValueError("Geçersiz kayıt türü.")
    row = entity(db, table, record_id)
    history = rows(db, """SELECT effective_day,recorded_at,state FROM definition_versions
                         WHERE entity=? AND entity_id=? ORDER BY effective_day DESC,id DESC""", (table, row["id"]))
    for version in history:
        version["state"] = json.loads(version["state"])
        if table == "people":
            department = definition_on(db, "departments", version["state"]["department_id"], version["effective_day"])
            version["state"]["department"] = department["name"]
        if table == "product_groups":
            state = version["state"]
            state.setdefault("size_kind", infer_size_kind(state["name"], state["prefix"]))
    return {"versions": history}


def normalized(value):
    return unicodedata.normalize("NFKC", " ".join(value.split())).translate(
        str.maketrans({"I": "ı", "İ": "i"})).casefold()


def infer_size_kind(name, prefix=""):
    return "shoe" if prefix == "AYK" or re.search(
        r"\b(?:ayakkabı\w*|bot(?:lar|u)?|çizme\w*)\b", normalized(name)) else "size"


def backfill_group_size_kind(db):
    for group in rows(db, "SELECT id,name,prefix FROM product_groups"):
        if infer_size_kind(group["name"], group["prefix"]) == "shoe":
            db.execute("UPDATE product_groups SET size_kind='shoe' WHERE id=?", (group["id"],))


def group_size_kind(data, name, prefix):
    kind = field(data, "size_kind", False) or infer_size_kind(name, prefix)
    if kind not in {"size", "shoe"}:
        raise ValueError("Beden veya ayakkabı numarası seçin.")
    return kind


def add_variant(db, product_id, size, color, minimum):
    existing = db.execute("SELECT size,color FROM variants WHERE product_id=?", (product_id,))
    if any((normalized(row["size"]), normalized(row["color"])) ==
           (normalized(size), normalized(color)) for row in existing):
        raise ValueError("Bu renk ve beden kombinasyonu zaten kayıtlı.")
    db.execute("INSERT INTO variants(product_id,size,color,minimum) VALUES (?,?,?,?)",
               (product_id, size, color, minimum))


def add_group(db, data):
    name, prefix, module = field(data, "name"), field(data, "prefix").upper(), field(data, "module")
    if module not in {"sarf", "kkd"} or not re.fullmatch(r"[A-Z]{2,8}", prefix, flags=re.ASCII):
        raise ValueError("Bir modül seçin ve grup önekine 2–8 Latin harfi girin (örneğin TSH).")
    names = db.execute("SELECT name FROM product_groups WHERE module=?", (module,))
    if any(normalized(row["name"]) == normalized(name) for row in names):
        raise ValueError("Bu modülde aynı isimde bir grup zaten var.")
    cursor = db.execute("INSERT INTO product_groups(name,prefix,module,size_kind) VALUES (?,?,?,?)",
                        (name, prefix, module, group_size_kind(data, name, prefix)))
    record_definition(db, "product_groups", cursor.lastrowid, BASELINE)


def next_product_code(db, group):
    sequence = group["next_number"]
    for row in db.execute("SELECT code FROM products WHERE code LIKE ?", (group["prefix"] + "-%",)):
        suffix = row["code"].rsplit("-", 1)[-1]
        if re.fullmatch(r"[0-9]{1,10}", suffix):
            sequence = max(sequence, int(suffix) + 1)
    db.execute("UPDATE product_groups SET next_number=? WHERE id=?", (sequence + 1, group["id"]))
    return f"{group['prefix']}-{sequence:04d}"


def validate_product_name(db, name, module, group_id, category, product_id=None):
    existing = db.execute("SELECT id,name,category,group_id FROM products WHERE module=?", (module,))
    if any(row["id"] != product_id and
           ((group_id is not None and row["group_id"] == group_id) or
            normalized(row["category"]) == normalized(category)) and
           normalized(row["name"]) == normalized(name) for row in existing):
        raise ValueError("Bu grupta aynı isimde bir ürün kartı var.")


def save_department(db, data, department_id=None):
    name = field(data, "name")
    existing = db.execute("SELECT id,name FROM departments")
    if any(row["id"] != department_id and normalized(row["name"]) == normalized(name) for row in existing):
        raise ValueError("Bu isimde bir departman zaten var.")
    if department_id is None:
        cursor = db.execute("INSERT INTO departments(name) VALUES (?)", (name,))
        record_definition(db, "departments", cursor.lastrowid, BASELINE)
    else:
        db.execute("UPDATE departments SET name=?,active=? WHERE id=?",
                   (name, active_field(data), department_id))
        record_definition(db, "departments", department_id, effective_day(data))


def active_field(data):
    value = data.get("active")
    if value not in ("0", "1"):
        raise ValueError("Kayıt durumunu seçin.")
    return int(value)


def save_product(db, data):
    product = entity(db, "products", data.get("id"))
    name, unit, scope = (field(data, key) for key in ("name", "unit", "scope"))
    if scope not in {"person", "department", "both"}:
        raise ValueError("Teslim kapsamını seçin.")
    validate_product_name(db, name, product["module"], product["group_id"],
                          product["category"], product["id"])
    if unit != product["unit"] and has_movements(db, product["id"]):
        raise ValueError("Stok hareketi bulunan ürünün stok birimi değiştirilemez.")
    db.execute("UPDATE products SET name=?, unit=?, scope=?, active=? WHERE id=?",
               (name, unit, scope, active_field(data), product["id"]))
    record_definition(db, "products", product["id"], effective_day(data))


def has_movements(db, product_id):
    return db.execute("""SELECT 1 FROM variants v JOIN movements m ON m.variant_id=v.id
                         WHERE v.product_id=? ORDER BY v.id LIMIT 1""", (product_id,)).fetchone() is not None


def add_definition(db, route, data):
    if route == "departments":
        save_department(db, data)
    elif route == "department-update":
        department = entity(db, "departments", data.get("id"))
        save_department(db, data, department["id"])
    elif route == "product-update":
        save_product(db, data)
    elif route in {"person-status", "person-update"}:
        person = entity(db, "people", data.get("id"))
        if route == "person-update":
            department = entity(db, "departments", data.get("department_id"))
            name = field(data, "name")
        else:
            department = entity(db, "departments", person["department_id"])
            name = person["name"]
        if department["id"] != person["department_id"] and not department["active"]:
            raise ValueError("Personel için aktif bir departman seçin.")
        db.execute("UPDATE people SET name=?,department_id=?,active=? WHERE id=?",
                   (name, department["id"], active_field(data), person["id"]))
        record_definition(db, "people", person["id"], effective_day(data))
    elif route == "people":
        department = entity(db, "departments", data.get("department_id"))
        if not department["active"]:
            raise ValueError("Personel için aktif bir departman seçin.")
        active = active_field(data) if "active" in data else 1
        cursor = db.execute("INSERT INTO people(code,name,department_id,active) VALUES (?,?,?,?)",
                            (field(data, "code").upper(), field(data, "name"), department["id"], active))
        record_definition(db, "people", cursor.lastrowid, BASELINE)
    elif route == "product-groups":
        add_group(db, data)
    elif route == "product-group-size-update":
        group = entity(db, "product_groups", data.get("id"))
        kind = group_size_kind(data, group["name"], group["prefix"])
        db.execute("UPDATE product_groups SET size_kind=? WHERE id=?", (kind, group["id"]))
        record_definition(db, "product_groups", group["id"], effective_day(data))
    elif route == "variants":
        product = entity(db, "products", data.get("product_id"))
        default = db.execute("SELECT 1 FROM variants WHERE product_id=? AND size='' AND color=''",
                             (product["id"],)).fetchone()
        if default:
            raise ValueError("Bu ürün kartı tek stok satırıyla açılmış.")
        size, color = field(data, "size", False), field(data, "color", False)
        if not (size or color):
            raise ValueError("Bir renk veya beden girin.")
        add_variant(db, product["id"], size, color, number(data.get("minimum", 0), 1000, True))
        record_definition(db, "products", product["id"], effective_day(data))
    else:
        if "code" in data or "category" in data:
            raise ValueError("Ürün kodu ve kategori, seçilen ürün grubundan belirlenir.")
        group = entity(db, "product_groups", data.get("group_id"))
        name, scope = field(data, "name"), field(data, "scope")
        if scope not in {"person", "department", "both"}:
            raise ValueError("Teslim kapsamını seçin.")
        validate_product_name(db, name, group["module"], group["id"], group["name"])
        sizes = [size.strip() for size in field(data, "sizes", False).split(",") if size.strip()] or [""]
        color = field(data, "color", False)
        minimum = number(data.get("minimum", 0), 1000, True)
        code = next_product_code(db, group)
        cursor = db.execute(
            "INSERT INTO products(code,name,module,category,unit,scope,group_id) VALUES (?,?,?,?,?,?,?)",
            (code, name, group["module"], group["name"],
             field(data, "unit"), scope, group["id"]))
        for size in sizes:
            add_variant(db, cursor.lastrowid, size, color, minimum)
        record_definition(db, "products", cursor.lastrowid, BASELINE)
        return code


def rounded_ratio(value, numerator, denominator):
    return int((Decimal(value) * numerator / denominator).quantize(Decimal(1), rounding=ROUND_HALF_UP))


def return_value(db, source, variant_id, qty):
    original = db.execute(
        "SELECT qty,value FROM movements WHERE document_id=? AND variant_id=?",
        (source["id"], variant_id)).fetchone()
    if not original:
        raise ValueError("Bu ürün seçilen teslim fişinde bulunmuyor.")
    returned = db.execute(
        """SELECT COALESCE(SUM(m.qty),0), COALESCE(SUM(m.value),0)
           FROM movements m JOIN documents d ON d.id=m.document_id
           WHERE d.source_id=? AND m.variant_id=?""", (source["id"], variant_id)).fetchone()
    if qty + returned[0] > -original["qty"]:
        raise ValueError("İade miktarı, teslim fişindeki kalan iade edilebilir miktarı aşıyor.")
    return rounded_ratio(-original["value"], qty + returned[0], -original["qty"]) - returned[1]


def add_document(db, data):
    request_id = field(data, "request_id")
    previous = db.execute("SELECT id FROM documents WHERE request_id=?", (request_id,)).fetchone()
    if previous:
        return previous["id"]
    kind = field(data, "kind")
    if kind not in {"receipt", "person", "department", "return", "adjustment"}:
        raise ValueError("Geçersiz hareket türü.")
    day = day_field(data, "day")
    purchase_day = day_field(data, "purchase_day") if kind == "receipt" else None
    if purchase_day and purchase_day > day:
        raise ValueError("Alış tarihi depo giriş tarihinden sonra olamaz.")
    supplier = field(data, "supplier") if kind == "receipt" else None
    reference = field(data, "reference", kind == "receipt")
    note = field(data, "note", kind == "adjustment")
    person = department = source = None
    if kind in {"person", "department"}:
        person = definition_on(db, "people", data.get("person_id"), day)
        if not person["active"]:
            raise ValueError("Teslim için çalışan bir personel seçin.")
        target = definition_on(db, "departments", person["department_id"] if kind == "person" else data.get("department_id"), day)
        if not target["active"]:
            raise ValueError("Teslim için aktif bir departman seçin.")
        if not definition_on(db, "departments", person["department_id"], day)["active"]:
            raise ValueError("Teslim alan kişinin departmanı pasif.")
        department = target["id"]
    if kind == "return":
        source = entity(db, "documents", data.get("source_id"))
        if source["kind"] not in {"person", "department"} or day < source["day"]:
            raise ValueError("İade için geçerli bir teslim fişi ve tarih seçin.")
        person = entity(db, "people", source["person_id"])
        department = source["department_id"]
    lines = data.get("lines")
    if not isinstance(lines, list) or not 1 <= len(lines) <= 100:
        raise ValueError("Fişe 1–100 ürün satırı ekleyin.")
    cursor = db.execute(
        """INSERT INTO documents(request_id,kind,day,purchase_day,supplier,reference,note,
           department_id,person_id,source_id) VALUES (?,?,?,?,?,?,?,?,?,?)""",
        (request_id, kind, day, purchase_day, supplier, reference, note, department,
         person["id"] if person else None, source["id"] if source else None))
    document_id = cursor.lastrowid
    seen = set()
    for line in lines:
        if not isinstance(line, dict):
            raise ValueError("Geçersiz ürün satırı.")
        variant = entity(db, "variants", line.get("variant_id"))
        if variant["id"] in seen:
            raise ValueError("Aynı ürün ve beden için tek satır kullanın.")
        seen.add(variant["id"])
        product = definition_on(db, "products", variant["product_id"], day)
        if not product["active"] and kind not in {"return", "adjustment"}:
            raise ValueError(f"{product['name']} pasif. Alış veya teslim için kartı aktifleştirin.")
        if kind in {"person", "department"} and product["scope"] not in {kind, "both"}:
            raise ValueError(f"{product['name']} için bu teslim türü seçilemez.")
        balance = db.execute(
            """SELECT COALESCE(SUM(qty),0), COALESCE(SUM(value),0), MAX(d.day)
               FROM movements m JOIN documents d ON d.id=m.document_id WHERE variant_id=?""",
            (variant["id"],)).fetchone()
        stock, stock_value, latest = balance
        if latest and day < latest:
            raise ValueError(f"{product['name']} için son hareket tarihi {latest}. Bu tarihten itibaren kayıt girin.")
        qty = number(line.get("qty"), 1000, kind == "adjustment")
        unit_price = None
        if kind == "adjustment":
            qty -= stock
            if qty == 0:
                raise ValueError("Sayılan miktar mevcut stokla aynı.")
        if kind == "receipt" or (kind == "adjustment" and qty > 0):
            unit_price = number(line.get("unit_price"), 10000, True)
            value = rounded_ratio(unit_price, qty, 100000)
        elif kind == "return":
            value = return_value(db, source, variant["id"], qty)
        else:
            qty = -qty if kind in {"person", "department"} else qty
            if -qty > stock:
                raise ValueError(f"{product['name']} için yetersiz stok.")
            value = rounded_ratio(stock_value, qty, stock)
        db.execute("INSERT INTO movements(document_id,variant_id,qty,value,unit_price) VALUES (?,?,?,?,?)",
                   (document_id, variant["id"], qty, value, unit_price))
    return document_id


def snapshot(db):
    result = {
        "demo": DB.name == "demo.sqlite3",
        "departments": rows(db, "SELECT * FROM departments ORDER BY name"),
        "people": rows(db, """SELECT p.*, d.name department,d.active department_active FROM people p
                               JOIN departments d ON d.id=p.department_id ORDER BY p.name"""),
        "definition_timeline": rows(db, """SELECT entity,entity_id,effective_day,
            json_extract(state,'$.active') active,json_extract(state,'$.name') name,
            json_extract(state,'$.department_id') department_id,json_extract(state,'$.scope') scope
            FROM definition_versions WHERE entity IN ('products','people','departments')
            ORDER BY effective_day DESC,id DESC"""),
        "products": rows(db, """SELECT p.*, CASE WHEN EXISTS(SELECT 1 FROM variants v
            JOIN movements m ON m.variant_id=v.id WHERE v.product_id=p.id) THEN 1 ELSE 0 END has_movements
            FROM products p ORDER BY name"""),
        "groups": rows(db, "SELECT id,name,prefix,module,size_kind FROM product_groups ORDER BY name"),
        "stock": rows(db, """SELECT v.*, p.code,p.name,p.module,p.category,p.unit,p.scope,p.active,g.size_kind,
                             COALESCE(m.qty,0) qty,COALESCE(m.value,0) value
                             FROM variants v JOIN products p ON p.id=v.product_id
                             LEFT JOIN product_groups g ON g.id=p.group_id
                             LEFT JOIN (SELECT variant_id,SUM(qty) qty,SUM(value) value
                                       FROM movements GROUP BY variant_id) m ON m.variant_id=v.id
                             ORDER BY p.name,v.color,v.size"""),
        "movements": rows(db, """SELECT m.*,d.kind,d.day,d.purchase_day,d.supplier,d.reference,d.note,d.source_id,
                                 d.department_id,d.person_id,s.kind source_kind,
                                 COALESCE(json_extract(pv.state,'$.name'),p.name) name,
                                 p.code,COALESCE(json_extract(pv.state,'$.unit'),p.unit) unit,p.module,v.size,v.color,
                                 COALESCE(json_extract(dv.state,'$.name'),dp.name) department,
                                 COALESCE(json_extract(ev.state,'$.name'),pe.name) person,
                                 COALESCE(json_extract(ev.state,'$.active'),pe.active) person_active
                                 FROM movements m JOIN documents d ON d.id=m.document_id
                                 JOIN variants v ON v.id=m.variant_id JOIN products p ON p.id=v.product_id
                                 LEFT JOIN departments dp ON dp.id=d.department_id
                                 LEFT JOIN people pe ON pe.id=d.person_id
                                 LEFT JOIN definition_versions pv ON pv.id=(
                                    SELECT id FROM definition_versions WHERE entity='products' AND entity_id=p.id
                                    AND effective_day<=d.day ORDER BY effective_day DESC,id DESC LIMIT 1)
                                 LEFT JOIN definition_versions dv ON dv.id=(
                                    SELECT id FROM definition_versions WHERE entity='departments' AND entity_id=dp.id
                                    AND effective_day<=d.day ORDER BY effective_day DESC,id DESC LIMIT 1)
                                 LEFT JOIN definition_versions ev ON ev.id=(
                                    SELECT id FROM definition_versions WHERE entity='people' AND entity_id=pe.id
                                    AND effective_day<=d.day ORDER BY effective_day DESC,id DESC LIMIT 1)
                                 LEFT JOIN documents s ON s.id=d.source_id
                                 ORDER BY d.day DESC,d.id DESC,m.id"""),
    }
    for item in result["stock"]:
        if item["size_kind"] is None:
            item["size_kind"] = infer_size_kind(item["category"])
    return result


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def send_body(self, status, body, content_type, attachment=None, headers=()):
        self.send_response(status)
        self.send_header("Content-Type", content_type if attachment else content_type + "; charset=utf-8")
        if attachment:
            self.send_header("Content-Disposition", f'attachment; filename="{attachment}"')
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy",
                         "default-src 'self'; object-src 'none'; frame-ancestors 'none'; base-uri 'none'")
        for name, value in headers:
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(body)

    def respond(self, status, payload, headers=()):
        self.send_body(status, json.dumps(payload, ensure_ascii=False).encode("utf-8"), "application/json",
                       headers=headers)

    def local_request(self):
        host = self.headers.get("Host", "").lower()
        port = self.server.server_port
        origin = self.headers.get("Origin")
        if (host not in {f"localhost:{port}", f"127.0.0.1:{port}"} or
                (origin and origin != "http://" + host) or
                self.headers.get("Sec-Fetch-Site") == "cross-site"):
            self.close_connection = True
            self.respond(403, {"error": "Bu adres yerel uygulama için geçerli değil."})
            return False
        return True

    def current_user(self, db):
        return access.authenticate(db, self.headers.get("Cookie"))

    def signed_in(self, db):
        user = self.current_user(db)
        if user is None:
            raise access.Forbidden("Oturum açmanız gerekiyor.")
        return user

    def windows_login(self):
        import pywintypes
        import sspi
        import win32api
        import win32con
        import win32security

        authorization = self.headers.get("Authorization", "")
        scheme, _, encoded = authorization.partition(" ")
        challenge = (("WWW-Authenticate", "Negotiate"),)
        if scheme.lower() != "negotiate" or not encoded:
            self.respond(401, {"error": "Windows oturumu doğrulanamadı. Tarayıcıda Windows oturum açmayı etkinleştirin."},
                         challenge)
            return
        try:
            context = getattr(self, "windows_context", None) or sspi.ServerAuth("Negotiate")
            self.windows_context = context
            _, outgoing = context.authorize(base64.b64decode(encoded, validate=True))
            token = outgoing[0].Buffer
            headers = (("WWW-Authenticate", "Negotiate " + base64.b64encode(token).decode("ascii")),) if token else ()
            if not context.authenticated:
                self.respond(401, {"error": "Windows doğrulaması sürüyor."}, headers or challenge)
                return
            context.ctxt.ImpersonateSecurityContext()
            try:
                thread = win32api.GetCurrentThread()
                identity = win32security.OpenThreadToken(thread, win32con.TOKEN_QUERY, True)
                try:
                    sid, _ = win32security.GetTokenInformation(identity, win32security.TokenUser)
                    sid_string = win32security.ConvertSidToStringSid(sid)
                finally:
                    identity.Close()
            finally:
                context.ctxt.RevertSecurityContext()
            with connect() as db:
                db.execute("BEGIN IMMEDIATE")
                user, cookie = access.start_session(db, sid_string)
            self.respond(200, {"user": access.profile(user)}, headers + (("Set-Cookie", cookie),))
            self.windows_context = None
        except (binascii.Error, pywintypes.error):
            self.windows_context = None
            self.respond(401, {"error": "Windows hesabı doğrulanamadı."}, challenge)
        except access.Forbidden as error:
            self.windows_context = None
            self.respond(403, {"error": str(error)})
        except sqlite3.OperationalError:
            self.respond(503, {"error": "SQL Server bağlantısı kurulamadı."})

    def do_GET(self):
        url = urlsplit(self.path)
        path = url.path
        if SQL_MODE and not self.local_request():
            return
        if path == "/api/session":
            if SQL_MODE:
                with connect() as db:
                    user = self.current_user(db)
                self.respond(200, {"required": True, "user": access.profile(user) if user else None})
            else:
                self.respond(200, {"required": False, "user": None})
        elif path == "/api/windows-login" and SQL_MODE:
            self.windows_login()
        elif path == "/api/state":
            try:
                with connect() as db:
                    db.execute("BEGIN")
                    user = self.signed_in(db) if SQL_MODE else None
                    result = snapshot(db)
                    if user:
                        result = access.state_for_user(result, user)
                        result["session"] = access.profile(user)
                        if user["is_admin"]:
                            result["period_grants"] = access.period_grant_records(db, user, sys.modules[__name__])
                            result["accounts"] = [access.profile(row) | {"active": row["active"]}
                                for row in db.execute("SELECT * FROM app_users ORDER BY account")]
                            result["permission_labels"] = access.PERMISSIONS
                            result["audit"] = rows(db, """SELECT TOP 200 a.action,a.record_ref,a.recorded_at,u.account
                                FROM app_audit a JOIN app_users u ON u.sid=a.sid ORDER BY a.id DESC""")
                        else:
                            result["open_periods"] = access.user_open_periods(db, user, sys.modules[__name__])
                    if not user or (access.permitted(user, "view.sayimlar") and
                                    (user["is_admin"] or user["department_id"] is None)):
                        result["counts"] = stock_counts.counts(db, sys.modules[__name__])
                    result["open_period_start"] = access.open_period_start(sys.modules[__name__])
                self.respond(200, result)
            except access.Forbidden as error:
                self.respond(401, {"error": str(error)})
            except sqlite3.OperationalError:
                self.respond(503, {"error": "Veritabanı bağlantısı kurulamadı."})
        elif path == "/api/definition-history":
            query = parse_qs(url.query)
            try:
                with connect() as db:
                    db.execute("BEGIN")
                    table, record_id = query.get("entity", [""])[0], query.get("id", [""])[0]
                    user = self.signed_in(db) if SQL_MODE else None
                    if user:
                        access.check_history(db, user, table, record_id, sys.modules[__name__])
                    result = definition_history(db, table, record_id)
                    if user and not user["is_admin"] and user["department_id"] is not None and table == "people":
                        result["versions"] = [version for version in result["versions"]
                            if version["state"]["department_id"] == user["department_id"]]
                self.respond(200, result)
            except access.Forbidden as error:
                self.respond(403, {"error": str(error)})
            except ValueError as error:
                self.respond(400, {"error": str(error)})
        elif path in {"/api/stock-count", "/sayim-sablonu.xlsx"}:
            self.count_info(parse_qs(url.query), path.endswith(".xlsx"))
        elif path in STATIC:
            file, content_type = STATIC[path]
            self.send_body(200, (ROOT / "web" / file).read_bytes(), content_type)
        elif path == "/sablonlar/toplu-aktarim.xlsx" and IMPORT_TEMPLATE.is_file():
            if SQL_MODE:
                with connect() as db:
                    user = self.current_user(db)
                if user is None or not user["is_admin"]:
                    self.respond(403, {"error": "Toplu aktarım yöneticilere açıktır."})
                    return
            self.send_body(200, IMPORT_TEMPLATE.read_bytes(), XLSX_TYPE, IMPORT_TEMPLATE.name)
        else:
            self.respond(404, {"error": "Sayfa bulunamadı."})

    def do_POST(self):
        if SQL_MODE and not self.local_request():
            return
        url = urlsplit(self.path)
        route = url.path.removeprefix("/api/")
        if route == "count-import":
            self.receive_count(parse_qs(url.query))
            return
        if route in {"excel-preview", "excel-import"}:
            self.receive_workbook(route == "excel-preview")
            return
        if route in {"export-xlsx", "export-pdf", "export-csv"}:
            self.export_report(route.removeprefix("export-"))
            return
        if route not in {"products", "product-update", "product-groups", "product-group-size-update", "variants",
                         "departments", "department-update", "people", "person-status", "person-update", "documents",
                         "user-update", "logout", "count-create", "count-line", "count-transition",
                         "period-add", "period-revoke"} or (
                             not SQL_MODE and route in {"user-update", "logout", "period-add", "period-revoke"}):
            self.respond(404, {"error": "İşlem bulunamadı."})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if self.headers.get_content_type() != "application/json" or not 0 < length <= 128000:
                raise ValueError("Geçersiz istek.")
            data = json.loads(self.rfile.read(length))
            if not isinstance(data, dict):
                raise ValueError("Geçersiz istek.")
            with connect() as db:
                db.execute("BEGIN IMMEDIATE")
                result = {}
                user = self.signed_in(db) if SQL_MODE else None
                model = sys.modules[__name__]
                if route == "logout":
                    token = access.read_cookie(self.headers.get("Cookie"))
                    db.execute("DELETE FROM app_sessions WHERE token_hash=?", (access.token_hash(token),))
                    access.audit(db, user, "logout")
                elif route == "user-update":
                    access.save_user(db, user, data, model)
                elif route in {"period-add", "period-revoke"}:
                    result["grant_id"] = access.change_period_grant(db, user, route, data, model)
                elif route in {"count-create", "count-line", "count-transition"}:
                    result["id"] = self.change_count(db, user, route, data, model)
                elif route == "documents":
                    if user:
                        access.check_write(db, user, route, data, model)
                    result["id"] = add_document(db, data)
                    if user:
                        access.audit(db, user, "document-" + data["kind"], result["id"])
                else:
                    if user:
                        access.check_write(db, user, route, data, model)
                    code = add_definition(db, route, data)
                    if code:
                        result["code"] = code
                    if user:
                        access.audit(db, user, route, code or data.get("id") or data.get("product_id") or data.get("name"))
            headers = (("Set-Cookie", access.cookie_value()),) if route == "logout" else ()
            self.respond(201, result, headers)
        except access.Forbidden as error:
            self.respond(403, {"error": str(error)})
        except (ValueError, OverflowError) as error:
            self.respond(400, {"error": str(error) or "Geçersiz istek."})
        except sqlite3.IntegrityError:
            self.respond(409, {"error": "Bu kod, ad veya beden zaten kayıtlı."})
        except sqlite3.OperationalError:
            self.respond(503, {"error": "Veritabanına ulaşılamadı. İşlemi tekrar deneyin."})

    def change_count(self, db, user, route, data, model):
        if route == "count-create":
            record_id = stock_counts.create(db, data, user, model)
            action = route
        else:
            stock_counts.allow(user, "view.sayimlar")
            count = entity(db, "stock_counts", data.get("id"))
            stock_counts.writable(db, count, user, model, data.get("revision"))
            operation = data.get("operation")
            if route == "count-line":
                stock_counts.allow(user, "count.edit")
                stock_counts.save_line(db, count, data, user, model)
            else:
                stock_counts.transition(db, count, operation, user, model)
            stock_counts.bump(db, count)
            record_id = count["id"]
            action = route + "-" + operation
        if user:
            access.audit(db, user, action, record_id)
        return record_id

    def count_info(self, query, template):
        try:
            model = sys.modules[__name__]
            with connect() as db:
                db.execute("BEGIN")
                user = self.signed_in(db) if SQL_MODE else None
                stock_counts.allow(user, "count.edit" if template else "view.sayimlar")
                count = entity(db, "stock_counts", query.get("id", [""])[0])
                if template:
                    stock_counts.writable(db, count, user, model, count["revision"])
                result = stock_counts.detail(db, count, user, model)
            if template:
                body = stock_counts.workbook(result, not user or access.permitted(user, "field.cost"))
                self.send_body(200, body, XLSX_TYPE, f"sayim-{count['id']}-sablon.xlsx")
            else:
                self.respond(200, result)
        except access.Forbidden as error:
            self.respond(403, {"error": str(error)})
        except ValueError as error:
            self.respond(400, {"error": str(error)})
        except sqlite3.OperationalError:
            self.respond(503, {"error": "Veritabanına ulaşılamadı."})
        except ModuleNotFoundError:
            self.respond(503, {"error": "Excel paketi bulunamadı."})

    def receive_count(self, query):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if (self.headers.get_content_type() not in {XLSX_TYPE, "application/octet-stream"}
                    or not 0 < length <= excel_import.MAX_BYTES):
                raise ValueError("En çok 5 MB boyutunda .xlsx dosyası seçin.")
            self.connection.settimeout(20)
            body = self.rfile.read(length)
            model = sys.modules[__name__]
            with connect() as db:
                db.execute("BEGIN IMMEDIATE")
                user = self.signed_in(db) if SQL_MODE else None
                stock_counts.allow(user, "count.edit")
                count = entity(db, "stock_counts", query.get("id", [""])[0])
                stock_counts.writable(db, count, user, model, query.get("revision", [""])[0])
                records = stock_counts.spreadsheet_rows(body)
                stock_counts.import_lines(db, count, records, user, model)
                stock_counts.bump(db, count)
                if user:
                    access.audit(db, user, "count-import", count["id"])
            self.respond(201, {"rows": len(records)})
        except access.Forbidden as error:
            self.respond(403, {"error": str(error)})
        except (ValueError, OverflowError) as error:
            self.respond(400, {"error": str(error)})
        except TimeoutError:
            self.respond(408, {"error": "Excel yüklemesi zaman aşımına uğradı."})
        except (sqlite3.OperationalError, ModuleNotFoundError):
            self.respond(503, {"error": "Sayım dosyası aktarılamadı. Tekrar deneyin."})

    def receive_workbook(self, preview):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if (self.headers.get_content_type() not in {XLSX_TYPE, "application/octet-stream"}
                    or not 0 < length <= excel_import.MAX_BYTES):
                raise ValueError("En çok 5 MB boyutunda .xlsx dosyası seçin.")
            self.connection.settimeout(20)
            workbook = excel_import.read_workbook(self.rfile.read(length))
            with connect() as db:
                db.execute("BEGIN IMMEDIATE")
                user = self.signed_in(db) if SQL_MODE else None
                if user and not user["is_admin"]:
                    raise access.Forbidden("Toplu aktarım yöneticilere açıktır.")
                result = excel_import.apply_workbook(db, workbook, sys.modules[__name__])
                if preview:
                    db.rollback()
                elif user:
                    access.audit(db, user, "excel-import", json.dumps(result["counts"], ensure_ascii=False))
            self.respond(200 if preview else 201, result)
        except access.Forbidden as error:
            self.respond(403, {"error": str(error)})
        except (ValueError, OverflowError) as error:
            self.respond(400, {"error": str(error) or "Excel verisi doğrulanamadı."})
        except TimeoutError:
            self.respond(408, {"error": "Excel yüklemesi zaman aşımına uğradı."})
        except ModuleNotFoundError:
            self.respond(503, {"error": "Excel desteği için requirements.txt içindeki paketleri kurun."})
        except sqlite3.OperationalError:
            self.respond(503, {"error": "Veritabanına ulaşılamadı. İşlemi tekrar deneyin."})

    def export_report(self, extension):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if self.headers.get_content_type() != "application/json" or not 0 < length <= export_files.MAX_BYTES:
                raise ValueError("Çıktı isteği geçersiz veya çok büyük. Filtre aralığını daraltın.")
            self.connection.settimeout(20)
            payload = json.loads(self.rfile.read(length))
            if SQL_MODE:
                with connect() as db:
                    user = self.signed_in(db)
                    require_screen = next((key for key, title in access.SCREENS.items()
                                           if isinstance(payload, dict) and title == payload.get("title")), None)
                    if require_screen is None:
                        raise access.Forbidden("Bu ekranın çıktısı alınamaz.")
                    access.require(user, "view." + require_screen)
                    access.require(user, "export")
            arguments = export_files.validate_export(payload)
            if extension == "xlsx":
                body, content_type = export_files.write_xlsx(*arguments), XLSX_TYPE
            elif extension == "csv":
                body, content_type = export_files.write_csv(*arguments), "text/csv"
            else:
                body, content_type = export_files.write_pdf(*arguments), "application/pdf"
            self.send_body(200, body, content_type, "sarf-malzeme-raporu." + extension)
        except access.Forbidden as error:
            self.respond(403, {"error": str(error)})
        except (ValueError, OverflowError, RuntimeError) as error:
            self.respond(400, {"error": str(error) or "Çıktı hazırlanamadı."})
        except TimeoutError:
            self.respond(408, {"error": "Çıktı isteği zaman aşımına uğradı."})
        except ModuleNotFoundError:
            self.respond(503, {"error": "Çıktı desteği için requirements.txt içindeki paketleri kurun."})


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--demo", action="store_true", help="Demo veritabanıyla çalıştır.")
    parser.add_argument("--legacy-sqlite", action="store_true", help="Eski yerel SQLite veritabanıyla çalıştır.")
    args = parser.parse_args()
    if args.demo:
        DB = ROOT / "data" / "demo.sqlite3"
    SQL_MODE = not (args.demo or args.legacy_sqlite)
    initialize()
    if SQL_MODE:
        sqlserver.load_runtime()
        import sspi  # Fail at startup if native Windows authentication is unavailable.
    with ThreadingHTTPServer(("127.0.0.1", args.port), Handler) as server:
        print(f"Sarf Malzeme Takip · http://127.0.0.1:{args.port}", flush=True)
        server.serve_forever()
