import json
import sqlite3
from datetime import datetime, date
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile, ZipFile

HEADERS = json.loads(Path(__file__).with_name("import_columns.json").read_text(encoding="utf-8"))
MAX_BYTES = 5 * 1024 * 1024
MAX_ROWS = 5000


def cell_text(cell):
    if cell.data_type in {"f", "e", "b"}:
        raise ValueError("Formül, Excel hatası ve mantıksal değer yerine düz değer kullanın.")
    value = cell.value
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if value is None:
        return ""
    return str(value).strip()


def open_workbook(payload):
    from openpyxl import load_workbook
    from openpyxl.utils.exceptions import InvalidFileException

    stream = BytesIO(payload)
    try:
        with ZipFile(stream) as archive:
            files = archive.infolist()
            if (len(files) > 250 or sum(info.file_size for info in files) > 32 * 1024 * 1024
                    or any(info.flag_bits & 1 for info in files)):
                raise ValueError("Excel dosyası çok büyük veya şifreli.")
        return load_workbook(stream, read_only=True, keep_links=False)
    except (BadZipFile, InvalidFileException, KeyError, OSError, SyntaxError) as error:
        raise ValueError("Geçerli bir .xlsx dosyası yükleyin.") from error


def read_workbook(payload):
    book = open_workbook(payload)
    result = {}
    try:
        unexpected = set(book.sheetnames) - set(HEADERS) - {"Kullanım"}
        if unexpected:
            raise ValueError("Tanınmayan sayfa: " + sorted(unexpected)[0])
        for title, columns in HEADERS.items():
            if title not in book.sheetnames:
                continue
            sheet = book[title]
            if (sheet.max_row or 0) > MAX_ROWS + 1 or (sheet.max_column or 0) > 40:
                raise ValueError(f"{title}: en çok {MAX_ROWS} veri satırı yükleyin.")
            records = sheet.iter_rows()
            headings = [cell_text(cell) for cell in next(records, ())]
            while headings and not headings[-1]:
                headings.pop()
            if headings != list(columns.values()):
                raise ValueError(f"{title}, satır 1: şablon sütun başlıklarını ve sırasını koruyun.")
            result[title] = []
            for index, row in enumerate(records, 2):
                if index > MAX_ROWS + 1 or len(row) > 40:
                    raise ValueError(f"{title}: en çok {MAX_ROWS} veri satırı yükleyin.")
                try:
                    values = [cell_text(cell) for cell in row]
                    if any(values[len(columns):]):
                        raise ValueError("Şablon dışında veri sütunu bulundu.")
                    if any(values):
                        result[title].append((index, dict(zip(columns, values + [""] * len(columns)))))
                except ValueError as error:
                    raise ValueError(f"{title}, satır {index}: {error}") from error
    finally:
        book.close()
    if not any(result.values()):
        raise ValueError("İçe aktarılacak veri satırı bulunamadı.")
    return result


def apply_workbook(db, workbook, model):
    counts = {}
    new_products = {}

    def find(table, field, key, label):
        rows = db.execute(f"SELECT * FROM {table} ORDER BY id")
        matches = [row for row in rows if model.normalized(row[field]) == model.normalized(key)]
        if len(matches) != 1 or not key:
            raise ValueError(f"{label} bulunamadı veya eşleşme birden fazla: {key or '(boş)'}")
        return matches[0]

    def product(value):
        code = new_products.get(model.normalized(value[1:]), "") if value.startswith("@") else value
        return find("products", "code", code, "Ürün kodu")

    def choice(value, options, label):
        selected = options.get(model.normalized(value))
        if selected is None:
            raise ValueError(f"{label} için şu değerlerden birini kullanın: {', '.join(options)}")
        return selected

    def apply_row(title, row):
        if title == "Ürün Grupları":
            row["module"] = choice(row["module"], {"sarf": "sarf", "kkd": "kkd"}, "Modül")
            model.add_definition(db, "product-groups", row)
        elif title == "Departmanlar":
            model.add_definition(db, "departments", row)
        elif title == "Personel":
            row["department_id"] = find("departments", "name", row["department"], "Departman")["id"]
            row["active"] = choice(row["active"] or "Çalışıyor", {"çalışıyor": "1", "çalışmıyor": "0"}, "Çalışma durumu")
            model.add_definition(db, "people", row)
        elif title == "Ürün Kartları":
            key = model.normalized(model.field(row, "key", False))
            if key and key in new_products:
                raise ValueError("Kart referansı dosyada birden fazla kez kullanılmış.")
            row["group_id"] = find("product_groups", "prefix", row["prefix"], "Grup öneki")["id"]
            row["scope"] = choice(row["scope"], {"personel": "person", "departman": "department",
                                                 "her ikisi": "both"}, "Teslim kapsamı")
            row["minimum"] = row["minimum"] or "0"
            code = model.add_definition(db, "products", row)
            if key:
                new_products[key] = code
        else:
            row["product_id"] = product(row["product"])["id"]
            row["minimum"] = row["minimum"] or "0"
            model.add_definition(db, "variants", row)

    def at_row(title, index, callback, *args):
        try:
            return callback(*args)
        except sqlite3.IntegrityError as error:
            raise ValueError(f"{title}, satır {index}: mükerrer kayıt veya geçersiz bağlantı.") from error
        except ValueError as error:
            raise ValueError(f"{title}, satır {index}: {error}") from error

    for title in HEADERS:
        if title == "Stok Hareketleri":
            break
        for index, record in workbook.get(title, []):
            at_row(title, index, apply_row, title, record.copy())
            counts[title] = counts.get(title, 0) + 1

    title = "Stok Hareketleri"
    documents = {}
    line_fields = {"product", "size", "color", "qty", "unit_price"}
    for index, row in workbook.get(title, []):
        key = at_row(title, index, model.field, row, "key")
        header = {name: value for name, value in row.items() if name not in line_fields}
        key = model.normalized(key)
        if key not in documents:
            documents[key] = (index, header, [])
        elif documents[key][1] != header:
            raise ValueError(f"{title}, satır {index}: aynı fişin üst bilgileri her satırda aynı olmalı.")
        documents[key][2].append((index, row))

    def add_movement(key, header, lines):
        request_id = "excel:" + key
        if db.execute("SELECT 1 FROM documents WHERE request_id=?", (request_id,)).fetchone():
            raise ValueError(f"Bu fiş referansı daha önce aktarılmış: {header['key']}")
        kind = choice(header["kind"], {
            "alış": "receipt", "personel teslimi": "person", "departman teslimi": "department",
            "iade": "return", "sayım": "adjustment",
        }, "İşlem türü")
        payload = {**header, "kind": kind, "request_id": request_id, "lines": []}
        if kind in {"person", "department"}:
            payload["person_id"] = find("people", "code", header["person"], "Personel sicili")["id"]
        if kind == "department":
            payload["department_id"] = find("departments", "name", header["department"], "Departman")["id"]
        if kind == "return":
            source = header["source"]
            if source.startswith("@"):
                row = db.execute("SELECT id FROM documents WHERE request_id=?",
                                 ("excel:" + model.normalized(source[1:]),)).fetchone()
                if row is None:
                    raise ValueError("Kaynak Excel teslim fişi bulunamadı. Fiş sırasını kontrol edin.")
                payload["source_id"] = row["id"]
            else:
                payload["source_id"] = source.removeprefix("#")
        for index, row in lines:
            try:
                product_id = product(row["product"])["id"]
                variants = db.execute("SELECT id,size,color FROM variants WHERE product_id=?", (product_id,))
                variant = next((v for v in variants if model.normalized(v["color"]) == model.normalized(row["color"])
                                and model.normalized(v["size"]) == model.normalized(row["size"])), None)
                if variant is None:
                    raise ValueError("Ürün için belirtilen renk / beden bulunamadı.")
                model.number(row["qty"], 1000, kind == "adjustment")
                if kind == "receipt" or row["unit_price"]:
                    model.number(row["unit_price"], 10000, True)
            except ValueError as error:
                raise ValueError(f"Fiş satırı {index}: {error}") from error
            payload["lines"].append({"variant_id": variant["id"], "qty": row["qty"], "unit_price": row["unit_price"]})
        model.add_document(db, payload)

    for key, (index, header, lines) in sorted(documents.items(), key=lambda item: (item[1][1]["day"], item[1][0])):
        at_row(title, index, add_movement, key, header, lines)
        counts[title] = counts.get(title, 0) + len(lines)
    return {"counts": counts, "documents": len(documents), "codes": list(new_products.items())}
