"""Create sample records in the separate demo database once."""

from datetime import datetime, timedelta, timezone

import server


def seed():
    server.DB = server.ROOT / "data" / "demo.sqlite3"
    server.initialize()
    today = datetime.now(timezone(timedelta(hours=3))).date()
    months = [today.replace(day=1)]
    for _ in range(2):
        months.insert(0, (months[0] - timedelta(days=1)).replace(day=1))

    with server.connect() as db:
        db.execute("BEGIN IMMEDIATE")
        if db.execute("SELECT 1 FROM documents WHERE request_id='demo-opening'").fetchone():
            return

        def define(route, **fields):
            return server.add_definition(db, route, fields)

        def lookup(table, column, value):
            return db.execute(f"SELECT id FROM {table} WHERE {column}=?", (value,)).fetchone()["id"]

        for department in ["Üretim", "Bakım", "İdari İşler"]:
            define("departments", name=department)
        production = lookup("departments", "name", "Üretim")
        maintenance = lookup("departments", "name", "Bakım")
        office = lookup("departments", "name", "İdari İşler")
        for code, name, department in [
            ("D001", "Deniz Kaya", production), ("D002", "Ece Aydın", production),
            ("D003", "Mert Akın", maintenance), ("D004", "Selin Demir", office),
        ]:
            define("people", code=code, name=name, department_id=department)
        deniz, ece, mert, selin = [lookup("people", "code", f"D00{i}") for i in range(1, 5)]

        for prefix, name, module in [
            ("MTF", "Mutfak sarfı", "sarf"), ("TMZ", "Temizlik sarfı", "sarf"),
            ("TSH", "Tişört", "kkd"), ("AYK", "İş ayakkabısı", "kkd"),
            ("ELD", "Eldiven", "kkd"), ("GZL", "Koruyucu gözlük", "kkd"),
        ]:
            define("product-groups", name=name, prefix=prefix, module=module)

        def product(prefix, name, unit, scope, minimum, sizes="", color=""):
            code = define("products", group_id=lookup("product_groups", "prefix", prefix),
                          name=name, unit=unit, scope=scope, minimum=minimum, sizes=sizes, color=color)
            return lookup("products", "code", code)

        def variant(product_id, size="", color=""):
            return db.execute("SELECT id FROM variants WHERE product_id=? AND size=? AND color=?",
                              (product_id, size, color)).fetchone()["id"]

        tea = variant(product("MTF", "Siyah çay", "kg", "department", 5))
        sugar = variant(product("MTF", "Toz şeker", "kg", "department", 3))
        coffee = variant(product("MTF", "Filtre kahve", "kg", "department", 2))
        tissue = variant(product("TMZ", "Tuvalet kâğıdı", "rulo", "department", 20))
        napkin = variant(product("TMZ", "Servis peçetesi", "paket", "department", 8))
        shirt = product("TSH", "Polo yaka iş tişörtü", "adet", "person", 2, "M,L", "Lacivert")
        define("variants", product_id=shirt, size="M", color="Siyah", minimum=2)
        navy_m = variant(shirt, "M", "Lacivert")
        navy_l = variant(shirt, "L", "Lacivert")
        black_m = variant(shirt, "M", "Siyah")
        shoes = product("AYK", "Çelik burunlu iş ayakkabısı", "çift", "person", 1, "41,42,43")
        shoe_41, shoe_42, shoe_43 = [variant(shoes, str(size)) for size in (41, 42, 43)]
        gloves = product("ELD", "Nitril kaplı iş eldiveni", "çift", "both", 15, "M,L")
        glove_m, glove_l = [variant(gloves, size) for size in ("M", "L")]
        glasses = variant(product("GZL", "Şeffaf koruyucu gözlük", "adet", "both", 5))

        def document(key, kind, day, quantities, **fields):
            lines = []
            for item in quantities:
                variant_id, qty, *price = item
                line = {"variant_id": variant_id, "qty": str(qty)}
                if price:
                    line["unit_price"] = str(price[0])
                lines.append(line)
            return server.add_document(db, dict(request_id=f"demo-{key}", kind=kind,
                                               day=day.isoformat(), lines=lines, **fields))

        def receipt(key, day, supplier, quantities):
            return document(key, "receipt", day, quantities, supplier=supplier,
                            purchase_day=day.isoformat(), reference=f"DEMO-{key.upper()}")

        oldest, previous, current = months
        receipt("opening", oldest, "Örnek Genel Tedarik", [
            (tea, 18, 180), (sugar, 14, 36), (coffee, 6, 480), (tissue, 96, 14),
            (napkin, 24, 45), (navy_m, 15, 240), (navy_l, 8, 240), (black_m, 4, 260),
            (shoe_41, 2, 1250), (shoe_42, 5, 1250), (shoe_43, 1, 1250),
            (glove_m, 100, 40), (glove_l, 55, 42), (glasses, 24, 80),
        ])
        document("office-old", "department", oldest, [(tea, 4), (sugar, 3), (tissue, 36)],
                 person_id=selin, department_id=office)
        document("production-old", "department", oldest, [(glove_m, 40), (glasses, 8)],
                 person_id=deniz, department_id=production)

        receipt("second", previous, "Örnek KKD ve Sarf", [
            (tea, 10, 220), (glove_m, 40, 52), (glasses, 10, 95), (shoe_42, 2, 1450),
        ])
        document("personal-kit", "person", previous, [(navy_m, 2), (shoe_42, 1), (glove_m, 2)],
                 person_id=ece)
        document("maintenance-old", "department", previous, [(glove_l, 30), (glasses, 5)],
                 person_id=mert, department_id=maintenance)
        document("last-shoe", "person", previous, [(shoe_43, 1), (black_m, 1)], person_id=mert)

        issue = document("production", "department", current, [(glove_m, 50), (glasses, 12)],
                         person_id=deniz, department_id=production)
        document("gloves-return", "return", current, [(glove_m, 8)],
                 source_id=issue, note="Açılmamış paket depoya iade edildi.")
        document("office", "department", current,
                 [(coffee, 3.5), (tea, 6.25), (sugar, 7), (tissue, 42), (napkin, 18)],
                 person_id=selin, department_id=office)
        document("count", "adjustment", current, [(napkin, 5)], note="Fiziksel sayımda bir paket eksik.")
        receipt("latest", today, "Örnek KKD ve Sarf", [(glove_l, 20, 49.5), (navy_l, 6, 285)])


if __name__ == "__main__":
    seed()
    print("Demo veritabani hazir.")
