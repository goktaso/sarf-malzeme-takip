"""Windows login, application permissions, and session handling."""

import hashlib
import json
import secrets
import time
from datetime import datetime, timedelta, timezone
from http.cookies import CookieError, SimpleCookie

import sqlserver

SCREENS = {"stok": "Stok durumu", "hareketler": "Stok hareketleri", "urunler": "Ürün kartları",
           "tanimlar": "Departman ve personel", "raporlar": "Tüketim raporları", "sayimlar": "Sayım"}
PERMISSIONS = {
    **{f"view.{key}": name for key, name in SCREENS.items()},
    "edit.products": "Ürün kartı ve grup tanımlama",
    "edit.departments": "Departman tanımlama",
    "edit.people": "Personel tanımlama",
    "movement.receipt": "Alış girişi",
    "movement.delivery": "Malzeme teslimi",
    "movement.return": "İade",
    "movement.adjustment": "Sayım düzeltme",
    "count.edit": "Sayım oluşturma, satır düzenleme ve kilitleme",
    "count.approve": "Açık sayımı onaylama veya reddetme",
    "field.cost": "Birim maliyet ve TL değerleri",
    "field.supplier": "Tedarikçi bilgisi",
    "field.notes": "Fiş açıklaması ve düzeltme gerekçesi",
    "export": "Excel ve PDF çıktısı",
}
COOKIE = "SarfSession"
SESSION_SECONDS = 8 * 60 * 60
DEFINITIONS = {
    "products": "edit.products", "product-update": "edit.products",
    "product-groups": "edit.products", "product-group-size-update": "edit.products", "variants": "edit.products",
    "departments": "edit.departments", "department-update": "edit.departments",
    "people": "edit.people", "person-status": "edit.people", "person-update": "edit.people",
}
MOVEMENTS = {"receipt": "movement.receipt", "person": "movement.delivery", "department": "movement.delivery",
             "return": "movement.return", "adjustment": "movement.adjustment"}


class Forbidden(Exception):
    pass


def permitted(user, permission):
    return bool(user["is_admin"] or permission in json.loads(user["permissions"]))


def require(user, permission):
    if not permitted(user, permission):
        raise Forbidden("Bu işlem için yetkiniz yok.")


def require_admin(user):
    if not user or not user["is_admin"]:
        raise Forbidden("Bu işlem yalnızca uygulama yöneticilerine açıktır.")


def open_period_start(model):
    previous = datetime.now(model.ISTANBUL).date().replace(day=1) - timedelta(days=1)
    return previous.replace(day=1).isoformat()


def check_period(db, user, day, model):
    if not user or user["is_admin"] or day >= open_period_start(model):
        return
    if db.execute("""SELECT 1 FROM app_period_grants
        WHERE sid=? AND revoked_at IS NULL AND start_day<=? AND end_day>=?""",
                  (user["sid"], day, day)).fetchone() is None:
        raise Forbidden("Bu tarih dönem kilidi kapsamında. Yönetici bu tarih için size istisna tanımlayabilir.")


def period_grant_records(db, user, model):
    require_admin(user)
    return model.rows(db, """SELECT g.id,u.account,u.display_name,g.start_day,g.end_day,g.reason,
        g.granted_at,g.revoked_at,a.account granted_by,r.account revoked_by
        FROM app_period_grants g JOIN app_users u ON u.sid=g.sid
        JOIN app_users a ON a.sid=g.granted_by LEFT JOIN app_users r ON r.sid=g.revoked_by
        ORDER BY g.id DESC""")


def user_open_periods(db, user, model):
    return model.rows(db, """SELECT start_day,end_day FROM app_period_grants
        WHERE sid=? AND revoked_at IS NULL ORDER BY start_day""", (user["sid"],))


def change_period_grant(db, actor, route, data, model):
    require_admin(actor)
    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    if route == "period-add":
        account = model.field(data, "account")
        users = list(db.execute("SELECT sid,account FROM app_users WHERE account=? AND active=1 AND is_admin=0",
                                (account,)))
        if len(users) != 1:
            raise ValueError("Aktif bir kullanıcı hesabı seçin.")
        sid = users[0]["sid"]
        start = model.day_field(data, "start_day")
        end = model.day_field(data, "end_day")
        if start > end or end >= open_period_start(model):
            raise ValueError("Aralığın tamamı kilitli dönemde olmalı. Başlangıç bitişten sonra olamaz.")
        if db.execute("""SELECT 1 FROM app_period_grants
            WHERE sid=? AND revoked_at IS NULL AND start_day<=? AND end_day>=?""", (sid, end, start)).fetchone():
            raise ValueError("Bu kullanıcı için tarih aralığı mevcut bir istisnayla kesişiyor.")
        reason = model.field(data, "reason", False)
        cursor = db.execute("""INSERT INTO app_period_grants(sid,start_day,end_day,reason,granted_by,granted_at)
                        VALUES (?,?,?,?,?,?)""", (sid, start, end, reason, actor["sid"], timestamp))
        audit(db, actor, route, f"#{cursor.lastrowid} {account} {start} / {end}")
        return cursor.lastrowid
    grant = model.entity(db, "app_period_grants", data.get("id"))
    if grant["revoked_at"]:
        raise ValueError("Bu istisna zaten kapalı.")
    db.execute("UPDATE app_period_grants SET revoked_by=?,revoked_at=? WHERE id=?",
               (actor["sid"], timestamp, grant["id"]))
    audit(db, actor, route, grant["id"])
    return grant["id"]


def audit(db, user, action, reference=None):
    db.execute("INSERT INTO app_audit(sid,action,record_ref) VALUES (?,?,?)",
               (user["sid"], action, str(reference)[:250] if reference is not None else None))


def windows_account(account):
    sqlserver.load_runtime()
    import win32security
    import pywintypes

    try:
        sid, _, account_type = win32security.LookupAccountName(None, account)
        name, domain, _ = win32security.LookupAccountSid(None, sid)
    except pywintypes.error as error:
        raise ValueError("Windows hesabı doğrulanamadı. ETKİALANI\\kullanıcı biçimini kontrol edin.") from error
    if account_type != 1:
        raise ValueError("Bir Windows kullanıcı hesabı seçin.")
    return win32security.ConvertSidToStringSid(sid), domain + "\\" + name


def token_hash(token):
    return hashlib.sha256(token.encode("ascii")).hexdigest()


def read_cookie(header):
    try:
        cookie = SimpleCookie(header or "").get(COOKIE)
        value = cookie.value if cookie else ""
        return value if len(value) == 43 and value.isascii() else ""
    except CookieError:
        return ""


def authenticate(db, header):
    token = read_cookie(header)
    if not token:
        return None
    return db.execute("""SELECT u.* FROM app_sessions s JOIN app_users u ON u.sid=s.sid
                         WHERE s.token_hash=? AND s.expires_epoch>? AND u.active=1""",
                      (token_hash(token), int(time.time()))).fetchone()


def start_session(db, sid):
    user = db.execute("SELECT * FROM app_users WHERE sid=? AND active=1", (sid,)).fetchone()
    if user is None:
        raise Forbidden("Bu Windows hesabı uygulamada tanımlı ve aktif olmalıdır.")
    token = secrets.token_urlsafe(32)
    now = int(time.time())
    db.execute("DELETE FROM app_sessions WHERE expires_epoch <= ?", (now,))
    db.execute("INSERT INTO app_sessions(token_hash,sid,expires_epoch) VALUES (?,?,?)",
               (token_hash(token), sid, now + SESSION_SECONDS))
    audit(db, user, "login")
    return user, cookie_value(token, SESSION_SECONDS)


def cookie_value(token="", age=0):
    return f"{COOKIE}={token}; Path=/; Max-Age={age}; HttpOnly; SameSite=Strict"


def profile(user):
    return {"account": user["account"], "name": user["display_name"], "admin": bool(user["is_admin"]),
            "department_id": user["department_id"],
            "permissions": list(PERMISSIONS) if user["is_admin"] else json.loads(user["permissions"])}


def state_for_user(state, user):
    def has(permission):
        return permitted(user, permission)

    department = None if user["is_admin"] else user["department_id"]
    if department is not None:
        state["movements"] = [row for row in state["movements"] if row["department_id"] == department]
        state["departments"] = [row for row in state["departments"] if row["id"] == department]
        state["people"] = [row for row in state["people"] if row["department_id"] == department]
    allowed = {
        "departments": has("view.tanimlar") or has("view.raporlar") or has("view.hareketler") or has("movement.delivery"),
        "people": has("view.tanimlar") or has("view.raporlar") or has("view.hareketler") or has("movement.delivery"),
        "products": has("view.urunler"),
        "groups": has("view.urunler"),
        "stock": any(has(key) for key in ("view.stok", "view.urunler", *set(MOVEMENTS.values()))) or
                 (department is None and has("view.sayimlar") and has("count.edit")),
        "movements": has("view.hareketler") or has("view.raporlar") or has("movement.return"),
    }
    for key, visible in allowed.items():
        if not visible:
            state[key] = []
    if not (has("view.hareketler") or has("view.raporlar")) and has("movement.return"):
        state["movements"] = [row for row in state["movements"] if row["kind"] in ("person", "department")]
    if not has("view.hareketler"):
        fields = {"kind", "source_kind", "day", "module", "department_id", "person_id", "department",
                  "variant_id", "name", "color", "size", "qty", "unit", "value"}
        if has("movement.return"):
            fields.update(("document_id", "person"))
        state["movements"] = [{key: value for key, value in row.items() if key in fields}
                              for row in state["movements"]]
    for row in state["movements"]:
        if not has("field.supplier"):
            row.pop("supplier", None)
        if not has("field.notes"):
            row.pop("note", None)
    ids = {"products": {row["id"] for row in state["products"]} | {row["product_id"] for row in state["stock"]},
           "departments": {row["id"] for row in state["departments"]},
           "people": {row["id"] for row in state["people"]}}
    state["definition_timeline"] = [row for row in state["definition_timeline"]
        if row["entity_id"] in ids[row["entity"]] and (department is None or row["entity"] != "people"
                                                        or row["department_id"] == department)]
    if not has("field.cost"):
        for row in state["stock"] + state["movements"]:
            row.pop("value", None)
            row.pop("unit_price", None)
    return state


def check_history(db, user, table, record_id, model):
    if table not in ("people", "departments", "products", "product_groups"):
        raise ValueError("Geçersiz kayıt türü.")
    require(user, "view.tanimlar" if table in ("people", "departments") else "view.urunler")
    if user["department_id"] is not None and not user["is_admin"] and table in ("people", "departments"):
        row = model.entity(db, table, record_id)
        belongs = row["department_id"] if table == "people" else row["id"]
        if belongs != user["department_id"]:
            raise Forbidden("Bu kayıt yetki kapsamınız dışında.")


def check_write(db, user, route, payload, model):
    if route != "documents":
        require(user, DEFINITIONS[route])
        if payload.get("effective_day"):
            check_period(db, user, model.effective_day(payload), model)
        if user["department_id"] is not None and not user["is_admin"]:
            raise Forbidden("Departman kapsamlı hesaplar tanım değiştiremez.")
        return
    kind = payload.get("kind")
    if kind not in tuple(MOVEMENTS):
        raise ValueError("Geçersiz hareket türü.")
    require(user, MOVEMENTS[kind])
    check_period(db, user, model.day_field(payload, "day"), model)
    if kind == "receipt":
        check_period(db, user, model.day_field(payload, "purchase_day"), model)
    if kind == "return" and not user["is_admin"]:
        check_period(db, user, model.entity(db, "documents", payload.get("source_id"))["day"], model)
    if kind == "adjustment" and not user["is_admin"]:
        raise Forbidden("Stok sayım farkını Sayım ekranında onaylayarak kaydedin.")
    if kind in ("receipt", "adjustment"):
        require(user, "field.cost")
    if kind == "receipt":
        require(user, "field.supplier")
    if kind == "adjustment" or payload.get("note"):
        require(user, "field.notes")
    department = user["department_id"]
    if department is None or user["is_admin"]:
        return
    if kind in ("receipt", "adjustment"):
        raise Forbidden("Bu işlem şirket geneli stok yetkisi gerektirir.")
    if kind == "return":
        target = model.entity(db, "documents", payload.get("source_id"))["department_id"]
    elif kind == "person":
        target = model.definition_on(db, "people", payload.get("person_id"), model.day_field(payload, "day"))["department_id"]
    else:
        target = model.entity(db, "departments", payload.get("department_id"))["id"]
        recipient = model.definition_on(db, "people", payload.get("person_id"), model.day_field(payload, "day"))
        if recipient["department_id"] != department:
            raise Forbidden("Teslim alan kişi yetki kapsamınız dışında.")
    if target != department:
        raise Forbidden("Bu departmanda işlem yetkiniz yok.")


def save_user(db, actor, payload, model):
    require_admin(actor)
    sid, account = windows_account(model.field(payload, "account"))
    display = model.field(payload, "name")
    active = model.active_field(payload)
    admin = int(payload.get("admin") is True)
    permissions = payload.get("permissions")
    if not isinstance(permissions, list) or any(not isinstance(key, str) or key not in PERMISSIONS for key in permissions):
        raise ValueError("Geçersiz kullanıcı yetkisi.")
    department = payload.get("department_id")
    department = None if not department or admin else model.entity(db, "departments", department)["id"]
    if sid == actor["sid"] and (not admin or not active):
        raise ValueError("Kendi yönetici hesabınızı devre dışı bırakamazsınız.")
    existing = db.execute("SELECT sid FROM app_users WHERE sid=?", (sid,)).fetchone()
    values = (account, display, admin, active, json.dumps(sorted(set(permissions))), department, sid)
    if existing:
        db.execute("UPDATE app_users SET account=?,display_name=?,is_admin=?,active=?,permissions=?,department_id=? WHERE sid=?", values)
    else:
        db.execute("INSERT INTO app_users(account,display_name,is_admin,active,permissions,department_id,sid) VALUES (?,?,?,?,?,?,?)", values)
    if not active:
        db.execute("DELETE FROM app_sessions WHERE sid=?", (sid,))
    audit(db, actor, "user-update", account)
