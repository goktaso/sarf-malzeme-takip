"""Business operations for the Excel version. No database connections are used."""

import json
import re
from copy import deepcopy
from datetime import datetime, timezone
from decimal import Decimal

import access
from server import (BASELINE, ISTANBUL, day_field, field, infer_size_kind, normalized,
                    number, rounded_ratio)
from excel_vba.file_store import insert

MODULES = {"sarf": "Sarf Malzemeleri", "kkd": "İş Kıyafeti ve KKD"}
SCOPES = {"person": "Personel", "department": "Departman", "both": "Her ikisi"}
KINDS = {"receipt": "Alış girişi", "person": "Personele teslim", "department": "Departmana teslim",
         "return": "İade", "adjustment": "Sayım düzeltmesi"}
COUNT_STATUS = {"draft": "Taslak", "locked": "Kilitli", "rejected": "Onaylanmadı", "approved": "Onaylandı"}
PERMISSIONS = dict(access.PERMISSIONS)


def today():
    return datetime.now(ISTANBUL).date().isoformat()


def timestamp():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def allow(user, key):
    if not (user["is_admin"] or key in user["permissions"]):
        raise access.Forbidden("Bu işlem için yetkiniz yok.")


def is_allowed(user, key):
    return user["is_admin"] or key in user["permissions"]


def global_only(user):
    if not user["is_admin"] and user["department_id"] is not None:
        raise access.Forbidden("Bu işlem şirket geneli yetki gerektirir.")


def admin_only(user):
    access.require_admin(user)


def period(state, user, day):
    import server

    if user["is_admin"] or day >= access.open_period_start(server):
        return
    if not any(row["sid"] == user["sid"] and not row["revoked_at"] and row["start_day"] <= day <= row["end_day"]
               for row in state["periods"]):
        raise access.Forbidden("Bu tarih dönem kilidi kapsamında.")


def get(state, table, key):
    if not isinstance(key, (int, str)) or not re.fullmatch(r"[0-9]{1,15}", str(key)):
        raise ValueError("Geçerli bir kayıt seçin.")
    return next((row for row in state[table] if row["id"] == int(key)), None) or missing()


def missing():
    raise ValueError("Seçilen kayıt bulunamadı.")


def match(state, table, column, value):
    if not value:
        raise ValueError("Bağlantı alanı boş.")
    found = [row for row in state[table] if normalized(row[column]) == normalized(value)]
    if len(found) != 1:
        raise ValueError(f"Eşleşmeyen veya mükerrer kayıt: {value}")
    return found[0]


def audit(state, user, action, reference=""):
    insert(state, "audit", account=user["account"], action=action, reference=str(reference), time=timestamp())


def choice(value, options, message):
    if value not in options:
        raise ValueError(message)
    return value


def version(state, table, row, effective):
    earlier = [item for item in state["versions"] if item["entity"] == table and item["entity_id"] == row["id"]]
    current = deepcopy(row)
    if table == "products":
        current["variants"] = [deepcopy(v) for v in state["variants"] if v["product_id"] == row["id"]]
    if table == "groups":
        current.pop("next_number", None)
    last = earlier[-1] if earlier else None
    if last:
        if effective < last["effective_day"]:
            raise ValueError(f"Geçerlilik tarihi son değişiklikten ({last['effective_day']}) önce olamaz.")
        if last["state"] == current:
            return
    insert(state, "versions", entity=table, entity_id=row["id"], effective_day=effective,
           state=current, recorded_at=timestamp())


def on_day(state, table, key, day):
    return next((entry["state"] for entry in reversed(state["versions"])
                 if entry["entity"] == table and entry["entity_id"] == int(key)
                 and entry["effective_day"] <= day), None) or get(state, table, key)


def effective_day(state, user, payload):
    day = day_field(payload, "effective_day") if payload.get("effective_day") else today()
    if day > today() or day == BASELINE:
        raise ValueError("Geçerlilik tarihi bugün veya öncesi olmalı.")
    period(state, user, day)
    return day


def definition(state, user, action, data):
    table = {"group":"groups", "product":"products", "variant":"products",
             "department":"departments", "person":"people"}[action.split("-")[0]]
    permission = {"groups":"edit.products", "products":"edit.products", "departments":"edit.departments",
                  "people":"edit.people"}[table]
    allow(user, permission)
    global_only(user)
    editing = action.endswith("-edit")
    day = effective_day(state, user, data) if editing or action == "variant" else BASELINE
    row = get(state, table, data.get("id")) if editing else None
    if action == "variant":
        row = get(state, "products", data.get("product_id"))
        if any(not v["size"] and not v["color"] for v in state["variants"] if v["product_id"] == row["id"]):
            raise ValueError("Bu ürün tek stok satırıyla açılmış.")
        size, color = field(data,"size",False), field(data,"color",False)
        if not size and not color:
            raise ValueError("Renk veya beden girin.")
        add_variant(state, row, size, color, data.get("minimum") or "0")
    elif table == "groups":
        name = row["name"] if row else field(data,"name")
        prefix = row["prefix"] if row else field(data,"prefix").upper()
        size_kind = choice(data.get("size_kind") or infer_size_kind(name, prefix), ("size","shoe"), "Ölçü türü seçin.")
        if row:
            row["size_kind"] = size_kind
        else:
            module = choice(data.get("module"), MODULES, "Modül seçin.")
            if not re.fullmatch("[A-Z]{2,8}", prefix) or any(g["prefix"] == prefix or
                 (g["module"] == module and normalized(g["name"]) == normalized(name)) for g in state["groups"]):
                raise ValueError("Geçerli ve benzersiz grup adı ve öneki kullanın.")
            row = insert(state, table, name=name, prefix=prefix, size_kind=size_kind, module=module, next_number=1)
    elif table == "products":
        name, unit = field(data,"name"), field(data,"unit")
        scope = choice(data.get("scope"),SCOPES,"Teslim kapsamı seçin.")
        group = get(state,"groups",row["group_id"] if row else data.get("group_id"))
        if any(p["group_id"]==group["id"] and normalized(p["name"])==normalized(name) and
               (not row or p["id"]!=row["id"]) for p in state["products"]):
            raise ValueError("Aynı grupta aynı isimde ürün var.")
        if row:
            if unit != row["unit"] and any(get(state,"variants",m["variant_id"])["product_id"]==row["id"] for m in state["movements"]):
                raise ValueError("Hareketi olan ürünün stok birimi değiştirilemez.")
            row.update(name=name, unit=unit, scope=scope, active=flag(data.get("active")))
        else:
            code=f"{group['prefix']}-{group['next_number']:04d}"
            group["next_number"] += 1
            row=insert(state,table,code=code,name=name,unit=unit,scope=scope,active=1,group_id=group["id"],
                       module=group["module"],category=group["name"])
            sizes=[s.strip() for s in field(data,"sizes",False).split(",") if s.strip()] or [""]
            color=field(data,"color",False)
            for size in sizes:
                add_variant(state,row,size,color,data.get("minimum") or "0")
    else:
        name = field(data,"name")
        if table == "departments":
            if any(normalized(other["name"])==normalized(name) and (not row or other["id"]!=row["id"])
                   for other in state[table]):
                raise ValueError("Bu departman adı kayıtlı.")
            if row:
                row.update(name=name,active=flag(data.get("active")))
            else:
                row=insert(state,table,name=name,active=1)
        else:
            department=get(state,"departments",data.get("department_id"))
            if not department["active"] and (not row or row["department_id"]!=department["id"]):
                raise ValueError("Aktif departman seçin.")
            if row:
                row.update(name=name,department_id=department["id"],active=flag(data.get("active")))
            else:
                code=field(data,"code").upper()
                if any(person["code"]==code for person in state[table]):
                    raise ValueError("Personel sicili kayıtlı.")
                row=insert(state,table,name=name,code=code,department_id=department["id"],active=flag(data.get("active","1")))
    version(state,table,row,day)
    audit(state,user,action,row.get("code",row["id"]))
    return row


def flag(value):
    return int(choice(str(value),("0","1"),"Durum seçin."))


def add_variant(state, product, size, color, minimum):
    if any(v["product_id"]==product["id"] and normalized(v["size"])==normalized(size) and
           normalized(v["color"])==normalized(color) for v in state["variants"]):
        raise ValueError("Bu renk ve beden zaten tanımlı.")
    return insert(state,"variants",product_id=product["id"],size=size,color=color,minimum=number(minimum,1000,True))


def movement_totals(state, variant_id, end=None):
    lines = [line for line in state["movements"] if line["variant_id"] == variant_id and
             (not end or line["day"] <= end)]
    return sum(x["qty"] for x in lines), sum(x["value"] for x in lines), max((x["day"] for x in lines),default="")


def record_document(state, user, kind, data, internal=False):
    day=day_field(data,"day")
    period(state,user,day)
    if day > today():
        raise ValueError("İşlem tarihi bugün veya öncesi olmalı.")
    if not internal:
        allow(user,access.MOVEMENTS[kind])
    department=person=source=None
    if kind in ("person","department"):
        person=on_day(state,"people",get(state,"people",data.get("person_id"))["id"],day)
        home=on_day(state,"departments",person["department_id"],day)
        department=home if kind=="person" else on_day(state,"departments",get(state,"departments",data.get("department_id"))["id"],day)
        if not (person["active"] and home["active"] and department["active"]):
            raise ValueError("Çalışan personel ve aktif departman seçin.")
        scope=user["department_id"]
        if not user["is_admin"] and scope is not None and (department["id"]!=scope or home["id"]!=scope):
            raise access.Forbidden("Departman yetki kapsamınız dışında.")
    if kind=="return":
        source=get(state,"documents",data.get("source_id"))
        period(state,user,source["day"])
        if source["kind"] not in ("person","department") or source["day"]>day:
            raise ValueError("Geçerli bir teslim fişi seçin.")
        department=get(state,"departments",source["department_id"])
        person=get(state,"people",source["person_id"])
        if not user["is_admin"] and user["department_id"] is not None and user["department_id"]!=department["id"]:
            raise access.Forbidden("Departman yetki kapsamınız dışında.")
    purchase_day=supplier=reference=""
    if kind=="receipt":
        global_only(user)
        allow(user,"field.cost")
        allow(user,"field.supplier")
        purchase_day=day_field(data,"purchase_day")
        period(state,user,purchase_day)
        if purchase_day>day:
            raise ValueError("Alış tarihi işlem tarihinden sonra olamaz.")
        supplier,reference=field(data,"supplier"),field(data,"reference")
    note=field(data,"note",False)
    if note and not internal:
        allow(user,"field.notes")
    line_data=data.get("lines") or [data]
    if not isinstance(line_data,list) or not 1<=len(line_data)<=100:
        raise ValueError("Fişe 1–100 satır girin.")
    request_id=data.get("request_id")
    if request_id:
        earlier=next((doc for doc in state["documents"] if doc.get("request_id")==request_id),None)
        if earlier:
            if earlier["actor"] != user["account"]:
                raise access.Forbidden("Bu işlem anahtarı başka kullanıcıya ait.")
            return earlier
    doc=insert(state,"documents",kind=kind,day=day,purchase_day=purchase_day,supplier=supplier,
               reference=reference,note=note,department_id=department["id"] if department else None,
               person_id=person["id"] if person else None,source_id=source["id"] if source else None,
               actor=user["account"],request_id=request_id)
    seen=set()
    for line in line_data:
        variant=get(state,"variants",line.get("variant_id"))
        vid=variant["id"]
        if vid in seen:
            raise ValueError("Aynı varyant fişte bir kez kullanılabilir.")
        seen.add(vid)
        product=on_day(state,"products",variant["product_id"],day)
        if not product["active"] and kind not in ("return","adjustment"):
            raise ValueError(f"{product['code']} pasif.")
        if kind in ("person","department") and product["scope"] not in (kind,"both"):
            raise ValueError(f"{product['code']} için bu teslim türü tanımlı değil.")
        stock,value,last=movement_totals(state,vid)
        if last and day<last:
            raise ValueError(f"{product['code']} için son hareket tarihi {last}.")
        qty=number(line.get("qty"),1000,kind=="adjustment")
        price=None
        if kind=="adjustment":
            qty-=stock
        if kind=="receipt" or (kind=="adjustment" and qty>0):
            price=number(line.get("unit_price"),10000,True)
            movement_value=rounded_ratio(price,qty,100000)
        elif kind=="return":
            original=next((m for m in state["movements"] if m["document_id"]==source["id"] and m["variant_id"]==vid),None)
            if not original:
                raise ValueError("Varyant kaynak teslim fişinde bulunamadı.")
            returned=[m for m in state["movements"] if m["source_id"]==source["id"] and m["variant_id"]==vid]
            net=qty+sum(m["qty"] for m in returned)
            if net>-original["qty"]:
                raise ValueError("İade miktarı kaynak teslimi aşıyor.")
            movement_value=rounded_ratio(-original["value"],net,-original["qty"])-sum(m["value"] for m in returned)
        else:
            if kind!="adjustment":
                qty=-qty
            if stock+qty<0:
                raise ValueError(f"{product['code']} için yetersiz stok.")
            movement_value=rounded_ratio(value,qty,stock) if qty else 0
        insert(state,"movements",document_id=doc["id"],variant_id=vid,qty=qty,value=movement_value,
               day=day,unit_price=price,source_id=doc["source_id"])
    audit(state,user,kind,doc["id"])
    return doc


def count_operation(state,user,action,data):
    allow(user,"view.sayimlar")
    global_only(user)
    allow(user,"count.approve" if action in ("count-approve","count-reject") else "count.edit")
    if action=="count-new":
        day=day_field(data,"day")
        period(state,user,day)
        if day>today():
            raise ValueError("Sayım tarihi bugün veya öncesi olmalı.")
        lines=[]
        for variant in state["variants"]:
            stock,_,_=movement_totals(state,variant["id"],day)
            product=on_day(state,"products",variant["product_id"],day)
            if stock or product["active"]:
                lines.append({"variant_id":variant["id"],"expected_qty":stock,"counted_qty":None,"unit_price":None})
        result=insert(state,"counts",day=day,name=field(data,"name",False),note=field(data,"note",False),
                      owner=user["account"],status="draft",revision=0,lines=lines,documents=[])
    else:
        result=get(state,"counts",data.get("id"))
        period(state,user,result["day"])
        if result["status"]=="approved" or (result["status"]!="draft" and not user["is_admin"]):
            raise access.Forbidden("Sayım kilitli veya onaylanmış.")
        if str(result["revision"])!=str(data.get("revision")):
            raise ValueError("Sayım değişti. Bilgileri yenileyin.")
        if action in ("count-line","count-delete"):
            vid=get(state,"variants",data.get("variant_id"))["id"]
            line=next((line for line in result["lines"] if line["variant_id"]==vid),None)
            if action=="count-delete":
                if not line:
                    missing()
                result["lines"].remove(line)
            else:
                if line is None:
                    line={"variant_id":vid,"expected_qty":movement_totals(state,vid,result["day"])[0],"unit_price":None}
                    result["lines"].append(line)
                supplied=data.get("qty","")
                line["counted_qty"]=number(supplied,1000,True) if supplied!="" else None
                if data.get("unit_price") is not None:
                    allow(user,"field.cost")
                    line["unit_price"]=number(data["unit_price"],10000,True) if data["unit_price"]!="" else None
        elif action=="count-refresh":
            for line in result["lines"]:
                line["expected_qty"]=movement_totals(state,line["variant_id"],result["day"])[0]
        elif action=="count-lock" and result["status"]=="draft":
            result["status"]="locked"
        elif action=="count-unlock" and result["status"] in ("locked","rejected"):
            admin_only(user)
            result["status"]="draft"
        elif action=="count-reject" and result["status"] in ("draft","locked"):
            result["status"]="rejected"
        elif action=="count-approve" and result["status"] in ("draft","locked"):
            approve_count(state,user,result)
        else:
            raise ValueError("Bu sayım durumunda işlem yapılamaz.")
        result["revision"]+=1
    audit(state,user,action,result["id"])
    return result


def approve_count(state,user,count):
    checked=[line for line in count["lines"] if line["counted_qty"] is not None]
    if not checked:
        raise ValueError("Önce sayılan miktar girin.")
    posting=count["day"]
    adjustments=[]
    for line in checked:
        vid=line["variant_id"]
        if movement_totals(state,vid,count["day"])[0]!=line["expected_qty"]:
            raise ValueError("Sayım tarihi bakiyesi değişti. Sistem bakiyelerini yenileyin.")
        delta=line["counted_qty"]-line["expected_qty"]
        if not delta:
            continue
        stock,value,latest=movement_totals(state,vid)
        posting=max(posting,latest)
        if stock+delta<0:
            raise ValueError("Sayım farkı güncel stoğu eksiye düşürüyor.")
        row={"variant_id":vid,"qty":str(Decimal(stock+delta)/1000)}
        if delta>0:
            price=line["unit_price"]
            if price is None and stock>0:
                price=rounded_ratio(value,100000,stock)
            if price is None:
                price=next((m["unit_price"] for m in reversed(state["movements"]) if m["variant_id"]==vid
                            and m["unit_price"] is not None),None)
            if price is None:
                raise ValueError("Pozitif fark için birim maliyet girin.")
            row["unit_price"]=str(Decimal(price)/10000)
        adjustments.append(row)
    for start in range(0,len(adjustments),100):
        result=record_document(state,user,"adjustment",{"day":posting,
            "note":f"Onaylı sayım #{count['id']} · {count['day']}", "lines":adjustments[start:start+100]},True)
        count["documents"].append(result["id"])
    count["status"]="approved"


def administration(state,user,action,data):
    admin_only(user)
    import server

    if action=="user-save":
        sid,account=access.windows_account(field(data,"account"))
        is_admin,active=flag(data.get("is_admin")),flag(data.get("active"))
        if user["sid"]==sid and (not active or not is_admin):
            raise ValueError("Kendi yönetici hesabınızı kapatamazsınız.")
        permissions=sorted(set(filter(None,data.get("permissions","").split(","))))
        if any(key not in PERMISSIONS for key in permissions):
            raise ValueError("Geçersiz yetki.")
        scope=data.get("department_id")
        info=dict(account=account,display_name=field(data,"name"),is_admin=is_admin,active=active,
                  permissions=permissions,department_id=get(state,"departments",scope)["id"] if scope and not is_admin else None)
        row=next((u for u in state["users"] if u["sid"]==sid),None)
        if row:
            row.update(info)
        else:
            row=insert(state,"users",sid=sid,**info)
    elif action=="period-open":
        target=get(state,"users",data.get("user_id"))
        if not target["active"] or target["is_admin"]:
            raise ValueError("Aktif ve yönetici olmayan kullanıcı seçin.")
        start,end=day_field(data,"start_day"),day_field(data,"end_day")
        if start>end or end>=access.open_period_start(server):
            raise ValueError("Geçerli bir kapalı tarih aralığı seçin.")
        if any(p["sid"]==target["sid"] and not p["revoked_at"] and p["start_day"]<=end and p["end_day"]>=start
               for p in state["periods"]):
            raise ValueError("Aynı kullanıcı için çakışan istisna var.")
        row=insert(state,"periods",sid=target["sid"],account=target["account"],start_day=start,end_day=end,
                   granted_at=timestamp(),granted_by=user["account"],revoked_at=None,revoked_by="",
                   reason=field(data,"note",False))
    else:
        row=get(state,"periods",data.get("id"))
        if row["revoked_at"]:
            raise ValueError("Bu tarih izni zaten kapanmış.")
        row.update(revoked_at=timestamp(),revoked_by=user["account"])
    audit(state,user,action,row["id"])
    return row


def import_workbook(state, user, payload):
    import excel_import

    admin_only(user)
    workbook=excel_import.read_workbook(payload)
    codes={}
    def product_key(value):
        return codes.get(normalized(value[1:])) if value.startswith("@") else value

    def variant(data):
        product=match(state,"products","code",product_key(data["product"]))
        return next((row for row in state["variants"] if row["product_id"]==product["id"] and
                     normalized(row["size"])==normalized(data["size"]) and normalized(row["color"])==normalized(data["color"])),None) or missing()

    for title, records in workbook.items():
        if title=="Stok Hareketleri":
            continue
        for index, data in records:
            try:
                if title=="Ürün Grupları":
                    definition(state,user,"group",data | {"module":normalized(data["module"])})
                elif title=="Departmanlar":
                    definition(state,user,"department",data)
                elif title=="Personel":
                    dep=match(state,"departments","name",data["department"])
                    active=choice(normalized(data["active"]) or "çalışıyor",("çalışıyor","çalışmıyor"),"Çalışma durumu geçersiz.")
                    definition(state,user,"person",data|{"department_id":dep["id"],"active":"1" if active=="çalışıyor" else "0"})
                elif title=="Ürün Kartları":
                    group=match(state,"groups","prefix",data["prefix"])
                    scope=next((key for key,label in SCOPES.items() if normalized(label)==normalized(data["scope"])),"")
                    item=definition(state,user,"product",data|{"group_id":group["id"],"scope":scope})
                    if data["key"]:
                        key=normalized(data["key"])
                        if key in codes:
                            raise ValueError("Kart referansı tekrarlanmış.")
                        codes[key]=item["code"]
                else:
                    product=match(state,"products","code",product_key(data["product"]))
                    definition(state,user,"variant",data|{"product_id":product["id"]})
            except (ValueError,access.Forbidden) as error:
                raise ValueError(f"{title}, satır {index}: {error}") from error
    documents={}
    line_fields={"product","color","size","qty","unit_price"}
    for index,data in workbook.get("Stok Hareketleri",[]):
        try:
            key=normalized(field(data,"key"))
            header={name:value for name,value in data.items() if name not in line_fields}
            first,stored,lines=documents.setdefault(key,(index,header,[]))
            if header!=stored:
                raise ValueError("Aynı fişin üst bilgileri farklı.")
            lines.append({"variant_id":variant(data)["id"],"qty":data["qty"],"unit_price":data["unit_price"]})
        except (ValueError,access.Forbidden) as error:
            raise ValueError(f"Stok Hareketleri, satır {index}: {error}") from error
    kinds={"alış":"receipt","personel teslimi":"person","departman teslimi":"department",
           "iade":"return","sayım":"adjustment"}
    for key,(index,header,lines) in sorted(documents.items(),key=lambda item:(item[1][1]["day"],item[1][0])):
        try:
            request_id="excel:"+key
            if any(doc.get("request_id")==request_id for doc in state["documents"]):
                raise ValueError("Bu fiş referansı daha önce aktarılmış.")
            kind=kinds.get(normalized(header["kind"]))
            if not kind:
                raise ValueError("Geçersiz hareket türü.")
            args=header|{"request_id":request_id,"lines":lines}
            if kind in ("person","department"):
                args["person_id"]=match(state,"people","code",header["person"])["id"]
            if kind=="department":
                args["department_id"]=match(state,"departments","name",header["department"])["id"]
            if kind=="return":
                source=header["source"]
                if source.startswith("@"):
                    original=next((doc for doc in state["documents"] if doc.get("request_id")=="excel:"+normalized(source[1:])),None)
                    if not original:
                        raise ValueError("Kaynak Excel fişi bulunamadı.")
                    args["source_id"]=original["id"]
                else:
                    args["source_id"]=source.removeprefix("#")
            record_document(state,user,kind,args,kind=="adjustment")
        except (ValueError,access.Forbidden) as error:
            raise ValueError(f"Stok Hareketleri, satır {index}: {error}") from error
    return sum(map(len,workbook.values()))
