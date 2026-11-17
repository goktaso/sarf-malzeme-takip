"""Excel screens and native form descriptions."""

from calendar import monthrange
from datetime import date
from decimal import Decimal

import access
from excel_vba import inventory as m

VIEWS = {
    "stock":("Stok Durumu","view.stok"), "movements":("Stok Hareketleri","view.hareketler"),
    "products":("Ürün Kartları","view.urunler"), "groups":("Ürün Grupları","view.urunler"),
    "departments":("Departmanlar","view.tanimlar"), "people":("Personel","view.tanimlar"),
    "report":("Tüketim Raporu","view.raporlar"), "counts":("Sayım Listesi","view.sayimlar"),
    "count-lines":("Sayım Ürünleri","view.sayimlar"), "history":("Tanım Geçmişi",""),
    "users":("Kullanıcılar ve Yetkiler","admin"), "periods":("Dönem Kilidi","admin"), "audit":("İşlem Günlüğü","admin"),
}
ACTIONS = {
    "group":("Ürün grubu ekle","edit.products"), "group-edit":("Grubun ölçü türünü değiştir","edit.products"),
    "product":("Ürün kartı ekle","edit.products"), "product-edit":("Ürün kartını düzenle","edit.products"),
    "variant":("Renk ve beden ekle","edit.products"),
    "department":("Departman ekle","edit.departments"),"department-edit":("Departmanı düzenle","edit.departments"),
    "person":("Personel ekle","edit.people"),"person-edit":("Personeli düzenle","edit.people"),
    "receipt":("Alış girişi","movement.receipt"),"person-issue":("Personel teslimi","movement.delivery"),
    "department-issue":("Departman teslimi","movement.delivery"),"return":("İade al","movement.return"),
    "count-new":("Yeni sayım","count.edit"), "count-line":("Sayım miktarını gir / düzelt","count.edit"),
    "count-delete":("Sayım satırını sil","count.edit"), "count-refresh":("Sayım sistem bakiyelerini yenile","count.edit"),
    "count-lock":("Sayımı kilitle","count.edit"), "count-reject":("Sayımı reddet","count.approve"),
    "count-approve":("Sayımı onayla","count.approve"), "count-unlock":("Sayım kilidini aç","admin"),
    "user-save":("Windows hesabı ve yetkileri","admin"), "period-open":("Tarih aralığı aç","admin"),
    "period-close":("Tarih iznini kapat","admin"),
}


def visible(user,key):
    if key=="admin": return bool(user["is_admin"])
    if key=="view.sayimlar" or key.startswith(("edit.","count.")):
        return user["department_id"] is None and m.is_allowed(user,key)
    if key=="movement.receipt":
        return user["department_id"] is None and all(m.is_allowed(user,p) for p in (key,"field.cost","field.supplier"))
    return not key or m.is_allowed(user,key)


def restricted_department(user):
    return None if user["is_admin"] else user["department_id"]


def people_in_scope(state,user):
    scope=restricted_department(user)
    return [p for p in state["people"] if scope is None or p["department_id"]==scope]


def document_in_scope(doc,user):
    scope=restricted_department(user)
    return scope is None or doc["department_id"]==scope


def active_label(value):
    return "Aktif" if value else "Pasif"


def variant_label(state,row):
    product=m.get(state,"products",row["product_id"])
    return " · ".join(filter(None,(product["code"],product["name"],row["color"],row["size"])))


def qty(value):
    return None if value is None else Decimal(value)/1000


def count_name(count):
    return count.get("name") or f"Sayım #{count['id']}"


def cost(value):
    return Decimal(value)/100


def stock_columns(state,user,args):
    selected=args.get("module")
    columns=["Stok kodu","Ürün","Grup","Renk","Beden","Ayakkabı numarası","Birim","Eldeki miktar"]
    see_cost=m.is_allowed(user,"field.cost")
    rows=[]
    for v in state["variants"]:
        p=m.get(state,"products",v["product_id"])
        if selected and p["module"]!=selected: continue
        group=m.get(state,"groups",p["group_id"])
        stock,value,_=m.movement_totals(state,v["id"])
        shoe=group["size_kind"]=="shoe"
        row=[p["code"],p["name"],p["category"],v["color"],"" if shoe else v["size"],v["size"] if shoe else "",
             p["unit"],qty(stock)]
        if see_cost: row.append(cost(value))
        row.append("Pasif" if not p["active"] else "Stok yok" if stock==0 else "Kritik stok" if stock<=v["minimum"] else "Mevcut")
        rows.append(row)
    return columns+(["Stok değeri"] if see_cost else [])+["Durum"],rows


def ranged_documents(state,user,args):
    start=m.day_field({"day":args.get("from") or m.today()[:7]+"-01"},"day")
    end=m.day_field({"day":args.get("to") or m.today()},"day")
    if start>end:
        raise ValueError("Başlangıç tarihi bitiş tarihinden sonra olamaz.")
    docs={d["id"]:d for d in state["documents"] if start<=d["day"]<=end and document_in_scope(d,user)}
    return start,end,docs


def movements_or_report(state,user,args,report):
    start,end,docs=ranged_documents(state,user,args)
    see_cost=m.is_allowed(user,"field.cost")
    selected=args.get("module")
    rows=[]
    totals={}
    for line in sorted(state["movements"],key=lambda item:(item["day"],item["id"]),reverse=True):
        doc=docs.get(line["document_id"])
        if not doc: continue
        v=m.get(state,"variants",line["variant_id"])
        p=m.on_day(state,"products",v["product_id"],doc["day"])
        if selected and p["module"]!=selected: continue
        dept=m.on_day(state,"departments",doc["department_id"],doc["day"])["name"] if doc["department_id"] else ""
        person=m.on_day(state,"people",doc["person_id"],doc["day"])["name"] if doc["person_id"] else ""
        line_name=" · ".join(filter(None,(p["name"],v["color"],v["size"])))
        if report:
            kind=m.get(state,"documents",doc["source_id"])["kind"] if doc["source_id"] else doc["kind"]
            if kind not in ("person","department"): continue
            key=(doc["day"][:7],dept,m.KINDS[kind],p["code"],line_name,p["unit"],v["id"])
            total=totals.setdefault(key,[0,0])
            total[0]-=line["qty"]
            total[1]-=line["value"]
        else:
            row=[doc["id"],doc["day"],m.KINDS[doc["kind"]],p["code"],line_name,dept,person,p["unit"],qty(line["qty"])]
            if see_cost: row.append(cost(line["value"]))
            row += [doc["reference"],doc["purchase_day"],doc["source_id"]]
            if m.is_allowed(user,"field.supplier"): row.append(doc["supplier"])
            if m.is_allowed(user,"field.notes"): row.append(doc["note"])
            rows.append(row)
    if not report:
        return (["Fiş no","Tarih","İşlem","Stok kodu","Ürün","Departman","Teslim alan","Birim","Miktar"]+
                (["Stok değeri farkı"] if see_cost else [])+["Belge numarası","Alış tarihi","Kaynak fiş"]+
                (["Tedarikçi"] if m.is_allowed(user,"field.supplier") else [])+
                (["Açıklama"] if m.is_allowed(user,"field.notes") else []),rows)
    for (month,dept,kind,code,name,unit,vid),(amount,value) in totals.items():
        row=[month,dept,kind,code,name,unit,qty(amount)]
        if see_cost: row.append(cost(value))
        row.append(f"{month}:{vid}")
        rows.append(row)
    columns=["Ay","Departman","Teslim türü","Stok kodu","Ürün","Birim","Tüketim miktarı"]
    return columns+(["Net tüketim maliyeti"] if see_cost else [])+["_Rapor"],rows


def report_balances(state,user,args):
    m.allow(user,"view.raporlar")
    start,end,docs=ranged_documents(state,user,args)
    allowed=grid(state,user,"report",args)[1]
    possible={row[-1] for row in allowed}
    keys=set(filter(None,args.get("keys","").split(","))) & possible
    filtered=args.get("filtered")=="1"
    first_month,last_month=date.fromisoformat(start),date.fromisoformat(end)
    months=(len({key.split(":")[0] for key in keys}) if filtered else
            (last_month.year-first_month.year)*12+last_month.month-first_month.month+1)
    if restricted_department(user) is not None: return {"months":months}
    variants={int(key.split(":")[1]) for key in keys}
    balance_day=end
    if args.get("month_filter")=="1" and keys:
        month=max(key.split(":")[0] for key in keys)
        part=date.fromisoformat(month+"-01")
        balance_day=min(end,f"{month}-{monthrange(part.year,part.month)[1]:02d}")
    incoming={}
    balance={}
    value=0
    for entry in state["movements"]:
        vid=entry["variant_id"]
        if filtered and vid not in variants: continue
        variant=m.get(state,"variants",vid)
        product=m.get(state,"products",variant["product_id"])
        if args.get("module") and product["module"]!=args["module"]: continue
        unit=product["unit"]
        doc=docs.get(entry["document_id"])
        if doc and doc["kind"]=="receipt" and (not filtered or f"{entry['day'][:7]}:{vid}" in keys):
            incoming[unit]=incoming.get(unit,0)+entry["qty"]
            value+=entry["value"]
        if entry["day"]<=balance_day:
            balance[unit]=balance.get(unit,0)+entry["qty"]
    result={"received":list(incoming.items()),"balance":list(balance.items()),"day":balance_day,"months":months}
    if m.is_allowed(user,"field.cost"): result["value"]=value
    return result


def grid(state,user,name,args):
    label,permission=VIEWS[name]
    if not visible(user,permission):
        raise access.Forbidden("Bu listeye erişim yetkiniz yok.")
    if name=="stock": return stock_columns(state,user,args)
    if name in ("movements","report"): return movements_or_report(state,user,args,name=="report")
    if name=="groups":
        return ["Grup","Önek","Modül","Ölçü türü"],[[g["name"],g["prefix"],m.MODULES[g["module"]],
            "Ayakkabı numarası" if g["size_kind"]=="shoe" else "Beden"] for g in state["groups"]]
    if name=="products":
        return ["Stok kodu","Ürün","Grup","Modül","Birim","Teslim kapsamı","Durum"],[
            [p["code"],p["name"],p["category"],m.MODULES[p["module"]],p["unit"],m.SCOPES[p["scope"]],active_label(p["active"])]
            for p in state["products"] if not args.get("module") or args["module"]==p["module"]]
    if name=="departments":
        return ["Departman","Durum"],[[p["name"],active_label(p["active"])] for p in state["departments"]
            if restricted_department(user) in (None,p["id"])]
    if name=="people":
        selected_day=m.day_field({"day":args.get("to") or m.today()},"day")
        rows=[]
        for item in state["people"]:
            p=m.on_day(state,"people",item["id"],selected_day)
            if restricted_department(user) not in (None,p["department_id"]): continue
            d=m.on_day(state,"departments",p["department_id"],selected_day)
            rows.append([p["code"],p["name"],d["name"],"Çalışıyor" if p["active"] else "Çalışmıyor",selected_day])
        return ["Sicil","Personel","Departman","Çalışma durumu","Tarih"],rows
    if name=="counts":
        return ["Sayım adı","Sayım no","Tarih","Durum","Sayılan satır","Toplam satır","Oluşturan","Açıklama"],[
            [count_name(row),row["id"],row["day"],m.COUNT_STATUS[row["status"]],sum(line["counted_qty"] is not None for line in row["lines"]),
             len(row["lines"]),row["owner"],row["note"]] for row in reversed(state["counts"])]
    if name=="count-lines":
        count=m.get(state,"counts",args.get("id"))
        rows=[]
        for line in count["lines"]:
            v=m.get(state,"variants",line["variant_id"])
            p=m.on_day(state,"products",v["product_id"],count["day"])
            sku=line["counted_qty"]
            rows.append([p["name"],v["color"],v["size"],qty(line["expected_qty"]),qty(sku),
                         qty(sku-line["expected_qty"]) if sku is not None else None,p["unit"],p["code"],v["id"]])
        return ["Ürün","Renk","Beden / numara","Sistem bakiyesi","Sayılan miktar","Fark","Birim","Stok kodu","_Varyant"],rows
    if name=="history":
        rows=[]
        for entry in reversed(state["versions"]):
            table=entry["entity"]
            scope_key=args.get("scope")
            if scope_key=="products" and table not in ("groups","products"): continue
            if scope_key=="people" and table not in ("departments","people"): continue
            required="view.urunler" if table in ("groups","products") else "view.tanimlar"
            if not visible(user,required): continue
            record=entry["state"]
            scope=restricted_department(user)
            if table=="people" and scope is not None and record["department_id"]!=scope: continue
            if table=="departments" and scope is not None and record["id"]!=scope: continue
            labels={"groups":"Ürün grubu","products":"Ürün kartı","departments":"Departman","people":"Personel"}
            common=[labels[table],record.get("code",record.get("prefix",record["name"])),
                    record["name"],entry["effective_day"] if entry["effective_day"]!=m.BASELINE else "Başlangıç",
                    entry["recorded_at"]]
            for key,label in (("active","Durum"),("module","Modül"),("size_kind","Ölçü türü"),("unit","Birim"),
                              ("category","Grup"),("scope","Teslim kapsamı"),("department_id","Departman")):
                if key not in record: continue
                value=record[key]
                if key=="active": value=active_label(value)
                if key=="module": value=m.MODULES[value]
                if key=="scope": value=m.SCOPES[value]
                if key=="department_id": value=m.on_day(state,"departments",value,entry["effective_day"])["name"]
                rows.append(common+[label,value])
            for v in record.get("variants",[]):
                label=" / ".join(filter(None,(v["color"],v["size"]))) or "Bedensiz"
                rows.append(common+["Renk / beden · Kritik stok",f"{label} · {qty(v['minimum'])}"])
        return ["Kayıt türü","Kod","Ad","Geçerlilik","Kayıt zamanı","Alan","Değer"],rows
    if name=="users":
        return ["Windows hesabı","Ad","Rol","Departman","Durum","Yetkiler"],[
            [row["account"],row["display_name"],"Yönetici" if row["is_admin"] else "Kullanıcı",
             m.get(state,"departments",row["department_id"])["name"] if row["department_id"] is not None else "",
             active_label(row["active"]),", ".join(m.PERMISSIONS[key] for key in row["permissions"])]
            for row in state["users"]]
    if name=="periods":
        return ["No","Hesap","Başlangıç","Bitiş","Durum","Açıklama","Açan","Kapatan"],[
            [r["id"],r["account"],r["start_day"],r["end_day"],"Kapalı" if r["revoked_at"] else "Açık",
             r["reason"],r["granted_by"],r["revoked_by"]] for r in reversed(state["periods"])]
    return ["Zaman","Windows hesabı","İşlem","Kayıt"],[[a["time"],a["account"],a["action"],a["reference"]]
        for a in reversed(state["audit"][-500:])]


def form(state,user,action,record=None):
    title,permission=ACTIONS[action]
    if not visible(user,permission):
        raise access.Forbidden("Bu işlem için yetkiniz yok.")
    fields=[]
    def text(key,label,value="",required=False):
        fields.append(dict(key=key,label=label,value=str(value),kind="required" if required else "text"))
    def select(key,label,options,value="",lookup=False):
        fields.append(dict(key=key,label=label,value=str(value),kind="select",options=options,lookup=lookup))
    def edit(table,display):
        if not record: initial={}
        else: initial=m.get(state,table,record)
        select("id","Kayıt",[(r["id"],display(r)) for r in state[table]],record or "",True)
        return initial
    def date_field(key,label):
        text(key,label,m.today(),True)
    def boolean(key,label,value="1"):
        select(key,label,[("1","Evet"),("0","Hayır")],value)
    departments=[(d["id"],d["name"]) for d in state["departments"] if d["active"] and
                 restricted_department(user) in (None,d["id"])]
    products=[(p["id"],p["code"]+" · "+p["name"]) for p in state["products"]]
    variants=[(v["id"],variant_label(state,v)) for v in state["variants"]]
    if action.startswith("group"):
        item=edit("groups",lambda row:row["prefix"]+" · "+row["name"]) if action.endswith("-edit") else {}
        if not action.endswith("-edit"):
            text("name","Grup adı","",True)
            text("prefix","Kod öneki","",True)
            select("module","Modül",list(m.MODULES.items()),"kkd")
        select("size_kind","Ölçü türü",[("size","Beden"),("shoe","Ayakkabı numarası")],item.get("size_kind",""))
    elif action in ("product","product-edit"):
        item=edit("products",lambda row:row["code"]+" · "+row["name"]) if action.endswith("-edit") else {}
        if action=="product": select("group_id","Ürün grubu",[(g["id"],g["prefix"]+" · "+g["name"]) for g in state["groups"]])
        text("name","Ürün adı",item.get("name",""),True)
        text("unit","Stok birimi",item.get("unit","adet"),True)
        select("scope","Teslim kapsamı",list(m.SCOPES.items()),item.get("scope","both"))
        if action=="product":
            text("sizes","Bedenler / numaralar (virgülle)")
            text("color","Renk")
            text("minimum","Kritik stok","0")
        else:
            boolean("active","Aktif",item.get("active","1"))
    elif action=="variant":
        select("product_id","Ürün kartı",products)
        text("color","Renk")
        text("size","Beden / numara")
        text("minimum","Kritik stok","0")
        date_field("effective_day","Geçerlilik tarihi")
    elif action in ("department","department-edit","person","person-edit"):
        table="departments" if action.startswith("department") else "people"
        item=edit(table,lambda row:row["name"]) if action.endswith("-edit") else {}
        text("name","Departman adı" if table=="departments" else "Ad soyad",item.get("name",""),True)
        if table=="people":
            if action=="person": text("code","Sicil no","",True)
            select("department_id","Departman",departments,item.get("department_id",""))
        if action!="department": boolean("active","Çalışıyor" if table=="people" else "Aktif",item.get("active","1"))
    elif action in ("receipt","person-issue","department-issue","return"):
        date_field("day","İşlem tarihi")
        if action=="receipt":
            date_field("purchase_day","Alış tarihi")
            text("supplier","Tedarikçi","",True)
            text("reference","Belge numarası","",True)
            text("unit_price","Birim maliyet TL","",True)
        if action in ("person-issue","department-issue"):
            if action=="department-issue": select("department_id","Departman",departments)
            select("person_id","Teslim alan",[(p["id"],p["code"]+" · "+p["name"]) for p in people_in_scope(state,user) if p["active"]])
        if action=="return":
            select("source_id","Kaynak teslim fişi",[(d["id"],f"#{d['id']} · {d['day']}") for d in reversed(state["documents"])
                  if d["kind"] in ("person","department") and document_in_scope(d,user)])
        select("variant_id","Ürün / renk / beden",variants)
        text("qty","Miktar","",True)
        if m.is_allowed(user,"field.notes"): text("note","Açıklama")
    elif action=="count-new":
        text("name","Sayım adı","",True)
        date_field("day","Sayım tarihi")
        text("note","Açıklama")
    elif action.startswith("count-"):
        options=[(f"{c['id']}:{c['revision']}",f"{count_name(c)} · {c['day']} · {m.COUNT_STATUS[c['status']]}") for c in reversed(state["counts"])
                 if c["status"]!="approved" and (c["status"]=="draft" or user["is_admin"])]
        variant_id=""
        selected_count=record or ""
        if record and record.count(":")==2:
            count_id,revision,variant_id=record.split(":")
            selected_count=f"{count_id}:{revision}"
        select("count_key","Sayım",options,selected_count)
        if action in ("count-line","count-delete"):
            select("variant_id","Ürün / renk / beden",variants,variant_id)
            if action=="count-line":
                current=None
                if variant_id and selected_count in {key for key,_ in options}:
                    count=m.get(state,"counts",selected_count.split(":")[0])
                    current=next((line for line in count["lines"] if str(line["variant_id"])==variant_id),None)
                text("qty","Sayılan miktar (boş = sayılmadı)",
                     qty(current["counted_qty"]) if current and current["counted_qty"] is not None else "")
                if m.is_allowed(user,"field.cost"):
                    text("unit_price","Birim maliyet TL (isteğe bağlı)",
                         Decimal(current["unit_price"])/10000 if current and current["unit_price"] is not None else "")
    elif action=="user-save":
        if record:
            item=m.get(state,"users",record)
        else:
            item={}
        select("existing","Mevcut kullanıcı",[(u["id"],u["account"]) for u in state["users"]],record or "",True)
        text("account","Windows hesabı",item.get("account",""),True)
        text("name","Görünen ad",item.get("display_name",""),True)
        boolean("is_admin","Uygulama yöneticisi",item.get("is_admin",0))
        boolean("active","Aktif",item.get("active",1))
        select("department_id","Departman kapsamı",departments,item.get("department_id") or "")
        fields.append(dict(key="permissions",label="Ekran, işlem ve alan yetkileri",kind="multi",
                           value=",".join(item.get("permissions",[])),options=list(m.PERMISSIONS.items())))
    elif action=="period-open":
        select("user_id","Windows kullanıcısı",[(u["id"],u["account"]) for u in state["users"] if u["active"] and not u["is_admin"]])
        text("start_day","Başlangıç tarihi (YYYY-AA-GG)","",True)
        text("end_day","Bitiş tarihi (YYYY-AA-GG)","",True)
        text("note","Açıklama")
    else:
        select("id","Açık dönem izni",[(p["id"],p["account"]+" · "+p["start_day"]+" / "+p["end_day"])
                                     for p in state["periods"] if not p["revoked_at"]])
    if action.endswith("-edit"):
        date_field("effective_day","Geçerlilik tarihi")
    return title,fields


def apply(state,user,action,args):
    if not visible(user,ACTIONS[action][1]):
        raise access.Forbidden("Bu işlem için yetkiniz yok.")
    if action.startswith("count-"):
        if action!="count-new":
            key=args.pop("count_key","")
            if ":" in key:
                args["id"],args["revision"]=key.split(":",1)
            if args.get("unit_price")=="":
                args.pop("unit_price")
        return m.count_operation(state,user,action,args)
    if action in ("user-save","period-open","period-close"):
        return m.administration(state,user,action,args)
    if action in ("receipt","person-issue","department-issue","return"):
        return m.record_document(state,user,action.removesuffix("-issue"),args)
    return m.definition(state,user,action,args)
