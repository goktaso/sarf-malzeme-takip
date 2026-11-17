from datetime import date, timedelta

from excel_vba import inventory as m


def seed_sample(state,user):
    today=date.fromisoformat(m.today())
    previous=today.replace(day=1)-timedelta(days=1)
    opening=previous.replace(day=1).isoformat()
    first=today.replace(day=1).isoformat()
    tea=m.definition(state,user,"group",dict(name="Mutfak sarfı",prefix="MTF",module="sarf"))
    shirts=m.definition(state,user,"group",dict(name="Tişört",prefix="TSH",module="kkd"))
    shoes=m.definition(state,user,"group",dict(name="İş ayakkabısı",prefix="AYK",module="kkd"))
    production=m.definition(state,user,"department",dict(name="Üretim"))
    office=m.definition(state,user,"department",dict(name="İdari İşler"))
    people=[m.definition(state,user,"person",dict(code=code,name=name,department_id=dep["id"]))
            for code,name,dep in [("D01","Deniz Kaya",production),("D02","Ece Aydın",production),("D03","Selin Demir",office)]]
    for group,name,unit,sizes,color,scope in [(tea,"Siyah çay","kg","","","department"),
        (tea,"Filtre kahve","kg","","","department"),(shirts,"Polo yaka iş tişörtü","adet","M,L,XL","Lacivert","both"),
        (shirts,"Uzun kollu iş tişörtü","adet","M,L","Siyah","person"),
        (shoes,"Çelik burunlu iş ayakkabısı","çift","41,42,43","","person")]:
        m.definition(state,user,"product",dict(group_id=group["id"],name=name,unit=unit,scope=scope,
                                              sizes=sizes,color=color,minimum="3"))
    lines=[dict(variant_id=variant["id"],qty=str(5+2*index),unit_price=str(90+25*index))
           for index,variant in enumerate(state["variants"])]
    m.record_document(state,user,"receipt",dict(day=opening,purchase_day=opening,supplier="Demo Tedarik",
                                                reference="DEMO-001",lines=lines))
    m.record_document(state,user,"department",dict(day=opening,person_id=people[2]["id"],department_id=office["id"],
                                                   variant_id=state["variants"][0]["id"],qty="1.5"))
    for variant in state["variants"][2:5]:
        m.record_document(state,user,"person",dict(day=first,person_id=people[1]["id"],variant_id=variant["id"],qty="2"))
    m.count_operation(state,user,"count-new",dict(day=today.isoformat(),note="Excel sayım denemesi"))
