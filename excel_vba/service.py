"""Windows-authenticated file-backed API used by the Excel workbook."""

import argparse
import base64
import secrets
import ssl
import sys
import time
from decimal import Decimal
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit
from xml.etree.ElementTree import Element, ParseError, SubElement, fromstring, tostring

import access
import excel_import
import export_files
import sqlserver
import stock_counts
from excel_vba import inventory as m, views
from excel_vba.file_store import FileStore, insert
from server import IMPORT_TEMPLATE, XLSX_TYPE


def plain(value):
    return "".join(c for c in str(value) if c == "\n" or c == "\t" or ord(c) >= 32)


def xml_grid(name, columns, rows):
    document=Element("grid", title=views.VIEWS[name][0], name=name)
    roles={"Birim":"unit","Eldeki miktar":"qty","Miktar":"qty","Tüketim miktarı":"qty",
           "Stok değeri":"cost","Stok değeri farkı":"cost","Net tüketim maliyeti":"cost",
           "Ay":"month","_Rapor":"report","_Varyant":"count-variant"}
    for title in columns:
        SubElement(document,"column",role=roles.get(title,"")).text=title
    for row in rows:
        element=SubElement(document,"row")
        for value in row:
            cell=SubElement(element,"cell",type="number" if isinstance(value,(int,Decimal)) else "text")
            if value is not None: cell.text=plain(value)
    return document


def count_file(state,user,key):
    m.allow(user,"view.sayimlar")
    m.global_only(user)
    count=m.get(state,"counts",key)
    lines=[]
    for row in count["lines"]:
        v=m.get(state,"variants",row["variant_id"])
        p=m.on_day(state,"products",v["product_id"],count["day"])
        lines.append(row|{key:v[key] for key in ("color","size")}|{key:p[key] for key in ("code","name","unit")})
    return count,lines


def upload_count(state,user,count_key,payload):
    count_id,__,revision=count_key.partition(":")
    count,lines=count_file(state,user,count_id)
    if str(count["revision"])!=revision:
        raise ValueError("Sayım değişti. Bilgileri yenileyin.")
    ids={(m.normalized(r["code"]),m.normalized(r["color"]),m.normalized(r["size"])):r["variant_id"] for r in lines}
    records=stock_counts.spreadsheet_rows(payload)
    used=set()
    for number,(code,_,color,size,_,_,qty,price) in records:
        try:
            variant=ids.get(tuple(map(m.normalized,(code,color,size))))
            if not variant or variant in used or not qty:
                raise ValueError("Varyant eşleşmedi, tekrar ediyor veya sayılan miktar eksik.")
            used.add(variant)
            data=dict(id=count["id"],revision=count["revision"],variant_id=variant,qty=qty)
            if price: data["unit_price"]=price
            m.count_operation(state,user,"count-line",data)
        except (access.Forbidden,ValueError) as error:
            raise ValueError(f"Sayım, satır {number}: {error}") from error
    return len(records)


class Server(ThreadingHTTPServer):
    daemon_threads=True

    def __init__(self,address,store,secure):
        self.store=store
        self.sessions={}
        self.login_attempts={}
        self.local_identity=None
        self.local_proof=None
        self.local_envelope=None
        self.secure=secure
        super().__init__(address,Handler)


class Handler(BaseHTTPRequestHandler):
    protocol_version="HTTP/1.1"

    def log_message(self, format, *args):
        if sys.stderr:
            super().log_message(format,*args)

    def send(self,status,content,kind="application/xml; charset=utf-8",headers=()):
        self.send_response(status)
        for key,value in headers: self.send_header(key,value)
        self.send_header("Content-Type",kind)
        self.send_header("Content-Length",str(len(content)))
        self.send_header("X-Content-Type-Options","nosniff")
        self.send_header("Cache-Control","no-store")
        self.end_headers()
        self.wfile.write(content)

    def xml(self,element,status=200,headers=()):
        self.send(status,tostring(element,encoding="UTF-8",xml_declaration=True),headers=headers)

    def message(self,message,status=200,headers=()):
        node=Element("result")
        SubElement(node,"message").text=plain(message)
        self.xml(node,status,headers)

    def session(self,state):
        token=self.headers.get("X-Sarf-Session","")
        with self.server.store.guard:
            login=self.server.sessions.get(token)
        if not login or login[1]<=time.monotonic():
            raise access.Forbidden("Windows oturumunu yeniden açın.")
        user=next((row for row in state["users"] if row["sid"]==login[0] and row["active"]),None)
        if user is None:
            raise access.Forbidden("Bu Windows hesabı uygulamada etkin değil.")
        return user

    def windows_login(self, raw):
        import pywintypes
        import win32con
        import win32security

        if len(raw)>8192 or self.headers.get_content_type()!="application/x-www-form-urlencoded":
            raise ValueError("Geçersiz giriş isteği.")
        fields=parse_qs(raw.decode(),max_num_fields=2)
        account=fields.get("account",[""])[0]
        password=fields.get("password",[""])[0]
        failure="Windows kullanıcı adı veya parolası doğrulanamadı."
        state=self.server.store.read()
        user=next((row for row in state["users"] if row["active"] and row["account"].casefold()==account.casefold()),None)
        if user is None or not password or "\x00" in password:
            self.message(failure,401)
            return
        domain,separator,username=user["account"].rpartition("\\")
        if not separator:
            self.message(failure,401)
            return
        now=time.monotonic()
        identity=user["sid"]
        with self.server.store.guard:
            recent=[stamp for stamp in self.server.login_attempts.get(identity,[]) if now-stamp<60]
            if len(recent)>=3:
                self.message("Bir dakika sonra tekrar deneyin.",429)
                return
            self.server.login_attempts[identity]=recent+[now]
        try:
            token=win32security.LogonUser(username,domain,password,win32con.LOGON32_LOGON_NETWORK,
                                         win32con.LOGON32_PROVIDER_DEFAULT)
            try:
                sid,_=win32security.GetTokenInformation(token,win32security.TokenUser)
                verified=win32security.ConvertSidToStringSid(sid)
            finally:
                token.Close()
        except pywintypes.error:
            self.message(failure,401)
            return
        if verified!=identity:
            self.message(failure,401)
            return
        self.issue_session(identity, now, failure)

    def issue_session(self, identity, now, failure):
        with self.server.store.update() as state:
            user=next((row for row in state["users"] if row["sid"]==identity and row["active"]),None)
            if user is None:
                raise access.Forbidden(failure)
            m.audit(state,user,"Windows oturumu")
        session=secrets.token_urlsafe(32)
        with self.server.store.guard:
            self.server.login_attempts.pop(identity,None)
            self.server.sessions={key:value for key,value in self.server.sessions.items() if value[1]>now}
            self.server.sessions[session]=(identity,now+8*60*60)
        self.xml(Element("login",account=user["account"],name=user["display_name"],session=session))

    def menu(self,state,user):
        root=Element("menu", account=user["account"], name=user["display_name"],
                     export="1" if m.is_allowed(user,"export") else "0",
                     admin="1" if user["is_admin"] else "0", today=m.today())
        for name,(label,permission) in views.VIEWS.items():
            if views.visible(user,permission):
                SubElement(root,"view",key=name,label=label)
        for name,(label,permission) in views.ACTIONS.items():
            if views.visible(user,permission):
                SubElement(root,"action",key=name,label=label)
        if views.visible(user,"view.sayimlar"):
            for row in reversed(state["counts"]):
                writable=row["status"]!="approved" and (row["status"]=="draft" or user["is_admin"])
                if writable:
                    try:
                        m.period(state,user,row["day"])
                    except access.Forbidden:
                        writable=False
                SubElement(root,"count",key=f"{row['id']}:{row['revision']}",
                           label=f"{views.count_name(row)} · {row['day']} · {m.COUNT_STATUS[row['status']]}",
                           editable="1" if writable and views.visible(user,"count.edit") else "0",
                           postable="1" if writable and row["status"] in ("draft","locked")
                           and views.visible(user,"count.approve")
                           and any(line["counted_qty"] is not None for line in row["lines"]) else "0")
        return root

    def do_GET(self):
        url=urlsplit(self.path)
        args={key:values[0] for key,values in parse_qs(url.query).items()}
        try:
            if url.path=="/health":
                root=Element("service",storage="files")
                if self.server.local_proof is not None:
                    root.set("local_login","1")
                self.xml(root)
                return
            if url.path=="/local-proof" and self.client_address[0]=="127.0.0.1" and self.server.local_envelope:
                self.send(200,self.server.local_envelope,"text/plain; charset=ascii")
                return
            state=self.server.store.read()
            user=self.session(state)
            if url.path=="/menu":
                self.xml(self.menu(state,user))
            elif url.path=="/list" and args.get("name") in views.VIEWS:
                name=args["name"]
                result=xml_grid(name,*views.grid(state,user,name,args))
                if name=="count-lines" and views.visible(user,"count.edit"):
                    count=m.get(state,"counts",args.get("id"))
                    if count["status"]!="approved" and (count["status"]=="draft" or user["is_admin"]):
                        try:
                            m.period(state,user,count["day"])
                        except access.Forbidden:
                            pass
                        else:
                            result.set("count_key",f"{count['id']}:{count['revision']}")
                self.xml(result)
            elif url.path=="/form" and args.get("name") in views.ACTIONS:
                title,fields=views.form(state,user,args["name"],args.get("record"))
                root=Element("form",title=title,key=args["name"],nonce=secrets.token_hex(16))
                for field in fields:
                    element=SubElement(root,"field",key=field["key"],label=field["label"],kind=field["kind"],
                                       value=plain(field["value"]),lookup="1" if field.get("lookup") else "0")
                    for value,label in field.get("options",[]):
                        SubElement(element,"option",key=str(value)).text=plain(label)
                self.xml(root)
            elif url.path=="/template":
                if args.get("count"):
                    count,lines=count_file(state,user,args["count"].split(":",1)[0])
                    body=stock_counts.workbook(count|{"lines":lines},m.is_allowed(user,"field.cost"))
                else:
                    m.admin_only(user)
                    body=IMPORT_TEMPLATE.read_bytes()
                self.send(200,body,XLSX_TYPE)
            else:
                self.message("Sayfa bulunamadı.",404)
        except access.Forbidden as error:
            self.message(str(error),403)
        except (ValueError,OverflowError) as error:
            self.message(str(error),400)
        except (OSError, ModuleNotFoundError):
            self.message("Dosya servisi işlemi tamamlayamadı.",503)

    def do_POST(self):
        url=urlsplit(self.path)
        try:
            length=int(self.headers.get("Content-Length","0"))
            if length < 0 or length > excel_import.MAX_BYTES:
                raise ValueError("En fazla 5 MB veri gönderin.")
            self.connection.settimeout(20)
            raw=self.rfile.read(length)
            if url.path=="/local-login":
                proof=self.server.local_proof
                if self.client_address[0]!="127.0.0.1" or proof is None or not secrets.compare_digest(raw,proof):
                    raise access.Forbidden("Yerel Windows oturumu doğrulanamadı.")
                self.issue_session(self.server.local_identity,time.monotonic(),
                                   "Bu Windows hesabı uygulamada etkin değil.")
                return
            if url.path=="/login":
                self.windows_login(raw)
                return
            if url.path=="/logout":
                with self.server.store.guard:
                    self.server.sessions.pop(self.headers.get("X-Sarf-Session",""),None)
                self.message("Oturum kapatıldı.")
                return
            if url.path=="/summary":
                state=self.server.store.read()
                user=self.session(state)
                args={key:values[0] for key,values in parse_qs(raw.decode(),max_num_fields=12).items()}
                result=views.report_balances(state,user,args)
                root=Element("summary",day=result.get("day",""),months=str(result.get("months",0)))
                for key in ("received","balance"):
                    for unit,value in sorted(result.get(key,[])):
                        SubElement(root,key,unit=unit).text=str(Decimal(value)/1000)
                if "value" in result: SubElement(root,"value").text=str(Decimal(result["value"])/100)
                self.xml(root)
                return
            if url.path=="/pdf" and self.headers.get_content_type()=="application/xml":
                state=self.server.store.read()
                user=self.session(state)
                m.allow(user,"export")
                if b"<!DOCTYPE" in raw.upper():
                    raise ValueError("Dış XML türleri kullanılmaz.")
                document=fromstring(raw)
                name=document.get("name","")
                if name not in views.VIEWS or not views.visible(user,views.VIEWS[name][1]):
                    raise access.Forbidden("Bu ekranın çıktısına erişim yetkiniz yok.")
                sections=[{"title": item.get("title","Liste"),"columns":[col.text or "" for col in item.findall("column")],
                    "rows":[[{"text": cell.text or ""} for cell in row.findall("cell")] for row in item.findall("row")]}
                    for item in document.findall("section")]
                arguments=export_files.validate_export({"title":views.VIEWS[name][0],
                    "context":document.get("context","")[:1500],"sections":sections})
                self.send(200,export_files.write_pdf(*arguments),"application/pdf")
                return
            created_count=None
            with self.server.store.update() as state:
                user=self.session(state)
                if url.path=="/apply" and self.headers.get_content_type()=="application/x-www-form-urlencoded":
                    args={key:values[0] for key,values in parse_qs(raw.decode(),keep_blank_values=True,max_num_fields=24).items()}
                    name=args.pop("action","")
                    if name not in views.ACTIONS: raise ValueError("İşlem bulunamadı.")
                    result=views.apply(state,user,name,args)
                    if name=="count-new":
                        created_count=result["id"]
                    message=f"Kayıt tamamlandı · {result.get('code') or '#'+str(result['id'])}"
                elif url.path=="/upload" and self.headers.get_content_type()=="application/octet-stream":
                    count=parse_qs(url.query).get("count",[""])[0]
                    total=upload_count(state,user,count,raw) if count else m.import_workbook(state,user,raw)
                    message=f"{total} satır kaydedildi."
                else:
                    raise ValueError("İşlem bulunamadı.")
            if created_count is None:
                self.message(message,201)
            else:
                response=Element("result", count_id=str(created_count))
                SubElement(response,"message").text=plain(message)
                self.xml(response,201)
        except access.Forbidden as error:
            self.message(str(error),403)
        except (ValueError,OverflowError,ParseError) as error:
            self.message(str(error),400)
        except TimeoutError:
            self.message("Dosya aktarımı zaman aşımına uğradı.",408)
        except (OSError,ModuleNotFoundError):
            self.message("Dosya servisi işlemi tamamlayamadı.",503)


def bootstrap(store,account,name,seed=False):
    sid,windows_name=access.windows_account(account)
    with store.update() as state:
        if state["users"]:
            if not any(u["sid"]==sid and u["active"] and u["is_admin"] for u in state["users"]):
                raise ValueError("Kayıt dosyasının mevcut yöneticisini kullanın.")
            return
        user=insert(state,"users",sid=sid,account=windows_name,display_name=name,is_admin=1,active=1,
                    department_id=None,permissions=[])
        if seed:
            from excel_vba.samples import seed_sample

            seed_sample(state,user)


def create_server(config,store):
    host=config.get("host","127.0.0.1")
    certificate=config.get("certificate")
    key=config.get("private_key")
    if bool(certificate)!=bool(key):
        raise ValueError("HTTPS sertifikası ve anahtarı birlikte belirtilmeli.")
    if host not in ("127.0.0.1","localhost") and not certificate:
        raise ValueError("Ağ erişimi için HTTPS sertifikası ve anahtarı gerekli.")
    http=Server((host,int(config.get("port",8768))),store,bool(certificate))
    try:
        if config.get("local_login"):
            if host!="127.0.0.1" or certificate:
                raise ValueError("Parolasız oturum yalnızca yerel Excel servisine açıktır.")
            import win32api
            import win32con
            import win32crypt
            import win32security

            token=win32security.OpenProcessToken(win32api.GetCurrentProcess(),win32con.TOKEN_QUERY)
            try:
                sid,_=win32security.GetTokenInformation(token,win32security.TokenUser)
                http.local_identity=win32security.ConvertSidToStringSid(sid)
            finally:
                token.Close()
            http.local_proof=secrets.token_hex(32).encode("ascii")
            protected=win32crypt.CryptProtectData(http.local_proof,None,None,None,None,1)
            http.local_envelope=base64.b64encode(protected)
        if certificate:
            context=ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            context.minimum_version=ssl.TLSVersion.TLSv1_2
            context.load_cert_chain(certificate,key)
            http.socket=context.wrap_socket(http.socket,server_side=True)
        return http
    except BaseException:
        http.server_close()
        raise


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--data",required=True,type=Path)
    parser.add_argument("--port",type=int,default=8768)
    parser.add_argument("--host",default="127.0.0.1")
    parser.add_argument("--admin")
    parser.add_argument("--name",default="Yönetici")
    parser.add_argument("--sample",action="store_true")
    parser.add_argument("--certificate")
    parser.add_argument("--private-key")
    parser.add_argument("--local-login",action="store_true")
    args=parser.parse_args()
    if args.host not in ("127.0.0.1","localhost") and not (args.certificate and args.private_key):
        parser.error("Ağ erişimi için HTTPS sertifikası ve anahtarı gerekli.")
    sqlserver.load_runtime()
    store=FileStore(args.data)
    try:
        if args.admin: bootstrap(store,args.admin,args.name,args.sample)
        if not store.read()["users"]: parser.error("İlk Windows yöneticisini --admin ile tanımlayın.")
        with create_server(vars(args),store) as http:
            print(f"Excel dosya servisi · {args.host}:{args.port}",flush=True)
            http.serve_forever()
    finally:
        store.close()


if __name__=="__main__":
    main()
