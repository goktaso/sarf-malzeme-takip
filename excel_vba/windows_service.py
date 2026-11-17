"""Optional Service Control Manager deployment. The local pilot uses service.py."""

import json
import os
from pathlib import Path
import sys
import threading

# The service host may load this file by its absolute path.
REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0,str(REPO))
from sqlserver import load_runtime

load_runtime()
import servicemanager
import win32service
import win32serviceutil

from excel_vba.file_store import FileStore
from excel_vba.service import create_server


class SarfExcelService(win32serviceutil.ServiceFramework):
    _svc_name_="SarfMalzemeExcel"
    _svc_display_name_="Sarf Malzeme Excel Dosya Servisi"
    _svc_description_="Excel sarf, KKD ve fiziksel sayım istemcisi için Windows kimliği ve dosya saklama servisi."

    def __init__(self,args):
        super().__init__(args)
        self.stop=threading.Event()

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.stop.set()

    def SvcDoRun(self):
        path=Path(os.environ["ProgramData"])/"SarfMalzemeExcel"/"service.json"
        settings=json.loads(path.read_text(encoding="utf-8"))
        with_scope=FileStore(settings["data"])
        try:
            if not with_scope.read()["users"]:
                raise RuntimeError("Önce dosya deposunu oluşturup Windows yöneticisini tanımlayın.")
            with create_server(settings,with_scope) as server:
                thread=threading.Thread(target=server.serve_forever,daemon=True)
                thread.start()
                servicemanager.LogInfoMsg("Sarf Malzeme Excel dosya servisi başladı.")
                self.stop.wait()
                server.shutdown()
                thread.join()
        finally:
            with_scope.close()


if __name__=="__main__":
    win32serviceutil.HandleCommandLine(SarfExcelService)
