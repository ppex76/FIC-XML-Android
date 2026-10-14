import json, os, re, threading, urllib.parse, urllib.request, urllib.error
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

API = "https://api-v2.fattureincloud.it"
APP = "FIC XML PC v4"

class ApiError(Exception):
    pass

def request(url, token, accept="application/json"):
    token = str(token or "").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    req = urllib.request.Request(url, headers={
        "Authorization": "Bearer " + token,
        "Accept": accept,
        "User-Agent": "FIC-XML-PC/4.0"
    })
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, r.headers, r.read()
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")
        if e.code == 401:
            raise ApiError("401 — token non valido o scaduto")
        if e.code == 403:
            raise ApiError("403 — permessi insufficienti")
        if e.code == 422 and "fattura elettronica" in body.lower():
            raise ApiError("422 — XML non disponibile tramite API pubblica")
        raise ApiError(f"{e.code} — {body[:300]}")
    except Exception as e:
        raise ApiError(str(e))

def safe(s):
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", str(s or "").strip()).strip("_")
    return s[:100]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Fatture in Cloud — XML PC v4")
        self.geometry("1080x700")
        self.minsize(930, 600)
        self.invoices = []
        self.cfg = Path(os.getenv("APPDATA", str(Path.home()))) / "FIC_XML_PC_v4.json"
        self.download_dir = None
        self._ui()
        self._load_cfg()

    def _ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        top = ttk.Frame(self, padding=14)
        top.grid(row=0, column=0, sticky="ew")
        for c in [1, 3]: top.columnconfigure(c, weight=1)
        ttk.Label(top, text="ID azienda").grid(row=0, column=0, sticky="w")
        self.company = tk.StringVar()
        ttk.Entry(top, textvariable=self.company, width=18).grid(row=0, column=1, sticky="ew", padx=(8, 18))
        ttk.Label(top, text="Token personale").grid(row=0, column=2, sticky="w")
        self.token = tk.StringVar()
        self.token_entry = ttk.Entry(top, textvariable=self.token, show="•")
        self.token_entry.grid(row=0, column=3, sticky="ew", padx=(8, 8))
        self.show_token = tk.BooleanVar()
        ttk.Checkbutton(top, text="Mostra", variable=self.show_token,
                        command=lambda:self.token_entry.config(show="" if self.show_token.get() else "•")).grid(row=0,column=4)
        ttk.Label(top, text="Mese").grid(row=1,column=0,sticky="w",pady=(12,0))
        months=["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno","Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"]
        self.month=tk.StringVar(value=months[date.today().month-1])
        self.monthbox=ttk.Combobox(top,textvariable=self.month,values=months,state="readonly",width=14)
        self.monthbox.grid(row=1,column=1,sticky="w",padx=(8,18),pady=(12,0))
        ttk.Label(top,text="Anno").grid(row=1,column=2,sticky="w",pady=(12,0))
        years=[str(y) for y in range(date.today().year-5,date.today().year+2)]
        self.year=tk.StringVar(value=str(date.today().year))
        ttk.Combobox(top,textvariable=self.year,values=years,state="readonly",width=9).grid(row=1,column=3,sticky="w",padx=(8,18),pady=(12,0))
        ttk.Label(top,text="Filtro").grid(row=2,column=0,sticky="w",pady=(12,0))
        self.doc_filter=tk.StringVar(value="Tutti i documenti emessi")
        ttk.Combobox(top,textvariable=self.doc_filter,values=["Tutti i documenti emessi","Solo fatture"],state="readonly",width=24).grid(row=2,column=1,sticky="w",padx=(8,18),pady=(12,0))
        ttk.Button(top,text="Verifica collegamento",command=self.verify).grid(row=2,column=3,pady=(12,0),padx=(0,8),sticky="e")
        ttk.Button(top,text="Carica documenti",command=self.load).grid(row=2,column=4,pady=(12,0),sticky="e")

        mid=ttk.Frame(self,padding=(14,0,14,8)); mid.grid(row=1,column=0,sticky="nsew"); mid.columnconfigure(0,weight=1); mid.rowconfigure(0,weight=1)
        cols=("numero","data","cliente","tipo","stato")
        self.tree=ttk.Treeview(mid,columns=cols,show="headings",selectmode="browse")
        for c,t,w in [("numero","Numero",110),("data","Data",110),("cliente","Cliente",320),("tipo","Tipo",160),("stato","Stato",320)]:
            self.tree.heading(c,text=t); self.tree.column(c,width=w,anchor="w")
        self.tree.grid(row=0,column=0,sticky="nsew")
        sb=ttk.Scrollbar(mid,orient="vertical",command=self.tree.yview); sb.grid(row=0,column=1,sticky="ns"); self.tree.configure(yscrollcommand=sb.set)

        actions=ttk.Frame(self,padding=14); actions.grid(row=2,column=0,sticky="ew"); actions.columnconfigure(4,weight=1)
        ttk.Button(actions,text="Esporta XML selezionato",command=self.export_selected).grid(row=0,column=0,padx=(0,8))
        ttk.Button(actions,text="Esporta tutte — una alla volta",command=self.export_all).grid(row=0,column=1,padx=(0,8))
        ttk.Button(actions,text="Apri cartella XML",command=self.open_folder).grid(row=0,column=2,padx=(0,8))
        ttk.Button(actions,text="Scegli cartella…",command=self.choose_folder).grid(row=0,column=3)
        self.status=tk.StringVar(value="Pronto."); ttk.Label(actions,textvariable=self.status,anchor="e").grid(row=0,column=4,sticky="e")

    def _load_cfg(self):
        try:
            d=json.loads(self.cfg.read_text("utf-8")); self.company.set(d.get("company_id","")); self.doc_filter.set(d.get("doc_filter","Tutti i documenti emessi"))
            if d.get("download_dir"): self.download_dir=Path(d["download_dir"])
        except: pass

    def _save_cfg(self):
        try:
            self.cfg.write_text(json.dumps({"company_id":self.company.get().strip(),"doc_filter":self.doc_filter.get(),"download_dir":str(self.download_dir) if self.download_dir else ""}),encoding="utf-8")
        except: pass

    def creds(self):
        c=str(self.company.get() or "").strip(); t=str(self.token.get() or "").strip()
        if not c or not t: raise ApiError("Inserisci ID azienda e token personale.")
        self._save_cfg(); return c,t

    def month_bounds(self):
        m=self.monthbox.current()+1; y=int(self.year.get()); import calendar
        last=calendar.monthrange(y,m)[1]
        return f"{y:04d}-{m:02d}-01",f"{y:04d}-{m:02d}-{last:02d}",f"{y:04d}-{m:02d}"

    def run_bg(self,fn): threading.Thread(target=fn,daemon=True).start()
    def set_status(self,s): self.after(0,lambda:self.status.set(str(s)))

    def verify(self):
        def job():
            try:
                c,t=self.creds(); _,_,b=request(f"{API}/c/{urllib.parse.quote(c)}/company/info",t); data=json.loads(b.decode()); name=str((data.get("data") or {}).get("name","azienda")); self.set_status("Collegamento riuscito: "+name)
            except Exception as e: self.set_status("Errore: "+str(e))
        self.run_bg(job)

    def load(self):
        def job():
            try:
                c,t=self.creds(); start,end,key=self.month_bounds(); rows=[]; page=1; last_page=None
                while True:
                    q=f"date >= '{start}' and date <= '{end}'"; params={"q":q,"page":str(page),"per_page":"100","fieldset":"basic"}
                    if self.doc_filter.get()=="Solo fatture": params["type"]="invoice"
                    url=f"{API}/c/{urllib.parse.quote(c)}/issued_documents?{urllib.parse.urlencode(params)}"; self.set_status(f"Carico pagina {page}…")
                    _,_,b=request(url,t); js=json.loads(b.decode()); data=js.get("data") or []
                    if page==1:
                        last_page=int(js.get("last_page") or 1); total=int(js.get("total") or 0); self.set_status(f"Trovati {total} documenti, carico {last_page} pagina/e…")
                    for d in data:
                        rows.append({"id":d.get("id"),"number":str(d.get("number") if d.get("number") is not None else d.get("id") or ""),"date":str(d.get("date") or ""),"client":str((d.get("entity") or {}).get("name","") or ""),"type":str(d.get("type") or ""),"status":"Non esportata"})
                    if page >= (last_page or 1): break
                    page += 1
                rows.sort(key=lambda x:(x.get("date",""),x.get("number",""))); self.invoices=rows; self.after(0,self.render); self.set_status(f"Trovati {len(rows)} documenti per {key}.")
            except Exception as e: self.set_status("Errore: "+str(e))
        self.run_bg(job)

    def render(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for n,inv in enumerate(self.invoices): self.tree.insert("","end",iid=str(n),values=(str(inv["number"]),str(inv["date"]),str(inv["client"]),str(inv["type"]),str(inv["status"])))

    def output_dir(self):
        _,_,key=self.month_bounds(); base=self.download_dir or (Path.home()/"Downloads"/"FattureInCloud"); p=base/key; p.mkdir(parents=True,exist_ok=True); return p

    def choose_folder(self):
        p=filedialog.askdirectory(title="Scegli cartella principale per gli XML")
        if p: self.download_dir=Path(p); self._save_cfg(); self.set_status("Cartella impostata: "+str(p))

    def open_folder(self):
        p=self.output_dir()
        try: os.startfile(str(p))
        except Exception: messagebox.showinfo(APP,f"Cartella: {p}")

    def export_one(self,idx,c,t):
        inv=self.invoices[idx]
        url=f"{API}/c/{urllib.parse.quote(str(c))}/issued_documents/{inv['id']}/e_invoice/xml?avoid_einvoice_check=1&include_attachment=false"
        _,h,b=request(url,t,"text/xml,application/xml,*/*")
        original_name=""; cd=h.get("Content-Disposition",""); m=re.search(r'filename="?([^";]+)',cd,re.I)
        if m:
            original_name=safe(m.group(1))
            if original_name.lower().endswith(".xml"): original_name=original_name[:-4]
        unique_base=safe(f"{inv['date']}_fattura_{inv['number']}_id_{inv['id']}")
        fn=f"{unique_base}_{original_name}.xml" if original_name else f"{unique_base}.xml"
        path=self.output_dir()/fn
        if path.exists():
            n=2; stem=path.stem
            while True:
                candidate=path.with_name(f"{stem}_{n}.xml")
                if not candidate.exists(): path=candidate; break
                n += 1
        path.write_bytes(b); inv["status"]="✓ "+str(path.name); self.after(0,lambda:self.tree.set(str(idx),"stato",str(inv["status"]))); return path

    def export_selected(self):
        sel=self.tree.selection()
        if not sel: return messagebox.showinfo(APP,"Seleziona un documento.")
        idx=int(sel[0])
        def job():
            try:
                c,t=self.creds(); self.set_status("Esporto documento "+str(self.invoices[idx]["number"])+"…"); p=self.export_one(idx,c,t); self.set_status("Salvata: "+str(p))
            except Exception as e:
                self.invoices[idx]["status"]="✕ "+str(e); self.after(0,lambda:self.tree.set(str(idx),"stato",str(self.invoices[idx]["status"]))); self.set_status("Errore: "+str(e))
        self.run_bg(job)

    def export_all(self):
        if not self.invoices: return messagebox.showinfo(APP,"Carica prima i documenti del mese.")
        def job():
            ok=ko=0
            try:
                c,t=self.creds()
                for i,inv in enumerate(self.invoices):
                    self.set_status(f"{i+1}/{len(self.invoices)} — documento {inv['number']}")
                    try: self.export_one(i,c,t); ok += 1
                    except Exception as e:
                        ko += 1; inv["status"]="✕ "+str(e); self.after(0,lambda i=i,inv=inv:self.tree.set(str(i),"stato",str(inv["status"])))
                self.set_status(f"Completato: {ok} XML salvati, {ko} non esportati.")
            except Exception as e: self.set_status("Errore: "+str(e))
        self.run_bg(job)

if __name__=="__main__":
    App().mainloop()
