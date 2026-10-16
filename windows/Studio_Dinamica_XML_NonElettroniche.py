import csv
import json
import os
import re
import threading
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

API = "https://api-v2.fattureincloud.it"
APP_TITLE = "Studio Dinamica — XML fatture non elettroniche"
CONFIG_NAME = "Studio_Dinamica_XML_NonElettroniche.json"

MONTHS = [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"
]

class ApiError(Exception):
    pass

def clean_token(token):
    token = str(token or "").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    return token

def api_get(url, token, accept="application/json"):
    token = clean_token(token)
    req = urllib.request.Request(url, headers={
        "Authorization": "Bearer " + token,
        "Accept": accept,
        "User-Agent": "Studio-Dinamica-XML-NonElettroniche/1.0",
    })
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            return r.status, r.headers, r.read()
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")
        if e.code == 401:
            raise ApiError("401 — token non valido, scaduto o revocato")
        if e.code == 403:
            raise ApiError("403 — permessi insufficienti sul token")
        if e.code == 422:
            raise ApiError("422 — XML FatturaPA non disponibile per questo documento")
        raise ApiError(f"{e.code} — {body[:500]}")
    except urllib.error.URLError as e:
        raise ApiError("Errore di rete: " + str(e))
    except Exception as e:
        raise ApiError(str(e))

def to_float(v, default=0.0):
    try:
        if v is None or v == "":
            return default
        if isinstance(v, str):
            v = v.strip().replace(" ", "")
            if "," in v and "." not in v:
                v = v.replace(",", ".")
        return float(v)
    except Exception:
        return default

def boolish(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    return str(v or "").strip().lower() in {"1", "true", "yes", "si", "sì"}

def safe_name(s):
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", str(s or "").strip()).strip("_")
    return (s or "documento")[:100]

def fmt_money(v):
    s = f"{to_float(v):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def invoice_number(d):
    num = d.get("number")
    numeration = str(d.get("numeration") or "").strip()
    if num is None or num == "":
        return str(d.get("id") or "")
    n = str(num)
    return f"{n}{numeration}" if numeration else n

def stamp_info(d):
    sd = to_float(d.get("stamp_duty"), 0.0)
    if sd > 0:
        return sd, 1 if 1.95 <= sd <= 2.05 else 0
    total = 0.0
    count = 0
    for item in d.get("items_list") or []:
        text = " ".join([
            str(item.get("name") or ""),
            str(item.get("description") or ""),
            str(item.get("code") or ""),
        ]).upper()
        if "BOLLO" not in text and "MARCA DA BOLLO" not in text:
            continue
        qty = to_float(item.get("qty"), 1.0)
        unit = item.get("gross_price")
        if unit is None:
            unit = item.get("net_price")
        unit = to_float(unit, 0.0)
        amount = qty * unit
        if amount > 0:
            total += amount
            if 1.95 <= unit <= 2.05:
                count += max(1, int(round(qty)))
    return total, count

def normalize_for_xml_tag(key):
    tag = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(key or "campo"))
    if not tag or tag[0].isdigit():
        tag = "f_" + tag
    return tag

def append_json_xml(parent, key, value):
    tag = normalize_for_xml_tag(key)
    if isinstance(value, dict):
        node = ET.SubElement(parent, tag)
        for k, v in value.items():
            append_json_xml(node, k, v)
    elif isinstance(value, list):
        node = ET.SubElement(parent, tag)
        for item in value:
            append_json_xml(node, "elemento", item)
    else:
        node = ET.SubElement(parent, tag)
        if value is None:
            node.set("null", "true")
        elif isinstance(value, bool):
            node.text = "true" if value else "false"
        else:
            node.text = str(value)

def make_custom_xml(d):
    stamp_total, stamp_count = stamp_info(d)
    root = ET.Element("FatturaNonElettronica")
    root.set("origine", "FattureInCloud")
    root.set("formato", "XML dati fattura - non FatturaPA")
    summary = ET.SubElement(root, "Riepilogo")
    for key, val in [
        ("Id", d.get("id")), ("Numero", invoice_number(d)), ("Data", d.get("date")),
        ("Cliente", (d.get("entity") or {}).get("name")), ("TotaleNetto", d.get("amount_net")),
        ("TotaleIVA", d.get("amount_vat")), ("TotaleLordo", d.get("amount_gross")),
        ("BolloRilevato", stamp_total), ("NumeroBolliDa2Euro", stamp_count),
        ("FatturaElettronica", boolish(d.get("e_invoice"))),
    ]:
        n = ET.SubElement(summary, key)
        n.text = "" if val is None else str(val)
    full = ET.SubElement(root, "DatiCompletiFattureInCloud")
    for k, v in d.items():
        if str(k).startswith("_"):
            continue
        append_json_xml(full, k, v)
    try:
        ET.indent(root, space="  ")
    except AttributeError:
        pass
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def app_config_path():
    return Path(os.getenv("APPDATA", str(Path.home()))) / CONFIG_NAME

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x820")
        self.minsize(1050, 680)
        self.cfg_path = app_config_path()
        self.company_name = ""
        self.all_invoices = []
        self.visible_invoices = []
        self.summary_rows = {}
        self.output_base = None
        self.company = tk.StringVar()
        self.token = tk.StringVar()
        self.year = tk.StringVar(value=str(date.today().year))
        self.status = tk.StringVar(value="Pronto. Inserisci ID azienda e token di Studio Dinamica.")
        self.annual = tk.StringVar(value="")
        self.selected_month = 0
        self._build_ui()
        self._load_cfg()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        title = ttk.Frame(self, padding=(14, 12, 14, 4))
        title.grid(row=0, column=0, sticky="ew")
        ttk.Label(title, text="Studio Dinamica — Fatture non elettroniche", font=("Segoe UI", 19, "bold")).pack(anchor="w")
        ttk.Label(title, text="Carica le fatture non elettroniche dell'anno, riepiloga mese per mese e le esporta in XML.", font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))
        top = ttk.Frame(self, padding=(14, 8, 14, 10))
        top.grid(row=1, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=2)
        ttk.Label(top, text="ID azienda").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.company, width=18).grid(row=0, column=1, sticky="ew", padx=(8, 18))
        ttk.Label(top, text="Token personale").grid(row=0, column=2, sticky="w")
        self.token_entry = ttk.Entry(top, textvariable=self.token, show="•")
        self.token_entry.grid(row=0, column=3, sticky="ew", padx=(8, 8))
        self.show_token = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Mostra", variable=self.show_token, command=lambda: self.token_entry.configure(show="" if self.show_token.get() else "•")).grid(row=0, column=4, sticky="w")
        ttk.Label(top, text="Anno").grid(row=1, column=0, sticky="w", pady=(10, 0))
        years = [str(y) for y in range(date.today().year - 8, date.today().year + 2)]
        ttk.Combobox(top, textvariable=self.year, values=years, state="readonly", width=10).grid(row=1, column=1, sticky="w", padx=(8, 18), pady=(10, 0))
        ttk.Button(top, text="Verifica collegamento", command=self.verify).grid(row=1, column=2, padx=(0, 8), pady=(10, 0))
        ttk.Button(top, text="CARICA ANNO", command=self.load_year).grid(row=1, column=3, sticky="w", pady=(10, 0))
        ttk.Button(top, text="Scegli cartella XML…", command=self.choose_folder).grid(row=1, column=4, padx=(8, 0), pady=(10, 0))
        ttk.Label(top, textvariable=self.status, foreground="#344054").grid(row=2, column=0, columnspan=5, sticky="ew", pady=(10, 0))
        ttk.Label(top, textvariable=self.annual, font=("Segoe UI", 10, "bold")).grid(row=3, column=0, columnspan=5, sticky="ew", pady=(4, 0))
        panes = ttk.Panedwindow(self, orient="vertical")
        panes.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 8))
        summary_frame = ttk.Labelframe(panes, text="Riepilogo mese per mese", padding=8)
        detail_frame = ttk.Labelframe(panes, text="Fatture non elettroniche", padding=8)
        panes.add(summary_frame, weight=1)
        panes.add(detail_frame, weight=3)
        summary_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)
        scols = ("mese", "fatture", "totale", "bolli_n", "bolli_tot")
        self.summary_tree = ttk.Treeview(summary_frame, columns=scols, show="headings", height=8, selectmode="browse")
        for c, tt, w, a in [
            ("mese", "Mese", 170, "w"), ("fatture", "N. fatture", 100, "e"),
            ("totale", "Totale fatture €", 160, "e"), ("bolli_n", "Marche da 2 €", 130, "e"),
            ("bolli_tot", "Totale bolli €", 150, "e")
        ]:
            self.summary_tree.heading(c, text=tt)
            self.summary_tree.column(c, width=w, anchor=a)
        self.summary_tree.grid(row=0, column=0, sticky="nsew")
        sbar = ttk.Scrollbar(summary_frame, orient="vertical", command=self.summary_tree.yview)
        sbar.grid(row=0, column=1, sticky="ns")
        self.summary_tree.configure(yscrollcommand=sbar.set)
        self.summary_tree.bind("<<TreeviewSelect>>", self.on_month_select)
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(0, weight=1)
        dcols = ("numero", "data", "cliente", "totale", "bollo", "xml")
        self.invoice_tree = ttk.Treeview(detail_frame, columns=dcols, show="headings", selectmode="browse")
        for c, tt, w, a in [
            ("numero", "Numero", 100, "w"), ("data", "Data", 100, "w"),
            ("cliente", "Cliente", 330, "w"), ("totale", "Totale €", 110, "e"),
            ("bollo", "Bollo €", 90, "e"), ("xml", "XML", 390, "w")
        ]:
            self.invoice_tree.heading(c, text=tt)
            self.invoice_tree.column(c, width=w, anchor=a)
        self.invoice_tree.grid(row=0, column=0, sticky="nsew")
        dbar = ttk.Scrollbar(detail_frame, orient="vertical", command=self.invoice_tree.yview)
        dbar.grid(row=0, column=1, sticky="ns")
        self.invoice_tree.configure(yscrollcommand=dbar.set)
        actions = ttk.Frame(self, padding=(14, 2, 14, 12))
        actions.grid(row=3, column=0, sticky="ew")
        ttk.Button(actions, text="Esporta XML fattura selezionata", command=self.export_selected).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(actions, text="Esporta XML mese selezionato", command=self.export_month).grid(row=0, column=1, padx=6)
        ttk.Button(actions, text="ESPORTA TUTTE XML", command=self.export_all).grid(row=0, column=2, padx=6)
        ttk.Button(actions, text="Esporta riepilogo CSV", command=self.export_summary_csv).grid(row=0, column=3, padx=6)
        ttk.Button(actions, text="Apri cartella", command=self.open_folder).grid(row=0, column=4, padx=6)
        ttk.Label(actions, text="Nota: Totale fatture = amount_gross Fatture in Cloud. I bolli sono mostrati separatamente per evitare doppi conteggi.", foreground="#667085").grid(row=1, column=0, columnspan=5, sticky="w", pady=(8, 0))

    def _load_cfg(self):
        try:
            d = json.loads(self.cfg_path.read_text("utf-8"))
            self.company.set(str(d.get("company_id") or ""))
            if d.get("year"):
                self.year.set(str(d["year"]))
            if d.get("output_base"):
                self.output_base = Path(d["output_base"])
        except Exception:
            pass

    def _save_cfg(self):
        try:
            self.cfg_path.write_text(json.dumps({
                "company_id": self.company.get().strip(), "year": self.year.get(),
                "output_base": str(self.output_base) if self.output_base else "",
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def creds(self):
        c = str(self.company.get() or "").strip()
        t = clean_token(self.token.get())
        if not c or not t:
            raise ApiError("Inserisci ID azienda e token personale.")
        self._save_cfg()
        return c, t

    def run_bg(self, fn):
        threading.Thread(target=fn, daemon=True).start()

    def set_status(self, text):
        self.after(0, lambda: self.status.set(str(text)))

    def verify(self):
        def job():
            try:
                c, t = self.creds()
                _, _, body = api_get(f"{API}/c/{urllib.parse.quote(c)}/company/info", t)
                js = json.loads(body.decode("utf-8"))
                name = str((js.get("data") or {}).get("name") or "azienda")
                self.company_name = name
                extra = "" if "DINAMICA" in name.upper() else " — ATTENZIONE: il nome azienda non contiene 'Dinamica'"
                self.set_status("Collegamento riuscito: " + name + extra)
            except Exception as e:
                self.set_status("Errore: " + str(e))
        self.run_bg(job)

    def year_bounds(self):
        y = int(self.year.get())
        return f"{y:04d}-01-01", f"{y:04d}-12-31"

    def load_year(self):
        def job():
            try:
                c, t = self.creds()
                start, end = self.year_bounds()
                page, last_page, rows = 1, 1, []
                while True:
                    q = f"date >= '{start}' and date <= '{end}'"
                    params = {"q": q, "type": "invoice", "page": str(page), "per_page": "100", "fieldset": "detailed"}
                    url = f"{API}/c/{urllib.parse.quote(c)}/issued_documents?{urllib.parse.urlencode(params)}"
                    self.set_status(f"Carico fatture {self.year.get()} — pagina {page}…")
                    _, _, body = api_get(url, t)
                    js = json.loads(body.decode("utf-8"))
                    data = js.get("data") or []
                    if page == 1:
                        last_page = int(js.get("last_page") or 1)
                    rows.extend(data)
                    if page >= last_page:
                        break
                    page += 1
                non_e = []
                for d in rows:
                    if str(d.get("type") or "").lower() != "invoice" or boolish(d.get("e_invoice")):
                        continue
                    d["_stamp_total"], d["_stamp_count"] = stamp_info(d)
                    d["_xml_status"] = "Da esportare"
                    non_e.append(d)
                non_e.sort(key=lambda x: (str(x.get("date") or ""), invoice_number(x), str(x.get("id") or "")))
                self.all_invoices = non_e
                self.after(0, self.render_all)
                self.set_status(f"Caricate {len(non_e)} fatture NON elettroniche per il {self.year.get()}.")
            except Exception as e:
                self.set_status("Errore: " + str(e))
        self.run_bg(job)

    def render_all(self):
        self.build_summary(); self.render_summary(); self.selected_month = 0; self.render_invoices(); self.update_annual_line()

    def build_summary(self):
        rows = {m: {"count": 0, "total": 0.0, "stamp_count": 0, "stamp_total": 0.0} for m in range(1, 13)}
        for d in self.all_invoices:
            try: m = int(str(d.get("date") or "")[5:7])
            except Exception: continue
            if m not in rows: continue
            r = rows[m]; r["count"] += 1; r["total"] += to_float(d.get("amount_gross"), 0.0)
            r["stamp_total"] += to_float(d.get("_stamp_total"), 0.0); r["stamp_count"] += int(d.get("_stamp_count") or 0)
        self.summary_rows = rows

    def render_summary(self):
        for iid in self.summary_tree.get_children(): self.summary_tree.delete(iid)
        for m in range(1, 13):
            r = self.summary_rows[m]
            self.summary_tree.insert("", "end", iid=str(m), values=(MONTHS[m-1], r["count"], fmt_money(r["total"]), r["stamp_count"], fmt_money(r["stamp_total"])))

    def on_month_select(self, _event=None):
        sel = self.summary_tree.selection()
        if not sel: return
        try: self.selected_month = int(sel[0])
        except Exception: self.selected_month = 0
        self.render_invoices()

    def render_invoices(self):
        for iid in self.invoice_tree.get_children(): self.invoice_tree.delete(iid)
        rows = [d for d in self.all_invoices if str(d.get("date") or "")[5:7] == f"{self.selected_month:02d}"] if self.selected_month else list(self.all_invoices)
        self.visible_invoices = rows
        for i, d in enumerate(rows):
            ent = d.get("entity") or {}
            self.invoice_tree.insert("", "end", iid=str(i), values=(invoice_number(d), str(d.get("date") or ""), str(ent.get("name") or ""), fmt_money(d.get("amount_gross")), fmt_money(d.get("_stamp_total")), str(d.get("_xml_status") or "")))

    def update_annual_line(self):
        total = sum(to_float(d.get("amount_gross"), 0.0) for d in self.all_invoices)
        stamp_total = sum(to_float(d.get("_stamp_total"), 0.0) for d in self.all_invoices)
        stamp_count = sum(int(d.get("_stamp_count") or 0) for d in self.all_invoices)
        self.annual.set(f"ANNO {self.year.get()} — Fatture non elettroniche: {len(self.all_invoices)}   |   Totale fatture: € {fmt_money(total)}   |   Marche da 2 €: {stamp_count}   |   Totale bolli: € {fmt_money(stamp_total)}")

    def choose_folder(self):
        p = filedialog.askdirectory(title="Scegli cartella principale per l'esportazione")
        if p:
            self.output_base = Path(p); self._save_cfg(); self.set_status("Cartella impostata: " + str(self.output_base))

    def output_root(self):
        base = self.output_base or (Path.home() / "Downloads" / "StudioDinamica_XML")
        root = base / str(self.year.get()); root.mkdir(parents=True, exist_ok=True); return root

    def month_output_dir(self, d):
        try: m = int(str(d.get("date") or "")[5:7])
        except Exception: m = 0
        p = self.output_root() / (f"{m:02d}_{MONTHS[m-1]}" if 1 <= m <= 12 else "00_SenzaMese")
        p.mkdir(parents=True, exist_ok=True); return p

    def unique_path(self, d):
        num, dte, iid = safe_name(invoice_number(d)), safe_name(str(d.get("date") or "senza_data")), safe_name(str(d.get("id") or ""))
        base = f"{dte}_fattura_{num}_id_{iid}"; path = self.month_output_dir(d) / f"{base}.xml"
        if not path.exists(): return path
        n = 2
        while True:
            candidate = path.with_name(f"{base}_{n}.xml")
            if not candidate.exists(): return candidate
            n += 1

    def try_fic_xml(self, d, c, t):
        url = f"{API}/c/{urllib.parse.quote(str(c))}/issued_documents/{d.get('id')}/e_invoice/xml?avoid_einvoice_check=1&include_attachment=false"
        _, _, body = api_get(url, t, "text/xml,application/xml,*/*")
        if not body.lstrip().startswith(b"<"): raise ApiError("Risposta XML non valida")
        return body, "XML FIC"

    def export_one(self, d, c, t):
        path = self.unique_path(d)
        try: body, kind = self.try_fic_xml(d, c, t)
        except Exception: body, kind = make_custom_xml(d), "XML dati (non FatturaPA)"
        path.write_bytes(body); d["_xml_status"] = f"✓ {kind}: {path.name}"; return path, kind

    def sync_visible_status(self):
        for i, d in enumerate(self.visible_invoices):
            if self.invoice_tree.exists(str(i)): self.invoice_tree.set(str(i), "xml", str(d.get("_xml_status") or ""))

    def selected_invoice(self):
        sel = self.invoice_tree.selection()
        if not sel: return None
        idx = int(sel[0]); return self.visible_invoices[idx] if 0 <= idx < len(self.visible_invoices) else None

    def export_selected(self):
        d = self.selected_invoice()
        if d is None: messagebox.showinfo(APP_TITLE, "Seleziona una fattura."); return
        def job():
            try:
                c, t = self.creds(); self.set_status("Esporto fattura " + invoice_number(d) + "…")
                path, kind = self.export_one(d, c, t); self.after(0, self.sync_visible_status); self.set_status(f"Salvata ({kind}): {path}")
            except Exception as e:
                d["_xml_status"] = "✕ " + str(e); self.after(0, self.sync_visible_status); self.set_status("Errore: " + str(e))
        self.run_bg(job)

    def export_batch(self, docs, label):
        if not docs: messagebox.showinfo(APP_TITLE, "Nessuna fattura da esportare."); return
        def job():
            ok = fallback = errors = 0
            try:
                c, t = self.creds()
                for i, d in enumerate(docs, start=1):
                    self.set_status(f"{label}: {i}/{len(docs)} — fattura {invoice_number(d)}")
                    try:
                        _, kind = self.export_one(d, c, t); ok += 1
                        if kind != "XML FIC": fallback += 1
                    except Exception as e:
                        errors += 1; d["_xml_status"] = "✕ " + str(e)
                    self.after(0, self.sync_visible_status)
                self.write_summary_csv(self.output_root() / f"Riepilogo_{self.year.get()}.csv")
                self.set_status(f"Completato: {ok} XML salvati ({fallback} XML dati fallback), {errors} errori. Riepilogo CSV creato.")
            except Exception as e: self.set_status("Errore: " + str(e))
        self.run_bg(job)

    def export_month(self):
        if not self.selected_month: messagebox.showinfo(APP_TITLE, "Seleziona un mese nella tabella del riepilogo."); return
        docs = [d for d in self.all_invoices if str(d.get("date") or "")[5:7] == f"{self.selected_month:02d}"]
        self.export_batch(docs, MONTHS[self.selected_month - 1])

    def export_all(self):
        self.export_batch(list(self.all_invoices), "Tutte")

    def write_summary_csv(self, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        total_year = stamp_year = 0.0; stamp_count_year = 0
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Studio/Azienda", self.company_name or "Studio Dinamica"]); w.writerow(["Anno", self.year.get()]); w.writerow([])
            w.writerow(["Mese", "N. fatture non elettroniche", "Totale fatture", "N. marche da 2 euro", "Totale bolli"])
            for m in range(1, 13):
                r = self.summary_rows.get(m, {}); total_year += to_float(r.get("total"), 0.0); stamp_year += to_float(r.get("stamp_total"), 0.0); stamp_count_year += int(r.get("stamp_count") or 0)
                w.writerow([MONTHS[m-1], int(r.get("count") or 0), fmt_money(r.get("total")), int(r.get("stamp_count") or 0), fmt_money(r.get("stamp_total"))])
            w.writerow([]); w.writerow(["TOTALE ANNO", len(self.all_invoices), fmt_money(total_year), stamp_count_year, fmt_money(stamp_year)]); w.writerow([])
            w.writerow(["Nota", "Il totale fatture deriva da amount_gross Fatture in Cloud. I bolli sono esposti separatamente."])

    def export_summary_csv(self):
        if not self.summary_rows: messagebox.showinfo(APP_TITLE, "Carica prima le fatture dell'anno."); return
        p = filedialog.asksaveasfilename(title="Salva riepilogo", defaultextension=".csv", initialfile=f"Studio_Dinamica_Riepilogo_{self.year.get()}.csv", filetypes=[("CSV", "*.csv")])
        if not p: return
        try: self.write_summary_csv(Path(p)); self.set_status("Riepilogo CSV salvato: " + p)
        except Exception as e: messagebox.showerror(APP_TITLE, str(e))

    def open_folder(self):
        p = self.output_root()
        try: os.startfile(str(p))
        except Exception: messagebox.showinfo(APP_TITLE, "Cartella: " + str(p))

if __name__ == "__main__":
    App().mainloop()
