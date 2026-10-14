import os
import re
import csv
import sys
import json
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

try:
    import fitz  # PyMuPDF
except Exception as exc:
    fitz = None
    FITZ_IMPORT_ERROR = exc
else:
    FITZ_IMPORT_ERROR = None

APP_TITLE = "NMI - Beauty Fly DDT + Listino"
DEFAULT_CUSTOMER = "BEAUTY FLY"
NMI_MARKER = "NMI DI PAOLO PEZZOTTA"
PRICE_FILE = "NMI_BeautyFly_listino.json"


def clean_line(s):
    return " ".join(str(s).replace("\xa0", " ").split()).strip()


def norm_name(s):
    return re.sub(r"[^A-Z0-9]+", " ", clean_line(s).upper()).strip()


def fmt_qty(v):
    try:
        f = float(v)
    except Exception:
        return str(v)
    if abs(f - round(f)) < 1e-9:
        return str(int(round(f)))
    return ("%.3f" % f).rstrip("0").rstrip(".").replace(".", ",")


def fmt_money(v):
    if v is None:
        return ""
    try:
        f = float(v)
    except Exception:
        return str(v)
    s = f"{f:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def parse_price(s):
    s = clean_line(s).replace("€", "").replace(" ", "")
    if not s:
        raise ValueError("Prezzo vuoto")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    return float(s)


def app_folder():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def price_path():
    return app_folder() / PRICE_FILE


def load_prices():
    p = price_path()
    if not p.exists():
        return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, list):
            out = []
            for x in data:
                if not isinstance(x, dict):
                    continue
                name = clean_line(x.get("name", ""))
                if not name:
                    continue
                code = clean_line(x.get("code", ""))
                try:
                    price = float(x.get("price", 0))
                except Exception:
                    price = 0.0
                out.append({"code": code, "name": name, "price": price})
            return out
    except Exception:
        pass
    return []


def save_prices(prices):
    price_path().write_text(json.dumps(prices, ensure_ascii=False, indent=2), encoding="utf-8")


def get_price_for(prices, code, name):
    code = clean_line(code)
    nn = norm_name(name)
    if code:
        for p in prices:
            if clean_line(p.get("code", "")) == code:
                return float(p.get("price", 0)), p
    for p in prices:
        if norm_name(p.get("name", "")) == nn:
            return float(p.get("price", 0)), p
    return None, None


def extract_document_info(text, customer_key):
    lines = [clean_line(x) for x in text.splitlines() if clean_line(x)]
    upper_lines = [x.upper() for x in lines]
    customer_idx = None
    key = customer_key.upper().strip()
    for i, line in enumerate(upper_lines):
        if key and key in line:
            customer_idx = i
            break

    upto = lines[:customer_idx] if customer_idx is not None else lines
    document_no = ""
    document_date = ""
    for i in range(len(upto) - 1):
        if re.fullmatch(r"\d{3,6}", upto[i]) and re.fullmatch(r"\d{2}/\d{2}/\d{4}", upto[i + 1]):
            document_no = upto[i]
            document_date = upto[i + 1]
    return document_no, document_date


def extract_order_refs(text):
    refs = []
    lines = [clean_line(x) for x in text.splitlines() if clean_line(x)]
    for i, line in enumerate(lines):
        u = line.upper()
        if ("RIF.VS.ORD" in u or "RIF. VS. ORD" in u or "RIF.VS.MAIL" in u or
                "RIF.VS. MAIL" in u or "VS. ORD." in u or "VS.ORD." in u or
                u.startswith("RIF.VS.MAIL") or u.startswith("RIF.VS. MAIL")):
            value = line
            if i + 1 < len(lines):
                nxt = lines[i + 1]
                if re.fullmatch(r"[A-Z0-9]{8,20}", nxt.upper()):
                    value += " " + nxt
            if value not in refs:
                refs.append(value)
    return refs


def parse_products_from_page(text):
    lines = [clean_line(x) for x in text.splitlines() if clean_line(x)]

    start = 0
    for i, line in enumerate(lines):
        if "BONIFICO BANCARIO" in line.upper():
            start = i + 1

    end = len(lines)
    for i in range(start, len(lines)):
        u = lines[i].upper()
        if u == "TOTALE" or u.startswith("NUMERO COLLI") or u == "NATURETHICS SRL":
            end = i
            break

    body = lines[start:end]
    items = []
    i = 0
    while i <= len(body) - 4:
        code = body[i]
        um = body[i + 1]
        qty_raw = body[i + 2]
        desc = body[i + 3]

        is_code = bool(re.fullmatch(r"\d{5,8}", code))
        is_um = bool(re.fullmatch(r"[A-Z]{1,5}", um.upper()))
        is_qty = bool(re.fullmatch(r"\d+(?:[\.,]\d+)?", qty_raw))
        bad_desc = desc.upper().startswith(("LOTTO:", "SCADENZA:", "TOTALE", "EUR"))

        if is_code and is_um and is_qty and not bad_desc:
            qty = float(qty_raw.replace(",", "."))
            items.append({"code": code, "description": desc, "qty": qty, "um": um.upper()})
            i += 4
        else:
            i += 1
    return items


def analyze_pdf(path, customer_key=DEFAULT_CUSTOMER):
    if fitz is None:
        raise RuntimeError("PyMuPDF non disponibile: %s" % FITZ_IMPORT_ERROR)

    key = clean_line(customer_key).upper()
    if not key:
        raise ValueError("Inserisci il cliente da cercare.")

    doc = fitz.open(path)
    totals = {}
    selected_pages = []
    docs = {}
    order_refs = []
    row_count = 0

    try:
        for page_index in range(len(doc)):
            text = doc[page_index].get_text("text") or ""
            up = text.upper()

            if key not in up or "VENDITE" not in up or NMI_MARKER not in up:
                continue

            page_no = page_index + 1
            selected_pages.append(page_no)
            document_no, document_date = extract_document_info(text, key)
            if document_no:
                docs.setdefault(document_no, document_date)

            for ref in extract_order_refs(text):
                if ref not in order_refs:
                    order_refs.append(ref)

            items = parse_products_from_page(text)
            row_count += len(items)
            for item in items:
                code = item["code"]
                if code not in totals:
                    totals[code] = {"code": code, "description": item["description"], "qty": 0.0, "um": item["um"]}
                totals[code]["qty"] += item["qty"]
                if len(item["description"]) > len(totals[code]["description"]):
                    totals[code]["description"] = item["description"]

        results = sorted(totals.values(), key=lambda x: (x["description"].upper(), x["code"]))
        total_qty = sum(x["qty"] for x in results)
        return {
            "products": results,
            "pages": selected_pages,
            "documents": docs,
            "order_refs": order_refs,
            "row_count": row_count,
            "total_qty": total_qty,
            "total_pdf_pages": len(doc),
            "customer": customer_key,
        }
    finally:
        doc.close()


class PriceDialog(tk.Toplevel):
    def __init__(self, parent, item=None):
        super().__init__(parent)
        self.title("Voce listino prezzi")
        self.resizable(False, False)
        self.result = None
        self.transient(parent)
        self.grab_set()

        item = item or {}
        self.code = tk.StringVar(value=item.get("code", ""))
        self.name = tk.StringVar(value=item.get("name", ""))
        self.price = tk.StringVar(value=fmt_money(item.get("price")) if item.get("price") is not None else "")

        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Codice (opzionale):").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.code, width=34).grid(row=0, column=1, pady=4)
        ttk.Label(frm, text="Nome prodotto:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.name, width=34).grid(row=1, column=1, pady=4)
        ttk.Label(frm, text="Prezzo unitario €:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.price, width=34).grid(row=2, column=1, pady=4)

        btn = ttk.Frame(frm)
        btn.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btn, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btn, text="Salva", command=self.on_save).pack(side="right", padx=(0, 8))
        self.bind("<Return>", lambda e: self.on_save())
        self.bind("<Escape>", lambda e: self.destroy())
        self.after(100, self.focus_force)

    def on_save(self):
        name = clean_line(self.name.get())
        if not name:
            messagebox.showwarning(APP_TITLE, "Inserisci il nome del prodotto.", parent=self)
            return
        try:
            price = parse_price(self.price.get())
        except Exception:
            messagebox.showwarning(APP_TITLE, "Inserisci un prezzo valido.", parent=self)
            return
        self.result = {"code": clean_line(self.code.get()), "name": name, "price": price}
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x760")
        self.minsize(900, 620)
        self.result = None
        self.prices = load_prices()
        self.pdf_path = tk.StringVar()
        self.customer = tk.StringVar(value=DEFAULT_CUSTOMER)
        self.status = tk.StringVar(value="Seleziona il PDF dei DDT di vendita.")
        self.summary = tk.StringVar(value="")
        self.order_line = tk.StringVar(value="DDT / ordini: -")
        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="NMI - Quantità + Valore venduto per cliente", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(outer, text="Analizza i DDT di VENDITA, somma i pezzi per prodotto e applica il tuo listino prezzi.", font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 12))

        frm = ttk.Frame(outer)
        frm.pack(fill="x")
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text="Cliente:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.customer).grid(row=0, column=1, columnspan=3, sticky="ew", pady=4)
        ttk.Label(frm, text="PDF DDT:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.pdf_path, state="readonly").grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(frm, text="Scegli PDF", command=self.choose_pdf).grid(row=1, column=2, padx=(8, 0), pady=4)
        ttk.Button(frm, text="Listino prezzi", command=self.open_price_list).grid(row=1, column=3, padx=(8, 0), pady=4)

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(10, 8))
        ttk.Button(btns, text="ANALIZZA DDT", command=self.run_analysis).pack(side="left")
        self.price_btn = ttk.Button(btns, text="Prezzo prodotto selezionato", command=self.set_selected_price, state="disabled")
        self.price_btn.pack(side="left", padx=8)
        self.export_btn = ttk.Button(btns, text="Esporta CSV", command=self.export_csv, state="disabled")
        self.export_btn.pack(side="left")
        self.copy_full_btn = ttk.Button(btns, text="Copia lista completa", command=self.copy_full, state="disabled")
        self.copy_full_btn.pack(side="left", padx=8)
        self.copy_two_btn = ttk.Button(btns, text="Copia Prodotto + Pezzi", command=self.copy_name_qty, state="disabled")
        self.copy_two_btn.pack(side="left")

        ttk.Label(outer, textvariable=self.status, foreground="#344054").pack(anchor="w", pady=(3, 2))
        ttk.Label(outer, textvariable=self.summary, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 2))
        ttk.Label(outer, textvariable=self.order_line, font=("Segoe UI", 9, "bold"), foreground="#475467", wraplength=1060).pack(anchor="w", pady=(0, 8))

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)
        columns = ("code", "description", "qty", "price", "total")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("code", text="Codice")
        self.tree.heading("description", text="Prodotto")
        self.tree.heading("qty", text="Pezzi venduti")
        self.tree.heading("price", text="Prezzo €")
        self.tree.heading("total", text="Totale €")
        self.tree.column("code", width=115, anchor="center")
        self.tree.column("description", width=500, anchor="w")
        self.tree.column("qty", width=105, anchor="e")
        self.tree.column("price", width=100, anchor="e")
        self.tree.column("total", width=110, anchor="e")
        self.tree.tag_configure("missing", background="#fff4e5")
        self.tree.tag_configure("meta", background="#eef2f6")
        self.tree.bind("<Double-1>", lambda e: self.set_selected_price())

        ysb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        ttk.Label(outer, text="Il listino viene salvato accanto all'EXE come NMI_BeautyFly_listino.json. Doppio clic su un prodotto per assegnargli il prezzo.", font=("Segoe UI", 9), foreground="#667085").pack(anchor="w", pady=(8, 0))

    def choose_pdf(self):
        path = filedialog.askopenfilename(title="Seleziona PDF DDT", filetypes=[("PDF", "*.pdf"), ("Tutti i file", "*.*")])
        if path:
            self.pdf_path.set(path)
            self.status.set("PDF selezionato. Premi ANALIZZA DDT.")

    def clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def build_order_text(self, result):
        docs = list(result.get("documents", {}).keys())
        refs = result.get("order_refs", [])
        parts = []
        if docs:
            parts.append("DDT: " + ", ".join(docs))
        if refs:
            parts.append("Rif. ordini stampati: " + " | ".join(refs))
        return " — ".join(parts) if parts else "DDT / ordini: -"

    def enriched_products(self):
        if not self.result:
            return []
        rows = []
        for item in self.result.get("products", []):
            price, _ = get_price_for(self.prices, item["code"], item["description"])
            total = item["qty"] * price if price is not None else None
            rows.append({**item, "price": price, "total": total})
        return rows

    def refresh_table(self):
        self.clear_table()
        if not self.result:
            return

        meta_text = self.build_order_text(self.result)
        self.tree.insert("", "end", iid="__meta__", values=("DDT/ORDINI", meta_text, "", "", ""), tags=("meta",))

        rows = self.enriched_products()
        missing = 0
        total_value = 0.0
        for row in rows:
            tags = ()
            if row["price"] is None:
                missing += 1
                tags = ("missing",)
            else:
                total_value += row["total"]
            self.tree.insert("", "end", iid="prod_" + row["code"], values=(
                row["code"], row["description"], fmt_qty(row["qty"]),
                "MANCANTE" if row["price"] is None else fmt_money(row["price"]),
                "" if row["total"] is None else fmt_money(row["total"]),
            ), tags=tags)

        doc_count = len(self.result["documents"])
        product_count = len(rows)
        total_qty = fmt_qty(self.result["total_qty"])
        self.summary.set(
            "DDT distinti: %d   |   Prodotti: %d   |   Pezzi totali: %s   |   Valore: € %s   |   Prezzi mancanti: %d" %
            (doc_count, product_count, total_qty, fmt_money(total_value), missing)
        )
        self.order_line.set(meta_text)

    def run_analysis(self):
        path = self.pdf_path.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "Seleziona prima il PDF dei DDT.")
            return
        if not os.path.isfile(path):
            messagebox.showerror(APP_TITLE, "Il file PDF non esiste.")
            return
        customer = self.customer.get().strip()
        if not customer:
            messagebox.showwarning(APP_TITLE, "Inserisci il cliente da cercare.")
            return

        self.status.set("Analisi in corso...")
        self.update_idletasks()
        try:
            result = analyze_pdf(path, customer)
        except Exception as exc:
            self.result = None
            self.clear_table()
            self.summary.set("")
            self.order_line.set("DDT / ordini: -")
            for b in (self.export_btn, self.copy_full_btn, self.copy_two_btn, self.price_btn):
                b.configure(state="disabled")
            messagebox.showerror(APP_TITLE, "Errore durante la lettura del PDF:\n\n%s" % exc)
            self.status.set("Errore durante l'analisi.")
            return

        self.result = result
        if not result["pages"]:
            self.clear_table()
            self.status.set("Nessun DDT di vendita trovato per '%s'." % customer)
            self.summary.set("0 DDT - 0 prodotti")
            self.order_line.set("DDT / ordini: -")
            for b in (self.export_btn, self.copy_full_btn, self.copy_two_btn, self.price_btn):
                b.configure(state="disabled")
            return

        self.refresh_table()
        self.status.set("Analisi completata. Pagine PDF incluse: %s" % ", ".join(map(str, result["pages"])))
        for b in (self.export_btn, self.copy_full_btn, self.copy_two_btn, self.price_btn):
            b.configure(state="normal")

    def selected_product(self):
        sel = self.tree.selection()
        if not sel:
            return None
        iid = sel[0]
        if not iid.startswith("prod_"):
            return None
        code = iid[5:]
        if not self.result:
            return None
        for x in self.result.get("products", []):
            if x["code"] == code:
                return x
        return None

    def set_selected_price(self):
        item = self.selected_product()
        if item is None:
            if self.result:
                messagebox.showinfo(APP_TITLE, "Seleziona una riga prodotto (non la riga DDT/ordini).")
            return
        current_price, current = get_price_for(self.prices, item["code"], item["description"])
        seed = {"code": item["code"], "name": item["description"], "price": current_price if current_price is not None else 0.0}
        dlg = PriceDialog(self, seed)
        self.wait_window(dlg)
        if not dlg.result:
            return
        self.upsert_price(dlg.result)
        self.refresh_table()
        self.status.set("Prezzo salvato nel listino.")

    def upsert_price(self, item, old_index=None):
        if old_index is not None:
            self.prices[old_index] = item
        else:
            found = None
            for i, p in enumerate(self.prices):
                if item["code"] and clean_line(p.get("code", "")) == item["code"]:
                    found = i
                    break
                if norm_name(p.get("name", "")) == norm_name(item["name"]):
                    found = i
                    break
            if found is None:
                self.prices.append(item)
            else:
                self.prices[found] = item
        self.prices.sort(key=lambda x: x["name"].upper())
        save_prices(self.prices)

    def open_price_list(self):
        win = tk.Toplevel(self)
        win.title("Listino prezzi NMI")
        win.geometry("720x520")
        win.transient(self)
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Listino prezzi", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(frm, text="Puoi aggiungere nome prodotto e prezzo; il codice è facoltativo ma consigliato.").pack(anchor="w", pady=(2, 10))

        tv = ttk.Treeview(frm, columns=("code", "name", "price"), show="headings", selectmode="browse")
        tv.heading("code", text="Codice")
        tv.heading("name", text="Nome prodotto")
        tv.heading("price", text="Prezzo €")
        tv.column("code", width=110, anchor="center")
        tv.column("name", width=430)
        tv.column("price", width=110, anchor="e")
        tv.pack(fill="both", expand=True)

        def fill():
            for x in tv.get_children():
                tv.delete(x)
            for i, p in enumerate(self.prices):
                tv.insert("", "end", iid=str(i), values=(p.get("code", ""), p.get("name", ""), fmt_money(p.get("price", 0))))

        def add():
            dlg = PriceDialog(win)
            win.wait_window(dlg)
            if dlg.result:
                self.upsert_price(dlg.result)
                fill()
                if self.result:
                    self.refresh_table()

        def edit():
            sel = tv.selection()
            if not sel:
                return
            idx = int(sel[0])
            if idx >= len(self.prices):
                return
            dlg = PriceDialog(win, self.prices[idx])
            win.wait_window(dlg)
            if dlg.result:
                self.upsert_price(dlg.result, old_index=idx)
                fill()
                if self.result:
                    self.refresh_table()

        def delete():
            sel = tv.selection()
            if not sel:
                return
            idx = int(sel[0])
            if idx >= len(self.prices):
                return
            if messagebox.askyesno(APP_TITLE, "Eliminare questa voce dal listino?", parent=win):
                del self.prices[idx]
                save_prices(self.prices)
                fill()
                if self.result:
                    self.refresh_table()

        def import_csv():
            path = filedialog.askopenfilename(parent=win, title="Importa listino CSV", filetypes=[("CSV", "*.csv"), ("Tutti i file", "*.*")])
            if not path:
                return
            added = 0
            try:
                with open(path, "r", encoding="utf-8-sig", newline="") as f:
                    sample = f.read(4096)
                    f.seek(0)
                    delim = ";" if sample.count(";") >= sample.count(",") else ","
                    r = csv.reader(f, delimiter=delim)
                    for row in r:
                        if not row:
                            continue
                        low = [clean_line(x).lower() for x in row]
                        if any(x in ("nome", "prodotto", "descrizione") for x in low) and any("prezzo" in x for x in low):
                            continue
                        if len(row) >= 3:
                            code, name, price_s = row[0], row[1], row[2]
                        elif len(row) >= 2:
                            code, name, price_s = "", row[0], row[1]
                        else:
                            continue
                        name = clean_line(name)
                        if not name:
                            continue
                        try:
                            price = parse_price(price_s)
                        except Exception:
                            continue
                        self.upsert_price({"code": clean_line(code), "name": name, "price": price})
                        added += 1
                fill()
                if self.result:
                    self.refresh_table()
                messagebox.showinfo(APP_TITLE, f"Importate/aggiornate {added} voci.", parent=win)
            except Exception as exc:
                messagebox.showerror(APP_TITLE, str(exc), parent=win)

        def export_prices():
            path = filedialog.asksaveasfilename(parent=win, title="Esporta listino", defaultextension=".csv", initialfile="NMI_BeautyFly_listino.csv", filetypes=[("CSV", "*.csv")])
            if not path:
                return
            try:
                with open(path, "w", encoding="utf-8-sig", newline="") as f:
                    w = csv.writer(f, delimiter=";")
                    w.writerow(["Codice", "Nome prodotto", "Prezzo"])
                    for p in self.prices:
                        w.writerow([p.get("code", ""), p.get("name", ""), fmt_money(p.get("price", 0))])
                messagebox.showinfo(APP_TITLE, "Listino esportato.", parent=win)
            except Exception as exc:
                messagebox.showerror(APP_TITLE, str(exc), parent=win)

        b = ttk.Frame(frm)
        b.pack(fill="x", pady=(10, 0))
        ttk.Button(b, text="Aggiungi", command=add).pack(side="left")
        ttk.Button(b, text="Modifica", command=edit).pack(side="left", padx=6)
        ttk.Button(b, text="Elimina", command=delete).pack(side="left")
        ttk.Button(b, text="Importa CSV", command=import_csv).pack(side="left", padx=(18, 6))
        ttk.Button(b, text="Esporta CSV", command=export_prices).pack(side="left")
        ttk.Button(b, text="Chiudi", command=win.destroy).pack(side="right")
        tv.bind("<Double-1>", lambda e: edit())
        fill()

    def export_csv(self):
        if not self.result or not self.result.get("products"):
            return
        base = os.path.splitext(os.path.basename(self.pdf_path.get()))[0]
        suggested = "%s_BEAUTY_FLY_quantita_valore.csv" % base
        path = filedialog.asksaveasfilename(title="Esporta risultati", defaultextension=".csv", initialfile=suggested, filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            rows = self.enriched_products()
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["DDT / ORDINI STAMPATI", self.build_order_text(self.result)])
                w.writerow(["Cliente", self.customer.get().strip()])
                w.writerow(["File PDF", os.path.basename(self.pdf_path.get())])
                w.writerow([])
                w.writerow(["Codice", "Prodotto", "Pezzi venduti", "Prezzo unitario", "Totale"])
                total_value = 0.0
                for row in rows:
                    if row["total"] is not None:
                        total_value += row["total"]
                    w.writerow([row["code"], row["description"], fmt_qty(row["qty"]), "" if row["price"] is None else fmt_money(row["price"]), "" if row["total"] is None else fmt_money(row["total"])])
                w.writerow([])
                w.writerow(["", "TOTALE PEZZI", fmt_qty(self.result["total_qty"]), "TOTALE €", fmt_money(total_value)])
            messagebox.showinfo(APP_TITLE, "CSV salvato:\n%s" % path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, "Impossibile salvare il CSV:\n%s" % exc)

    def copy_full(self):
        if not self.result:
            return
        rows = ["DDT / ORDINI STAMPATI\t" + self.build_order_text(self.result), "Codice\tProdotto\tPezzi venduti\tPrezzo €\tTotale €"]
        total_value = 0.0
        for row in self.enriched_products():
            if row["total"] is not None:
                total_value += row["total"]
            rows.append("%s\t%s\t%s\t%s\t%s" % (
                row["code"], row["description"], fmt_qty(row["qty"]),
                "" if row["price"] is None else fmt_money(row["price"]),
                "" if row["total"] is None else fmt_money(row["total"]),
            ))
        rows.append("\tTOTALE PEZZI\t%s\tTOTALE €\t%s" % (fmt_qty(self.result["total_qty"]), fmt_money(total_value)))
        self.clipboard_clear()
        self.clipboard_append("\n".join(rows))
        self.status.set("Lista completa copiata negli appunti.")

    def copy_name_qty(self):
        if not self.result:
            return
        rows = ["Prodotto\tPezzi venduti"]
        for row in self.result.get("products", []):
            rows.append("%s\t%s" % (row["description"], fmt_qty(row["qty"])))
        self.clipboard_clear()
        self.clipboard_append("\n".join(rows))
        self.status.set("Copiate solo le colonne Prodotto + Pezzi venduti.")


def main():
    if fitz is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, "PyMuPDF non disponibile:\n%s" % FITZ_IMPORT_ERROR)
        root.destroy()
        return
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
