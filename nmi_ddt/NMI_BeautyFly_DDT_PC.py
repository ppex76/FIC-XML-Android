import os
import re
import csv
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import fitz  # PyMuPDF
except Exception as exc:
    fitz = None
    FITZ_IMPORT_ERROR = exc
else:
    FITZ_IMPORT_ERROR = None

APP_TITLE = "NMI - Quantita DDT Beauty Fly"
DEFAULT_CUSTOMER = "BEAUTY FLY"
NMI_MARKER = "NMI DI PAOLO PEZZOTTA"


def clean_line(s):
    return " ".join(str(s).replace("\xa0", " ").split()).strip()


def fmt_qty(v):
    try:
        f = float(v)
    except Exception:
        return str(v)
    if abs(f - round(f)) < 1e-9:
        return str(int(round(f)))
    return ("%.3f" % f).rstrip("0").rstrip(".").replace(".", ",")


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
            items.append({
                "code": code,
                "description": desc,
                "qty": qty,
                "um": um.upper(),
            })
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
    row_count = 0

    try:
        for page_index in range(len(doc)):
            text = doc[page_index].get_text("text") or ""
            up = text.upper()

            if key not in up:
                continue
            if "VENDITE" not in up:
                continue
            if NMI_MARKER not in up:
                continue

            page_no = page_index + 1
            selected_pages.append(page_no)
            document_no, document_date = extract_document_info(text, key)
            if document_no:
                docs.setdefault(document_no, document_date)

            items = parse_products_from_page(text)
            row_count += len(items)
            for item in items:
                code = item["code"]
                if code not in totals:
                    totals[code] = {
                        "code": code,
                        "description": item["description"],
                        "qty": 0.0,
                        "um": item["um"],
                    }
                totals[code]["qty"] += item["qty"]
                if len(item["description"]) > len(totals[code]["description"]):
                    totals[code]["description"] = item["description"]

        results = sorted(totals.values(), key=lambda x: (x["description"].upper(), x["code"]))
        total_qty = sum(x["qty"] for x in results)
        return {
            "products": results,
            "pages": selected_pages,
            "documents": docs,
            "row_count": row_count,
            "total_qty": total_qty,
            "total_pdf_pages": len(doc),
            "customer": customer_key,
        }
    finally:
        doc.close()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("940x650")
        self.minsize(760, 520)
        self.result = None
        self.pdf_path = tk.StringVar()
        self.customer = tk.StringVar(value=DEFAULT_CUSTOMER)
        self.status = tk.StringVar(value="Seleziona il PDF dei DDT di vendita.")
        self.summary = tk.StringVar(value="")
        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="NMI - Quantita prodotti per cliente", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Legge i DDT di VENDITA, prende solo le pagine del cliente scelto e somma la colonna Quantita per prodotto.",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(2, 14))

        frm = ttk.Frame(outer)
        frm.pack(fill="x")
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Cliente:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.customer).grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)

        ttk.Label(frm, text="PDF DDT:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.pdf_path, state="readonly").grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(frm, text="Scegli PDF", command=self.choose_pdf).grid(row=1, column=2, padx=(8, 0), pady=4)

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(12, 8))
        ttk.Button(btns, text="ANALIZZA DDT", command=self.run_analysis).pack(side="left")
        self.export_btn = ttk.Button(btns, text="Esporta CSV", command=self.export_csv, state="disabled")
        self.export_btn.pack(side="left", padx=8)
        self.copy_btn = ttk.Button(btns, text="Copia elenco", command=self.copy_results, state="disabled")
        self.copy_btn.pack(side="left")

        ttk.Label(outer, textvariable=self.status, foreground="#344054").pack(anchor="w", pady=(4, 4))
        ttk.Label(outer, textvariable=self.summary, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)

        columns = ("code", "description", "qty")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.tree.heading("code", text="Codice")
        self.tree.heading("description", text="Prodotto")
        self.tree.heading("qty", text="Quantita")
        self.tree.column("code", width=120, anchor="center")
        self.tree.column("description", width=580, anchor="w")
        self.tree.column("qty", width=110, anchor="e")

        ysb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        ttk.Label(
            outer,
            text="Nota: il programma usa la Quantita dei DDT con causale VENDITE; non somma lotti, prezzi o righe di riferimento ordine.",
            font=("Segoe UI", 9),
            foreground="#667085",
        ).pack(anchor="w", pady=(8, 0))

    def choose_pdf(self):
        path = filedialog.askopenfilename(title="Seleziona PDF DDT", filetypes=[("PDF", "*.pdf"), ("Tutti i file", "*.*")])
        if path:
            self.pdf_path.set(path)
            self.status.set("PDF selezionato. Premi ANALIZZA DDT.")

    def clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

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
            self.export_btn.configure(state="disabled")
            self.copy_btn.configure(state="disabled")
            messagebox.showerror(APP_TITLE, "Errore durante la lettura del PDF:\n\n%s" % exc)
            self.status.set("Errore durante l'analisi.")
            return

        self.result = result
        self.clear_table()
        for row in result["products"]:
            self.tree.insert("", "end", values=(row["code"], row["description"], fmt_qty(row["qty"])))

        doc_count = len(result["documents"])
        page_count = len(result["pages"])
        product_count = len(result["products"])
        total_qty = fmt_qty(result["total_qty"])

        if page_count == 0:
            self.status.set("Nessun DDT di vendita trovato per '%s'." % customer)
            self.summary.set("0 DDT - 0 prodotti")
            self.export_btn.configure(state="disabled")
            self.copy_btn.configure(state="disabled")
            return

        self.status.set("Analisi completata. Pagine PDF incluse: %s" % ", ".join(map(str, result["pages"])))
        self.summary.set("DDT distinti: %d   |   Prodotti: %d   |   Quantita totale: %s" % (doc_count, product_count, total_qty))
        self.export_btn.configure(state="normal")
        self.copy_btn.configure(state="normal")

    def export_csv(self):
        if not self.result or not self.result.get("products"):
            return
        base = os.path.splitext(os.path.basename(self.pdf_path.get()))[0]
        suggested = "%s_BEAUTY_FLY_quantita.csv" % base
        path = filedialog.asksaveasfilename(
            title="Esporta risultati",
            defaultextension=".csv",
            initialfile=suggested,
            filetypes=[("CSV", "*.csv")],
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["Cliente", self.customer.get().strip()])
                w.writerow(["File PDF", os.path.basename(self.pdf_path.get())])
                w.writerow(["DDT distinti", len(self.result["documents"])])
                w.writerow(["Pagine incluse", ", ".join(map(str, self.result["pages"]))])
                w.writerow([])
                w.writerow(["Codice", "Prodotto", "Quantita"])
                for row in self.result["products"]:
                    w.writerow([row["code"], row["description"], fmt_qty(row["qty"])])
                w.writerow([])
                w.writerow(["", "TOTALE PEZZI", fmt_qty(self.result["total_qty"])])
            messagebox.showinfo(APP_TITLE, "CSV salvato:\n%s" % path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, "Impossibile salvare il CSV:\n%s" % exc)

    def copy_results(self):
        if not self.result or not self.result.get("products"):
            return
        rows = ["Codice\tProdotto\tQuantita"]
        for row in self.result["products"]:
            rows.append("%s\t%s\t%s" % (row["code"], row["description"], fmt_qty(row["qty"])))
        rows.append("\tTOTALE PEZZI\t%s" % fmt_qty(self.result["total_qty"]))
        self.clipboard_clear()
        self.clipboard_append("\n".join(rows))
        self.status.set("Elenco copiato negli appunti.")


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
