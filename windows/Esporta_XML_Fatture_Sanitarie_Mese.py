"""Esporta XML FPR12 delle fatture sanitarie NON elettroniche del solo mese scelto.

Programma di sola lettura: tutte le chiamate Fatture in Cloud sono GET.
Non usa STS come criterio di inclusione, non esegue alcun invio.
Non produce XML gestionale o surrogato FatturaPA: esporta ESCLUSIVAMENTE
l'XML nativo FPR12 restituito da Fatture in Cloud.
"""
from __future__ import annotations

import json
import re
import threading
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date, calendar
from pathlib import Path
from tkinter import Tk, ttk, StringVar, Text, filedialog, messagebox, END
from zipfile import ZipFile, ZIP_DEFLATED

API = "https://api-v2.fattureincloud.it"
MONTHS = (
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre",
)
NS = "http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"
APP_NAME = "Esporta XML Fatture Sanitarie - Mese"


class ExportError(Exception):
    pass


def token_clean(value: str) -> str:
    token = str(value or "").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    if not token or "traceback" in token.casefold() or any(
        ord(c) < 33 or ord(c) > 126 for c in token
    ):
        raise ExportError(
            "Token non valido. Cancella il campo e incolla soltanto il token API "
            "originale di Fatture in Cloud."
        )
    return token


def api_read(url: str, token: str, accept: str = "application/json") -> bytes:
    # Never forward the bearer token to a third-party URL.
    parts = urllib.parse.urlsplit(url)
    if parts.scheme != "https" or parts.netloc != "api-v2.fattureincloud.it":
        raise ExportError("Indirizzo API non consentito.")
    req = urllib.request.Request(
        url, headers={
            "Authorization": "Bearer " + token,
            "Accept": accept,
            "User-Agent": "FIC-Esporta-XML-Sanitari-Mese/1.0",
        }, method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=90) as response:
            return response.read()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")
        try:
            js = json.loads(detail)
            err = js.get("error") or {}
            detail = str(err.get("message") or detail)[:280]
        except ValueError:
            detail = detail[:280]
        if exc.code == 401:
            raise ExportError("401: token scaduto o errato.") from exc
        if exc.code == 403:
            raise ExportError("403: il token non ha i permessi di lettura.") from exc
        raise ExportError(f"HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise ExportError(f"Connessione non disponibile: {exc.reason}") from exc


def month_bounds(year: int, month: int) -> tuple[str, str]:
    if not 2000 <= year <= 2100 or not 1 <= month <= 12:
        raise ExportError("Anno o mese non valido.")
    return f"{year:04d}-{month:02d}-01", (
        f"{year:04d}-{month:02d}-{calendar.monthrange(year, month)[1]:02d}"
    )


def invoice_number(row: dict) -> str:
    num = row.get("number")
    numeration = str(row.get("numeration") or "")
    if num is None:
        return str(row.get("id") or "")
    return str(num) + numeration


def is_non_electronic(row: dict) -> bool:
    """Criterio numerazione, già verificato sui totali di riferimento:
    fatture elettroniche nel sezionale EL; altre fatture incluse.
    NON verificare o filtrare invio/pagamento TS.
    """
    if str(row.get("type") or "invoice").lower() != "invoice":
        return False
    number = invoice_number(row).upper()
    return "EL" not in number


def list_month_invoices(company: str, token: str, year: int, month: int) -> list[dict]:
    start, end = month_bounds(year, month)
    query = f"date >= '{start}' and date <= '{end}'"
    params = urllib.parse.urlencode({
        "type": "invoice",
        "q": query,
        "per_page": "100",
        "fieldset": "detailed",
        "sort": "date",
    })
    url = f"{API}/c/{urllib.parse.quote(company)}/issued_documents?{params}"
    rows: list[dict] = []
    seen = set()
    page = 0
    while url:
        page += 1
        if page > 1000:
            raise ExportError("Interrotto: paginazione API anomala.")
        try:
            js = json.loads(api_read(url, token))
        except ValueError as exc:
            raise ExportError("Risposta JSON non valida dall'API.") from exc
        data = js.get("data") or []
        if not isinstance(data, list):
            raise ExportError("Risposta fatture non valida.")
        for row in data:
            if not isinstance(row, dict):
                continue
            ident = str(row.get("id") or "")
            if ident in seen:
                continue
            if ident:
                seen.add(ident)
            # Second check: date still in the selected calendar month.
            if str(row.get("date") or "")[:7] == f"{year:04d}-{month:02d}":
                rows.append(row)
        next_url = js.get("next_page_url")
        if next_url:
            url = urllib.parse.urljoin(API, str(next_url))
        else:
            url = None
    return sorted(
        (r for r in rows if is_non_electronic(r)),
        key=lambda r: (str(r.get("date") or ""), invoice_number(r), str(r.get("id") or "")),
    )


def validate_fpr12(xml_bytes: bytes) -> None:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        raise ExportError("Il documento restituito non contiene un XML valido.") from exc
    if root.tag != "{" + NS + "}FatturaElettronica":
        raise ExportError(
            "Il documento non ha la struttura FatturaElettronica FPR12 richiesta."
        )
    version = str(root.attrib.get("versione") or "")
    transport = root.findtext(".//FormatoTrasmissione")
    if version != "FPR12" or transport != "FPR12":
        raise ExportError(
            f"Formato diverso dall'esempio (versione={version}, trasmissione={transport}). "
            "Non esporto un documento FPA12 o un XML di altro tipo."
        )
    if root.find("FatturaElettronicaHeader") is None or root.find("FatturaElettronicaBody") is None:
        raise ExportError("XML incompleto: mancano intestazione o corpo.")


def fetch_native_xml(company: str, token: str, row: dict) -> bytes:
    doc_id = str(row.get("id") or "")
    if not doc_id.isdigit():
        raise ExportError("ID del documento mancante o non valido.")
    # Richiede unicamente l'XML esattamente prodotto da Fatture in Cloud,
    # anche se la fattura NON e' elettronica. Non effettua mai invii.
    url = (
        f"{API}/c/{urllib.parse.quote(company)}/issued_documents/"
        f"{doc_id}/e_invoice/xml?avoid_einvoice_check=1&include_attachment=false"
    )
    contents = api_read(url, token, "application/xml,text/xml,*/*")
    validate_fpr12(contents)
    return contents


def safe_filename(row: dict, company: str) -> str:
    label = invoice_number(row)
    safe_num = re.sub(r"[^a-zA-Z0-9._-]+", "_", label).strip("._") or "senza_numero"
    stamp = str(row.get("date") or "").replace("-", "")
    return f"IT_{company}_{stamp}_{safe_num}_id{row.get('id')}.xml"


def export_month(company: str, token: str, year: int, month: int,
                 output_zip: Path, progress=None) -> tuple[int, int, list[str]]:
    rows = list_month_invoices(company, token, year, month)
    if progress:
        progress(f"Fatture non elettroniche trovate: {len(rows)}")
    if not rows:
        raise ExportError("Nessuna fattura non elettronica trovata per il mese scelto.")
    exported = 0
    failures: list[str] = []
    output_zip.parent.mkdir(parents=True, exist_ok=True)
    # Output is an ordinary ZIP of original XMLs; no substitute XML is generated.
    with ZipFile(output_zip, "w", compression=ZIP_DEFLATED) as archive:
        filenames = set()
        for index, doc in enumerate(rows, start=1):
            num = invoice_number(doc)
            if progress:
                progress(f"{index}/{len(rows)} — fattura {num}")
            try:
                content = fetch_native_xml(company, token, doc)
                name = safe_filename(doc, company)
                if name in filenames:
                    raise ExportError("Nome XML duplicato.")
                archive.writestr(name, content)
                filenames.add(name)
                exported += 1
            except Exception as exc:
                failures.append(f"Fattura {num} (ID {doc.get('id')}): {exc}")
                if progress:
                    progress("NON esportata: " + num + " — " + str(exc)[:180])
    # No partial export is presented as complete.
    if failures:
        report = output_zip.with_suffix(".errori.txt")
        report.write_text(
            f"Esportazione {MONTHS[month-1]} {year}\n"
            f"Fatture totali: {len(rows)}; XML FPR12 salvati: {exported}; errori: {len(failures)}\n\n"
            + "\n".join(failures), encoding="utf-8",
        )
        if progress:
            progress(f"ATTENZIONE: {len(failures)} fatture non esportate. Dettagli in {report.name}")
    else:
        report = output_zip.with_suffix(".errori.txt")
        if report.exists():
            report.unlink()
    return len(rows), exported, failures


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("840x440")
        self.minsize(690, 360)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.company = StringVar()
        self.token = StringVar()
        self.year = StringVar(value=str(date.today().year))
        self.month = StringVar(value=MONTHS[date.today().month - 1])
        self.status = StringVar(value="Seleziona anno e mese, poi premi ESPORTA XML DEL MESE.")
        self._ui()

    def _ui(self):
        frame = ttk.Frame(self, padding=18)
        frame.grid(row=0, column=0, sticky="ew")
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)

        ttk.Label(frame, text="ID azienda Fatture in Cloud").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.company).grid(
            row=0, column=1, sticky="ew", padx=(8, 24))
        ttk.Label(frame, text="Token personale").grid(row=0, column=2, sticky="w")
        ttk.Entry(frame, textvariable=self.token, show="•").grid(
            row=0, column=3, sticky="ew", padx=(8, 0))

        ttk.Label(frame, text="Anno").grid(row=1, column=0, sticky="w", pady=(14, 0))
        ttk.Combobox(frame, textvariable=self.year, state="readonly",
                     values=[str(y) for y in range(date.today().year-5, date.today().year+2)],
                     width=12).grid(row=1, column=1, sticky="w", padx=(8, 24), pady=(14, 0))
        ttk.Label(frame, text="Mese da esportare").grid(row=1, column=2, sticky="w", pady=(14, 0))
        ttk.Combobox(frame, textvariable=self.month, state="readonly",
                     values=MONTHS, width=20).grid(
                         row=1, column=3, sticky="w", padx=(8, 0), pady=(14, 0))

        self.start_button = ttk.Button(
            frame, text="ESPORTA XML DEL MESE", command=self.start)
        self.start_button.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(20, 8))
        ttk.Label(frame, textvariable=self.status, wraplength=750).grid(
            row=3, column=0, columnspan=4, sticky="w", pady=(3, 0))

        output = ttk.Frame(self, padding=(18, 0, 18, 14))
        output.grid(row=1, column=0, sticky="nsew")
        output.columnconfigure(0, weight=1)
        output.rowconfigure(0, weight=1)
        self.log = Text(output, wrap="word", height=11, state="disabled")
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(output, orient="vertical", command=self.log.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

    def log_line(self, value: str):
        self.after(0, lambda: self._append(value))

    def _append(self, line: str):
        self.log.configure(state="normal")
        self.log.insert(END, line + "\n")
        self.log.see(END)
        self.log.configure(state="disabled")

    def start(self):
        try:
            company = self.company.get().strip()
            if not company.isdigit():
                raise ExportError("Inserisci un ID azienda numerico.")
            token = token_clean(self.token.get())
            year = int(self.year.get())
            month = MONTHS.index(self.month.get()) + 1
            month_bounds(year, month)
        except (ExportError, ValueError) as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return

        path = filedialog.asksaveasfilename(
            title="Salva tutte le fatture sanitarie XML del mese in ZIP",
            defaultextension=".zip", filetypes=[("ZIP di file XML", "*.zip")],
            initialfile=f"Fatture_Sanitarie_{company}_{year}_{month:02d}.zip",
        )
        if not path:
            return
        self.start_button.configure(state="disabled")
        self.status.set("Lettura fatture in corso…")
        self._append(f"Esporto: {MONTHS[month - 1]} {year} — azienda {company}")
        threading.Thread(
            target=self._work,
            args=(company, token, year, month, Path(path)),
            daemon=True,
        ).start()

    def _work(self, company, token, year, month, path):
        try:
            total, count, failures = export_month(
                company, token, year, month, path, self.log_line)
            if failures:
                msg = (f"INCOMPLETO: {count}/{total} XML FPR12 esportati. "
                       f"Verifica il file .errori.txt accanto allo ZIP.")
            else:
                msg = f"COMPLETATO: {count}/{total} XML FPR12 in {path}"
            self.after(0, lambda: self.status.set(msg))
            self.log_line(msg)
        except Exception as exc:
            message = str(exc)
            self.after(0, lambda: self.status.set("Errore: " + message))
            self.log_line("Errore: " + message)
        finally:
            self.after(0, lambda: self.start_button.configure(state="normal"))


if __name__ == "__main__":
    App().mainloop()
