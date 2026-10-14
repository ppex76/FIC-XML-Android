import json
import threading
import urllib.error
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path
import os
import tkinter as tk
from tkinter import ttk, messagebox

API = "https://api-v2.fattureincloud.it"
APP = "NMI — Invio Fatture Elettroniche"
VERSION = "1.0"


class ApiError(Exception):
    def __init__(self, message, status=None):
        super().__init__(message)
        self.status = status


def _token_value(token):
    token = (token or "").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    return token


def api_request(method, url, token, payload=None):
    body = None
    headers = {
        "Authorization": "Bearer " + _token_value(token),
        "Accept": "application/json",
        "User-Agent": "NMI-SDI-PC/1.0",
    }
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"

    req = urllib.request.Request(url, data=body, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            raw = r.read()
            if not raw:
                return r.status, {}
            return r.status, json.loads(raw.decode("utf-8"))
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        detail = ""
        try:
            js = json.loads(raw)
            err = js.get("error") or {}
            msg = err.get("message") or ""
            validation = err.get("validation_result") or []
            if isinstance(validation, list) and validation:
                detail = msg + ": " + " | ".join(str(x) for x in validation)
            elif msg:
                detail = msg
            else:
                detail = raw[:500]
        except Exception:
            detail = raw[:500]

        if e.code == 401:
            raise ApiError("401 — token non valido o scaduto", e.code)
        if e.code == 403:
            raise ApiError("403 — permessi insufficienti per questa operazione", e.code)
        if e.code == 404:
            raise ApiError("404 — documento o azienda non trovato", e.code)
        if e.code == 422:
            raise ApiError("422 — " + (detail or "fattura non valida per l'invio"), e.code)
        raise ApiError(f"{e.code} — {detail or 'errore API'}", e.code)
    except Exception as e:
        if isinstance(e, ApiError):
            raise
        raise ApiError(str(e))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"NMI — Invio SDI PC v{VERSION}")
        self.geometry("1120x720")
        self.minsize(940, 600)
        self.candidates = []
        self.busy = False
        self.cfg = Path(os.getenv("APPDATA", str(Path.home()))) / "NMI_Invio_SDI_PC.json"
        self._build_ui()
        self._load_cfg()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=14)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)

        ttk.Label(top, text="ID azienda NMI").grid(row=0, column=0, sticky="w")
        self.company = tk.StringVar()
        ttk.Entry(top, textvariable=self.company, width=22).grid(row=0, column=1, sticky="ew", padx=(8, 20))

        ttk.Label(top, text="Token personale").grid(row=0, column=2, sticky="w")
        self.token = tk.StringVar()
        self.token_entry = ttk.Entry(top, textvariable=self.token, show="•")
        self.token_entry.grid(row=0, column=3, sticky="ew", padx=(8, 8))
        self.show_token = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top,
            text="Mostra",
            variable=self.show_token,
            command=lambda: self.token_entry.config(show="" if self.show_token.get() else "•"),
        ).grid(row=0, column=4, sticky="w")

        ttk.Label(top, text="Dal (AAAA-MM-GG)").grid(row=1, column=0, sticky="w", pady=(12, 0))
        self.date_from = tk.StringVar(value=f"{date.today().year}-01-01")
        ttk.Entry(top, textvariable=self.date_from, width=16).grid(row=1, column=1, sticky="w", padx=(8, 20), pady=(12, 0))

        ttk.Label(top, text="Al (AAAA-MM-GG)").grid(row=1, column=2, sticky="w", pady=(12, 0))
        self.date_to = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(top, textvariable=self.date_to, width=16).grid(row=1, column=3, sticky="w", padx=(8, 20), pady=(12, 0))

        self.btn_verify = ttk.Button(top, text="Verifica collegamento", command=self.verify_connection)
        self.btn_verify.grid(row=2, column=0, pady=(14, 0), sticky="w")

        self.btn_check = ttk.Button(top, text="Controlla fatture da spedire", command=self.check_only)
        self.btn_check.grid(row=2, column=1, pady=(14, 0), padx=(8, 8), sticky="w")

        self.btn_dry = ttk.Button(top, text="TEST senza inviare", command=self.dry_run_all)
        self.btn_dry.grid(row=2, column=2, pady=(14, 0), padx=(0, 8), sticky="w")

        self.btn_send = ttk.Button(top, text="CONTROLLA E SPEDISCI", command=self.check_and_send)
        self.btn_send.grid(row=2, column=3, pady=(14, 0), sticky="w")

        mid = ttk.Frame(self, padding=(14, 0, 14, 6))
        mid.grid(row=1, column=0, sticky="nsew")
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(0, weight=1)

        cols = ("numero", "data", "cliente", "totale", "stato_fic", "esito")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        definitions = [
            ("numero", "Numero", 90),
            ("data", "Data", 100),
            ("cliente", "Cliente", 300),
            ("totale", "Totale", 100),
            ("stato_fic", "Stato FIC", 120),
            ("esito", "Esito", 330),
        ]
        for c, title, width in definitions:
            self.tree.heading(c, text=title)
            self.tree.column(c, width=width, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)

        bottom = ttk.Frame(self, padding=14)
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(4, weight=1)

        self.count_checked = tk.StringVar(value="0")
        self.count_ready = tk.StringVar(value="0")
        self.count_sent = tk.StringVar(value="0")
        self.count_errors = tk.StringVar(value="0")
        ttk.Label(bottom, text="Controllate:").grid(row=0, column=0, sticky="w")
        ttk.Label(bottom, textvariable=self.count_checked).grid(row=0, column=1, padx=(5, 18), sticky="w")
        ttk.Label(bottom, text="Da spedire:").grid(row=0, column=2, sticky="w")
        ttk.Label(bottom, textvariable=self.count_ready).grid(row=0, column=3, padx=(5, 18), sticky="w")

        self.status = tk.StringVar(value="Pronto.")
        ttk.Label(bottom, textvariable=self.status, anchor="e").grid(row=0, column=4, sticky="ew")

        ttk.Label(bottom, text="Inviate:").grid(row=1, column=0, sticky="w", pady=(7, 0))
        ttk.Label(bottom, textvariable=self.count_sent).grid(row=1, column=1, padx=(5, 18), sticky="w", pady=(7, 0))
        ttk.Label(bottom, text="Errori/saltate:").grid(row=1, column=2, sticky="w", pady=(7, 0))
        ttk.Label(bottom, textvariable=self.count_errors).grid(row=1, column=3, padx=(5, 18), sticky="w", pady=(7, 0))
        ttk.Label(
            bottom,
            text="Il token non viene salvato. L'invio reale richiede conferma.",
            anchor="e",
        ).grid(row=1, column=4, sticky="e", pady=(7, 0))

    def _load_cfg(self):
        try:
            d = json.loads(self.cfg.read_text("utf-8"))
            self.company.set(d.get("company_id", ""))
            self.date_from.set(d.get("date_from", self.date_from.get()))
            self.date_to.set(d.get("date_to", self.date_to.get()))
        except Exception:
            pass

    def _save_cfg(self):
        try:
            self.cfg.write_text(
                json.dumps({
                    "company_id": self.company.get().strip(),
                    "date_from": self.date_from.get().strip(),
                    "date_to": self.date_to.get().strip(),
                }),
                encoding="utf-8",
            )
        except Exception:
            pass

    def creds(self):
        company = self.company.get().strip()
        token = _token_value(self.token.get())
        if not company:
            raise ApiError("Inserisci l'ID azienda NMI.")
        if not token:
            raise ApiError("Inserisci il token personale.")
        self._save_cfg()
        return company, token

    def date_range(self):
        start = self.date_from.get().strip()
        end = self.date_to.get().strip()
        try:
            d1 = date.fromisoformat(start)
            d2 = date.fromisoformat(end)
        except Exception:
            raise ApiError("Le date devono essere nel formato AAAA-MM-GG.")
        if d1 > d2:
            raise ApiError("La data iniziale è successiva alla data finale.")
        return start, end

    def set_status(self, text):
        self.after(0, lambda: self.status.set(text))

    def set_busy(self, value):
        self.busy = value
        state = "disabled" if value else "normal"
        def apply():
            for b in (self.btn_verify, self.btn_check, self.btn_dry, self.btn_send):
                b.config(state=state)
        self.after(0, apply)

    def run_bg(self, fn):
        threading.Thread(target=fn, daemon=True).start()

    def verify_connection(self):
        def job():
            self.set_busy(True)
            try:
                c, t = self.creds()
                self.set_status("Verifico il collegamento...")
                _, js = api_request("GET", f"{API}/c/{urllib.parse.quote(c)}/company/info", t)
                name = (js.get("data") or {}).get("name") or "azienda"
                self.set_status("Collegamento riuscito: " + str(name))
            except Exception as e:
                self.set_status("Errore: " + str(e))
            finally:
                self.set_busy(False)
        self.run_bg(job)

    def scan_candidates(self, company, token):
        start, end = self.date_range()
        page = 1
        last_page = 1
        scanned = 0
        ready = []
        fields = "id,type,entity,date,number,numeration,e_invoice,ei_status,amount_gross"

        while page <= last_page:
            self.set_status(f"Controllo pagina {page}...")
            q = f"date >= '{start}' and date <= '{end}'"
            params = {
                "type": "invoice",
                "q": q,
                "page": str(page),
                "per_page": "100",
                "fields": fields,
                "sort": "date",
            }
            url = f"{API}/c/{urllib.parse.quote(company)}/issued_documents?{urllib.parse.urlencode(params)}"
            _, js = api_request("GET", url, token)
            data = js.get("data") or []
            last_page = int(js.get("last_page") or 1)
            scanned += len(data)

            for d in data:
                e_invoice = d.get("e_invoice")
                is_e = e_invoice is True or e_invoice == 1 or str(e_invoice).lower() == "true"
                status = str(d.get("ei_status") or "").lower()
                if is_e and status == "not_sent":
                    d["result"] = "Da verificare"
                    ready.append(d)
            page += 1

        return scanned, ready

    def render(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, d in enumerate(self.candidates):
            entity = d.get("entity") or {}
            amount = d.get("amount_gross")
            if isinstance(amount, (int, float)):
                amount_text = f"€ {amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                amount_text = ""
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    str(d.get("number") or ""),
                    str(d.get("date") or ""),
                    str(entity.get("name") or ""),
                    amount_text,
                    str(d.get("ei_status") or ""),
                    str(d.get("result") or "Da verificare"),
                ),
            )
        self.count_ready.set(str(len(self.candidates)))

    def set_row_result(self, idx, text):
        self.candidates[idx]["result"] = text
        self.after(0, lambda: self.tree.set(str(idx), "esito", text))

    def check_only(self):
        def job():
            self.set_busy(True)
            try:
                c, t = self.creds()
                scanned, ready = self.scan_candidates(c, t)
                self.candidates = ready
                self.after(0, self.render)
                self.after(0, lambda: self.count_checked.set(str(scanned)))
                self.after(0, lambda: self.count_sent.set("0"))
                self.after(0, lambda: self.count_errors.set("0"))
                if ready:
                    self.set_status(f"Trovate {len(ready)} fatture elettroniche da spedire.")
                else:
                    self.set_status("Nessuna fattura elettronica da spedire nel periodo selezionato.")
            except Exception as e:
                self.set_status("Errore: " + str(e))
            finally:
                self.set_busy(False)
        self.run_bg(job)

    def verify_xml(self, company, token, document_id):
        url = f"{API}/c/{urllib.parse.quote(company)}/issued_documents/{document_id}/e_invoice/xml_verify"
        api_request("GET", url, token)
        return True

    def send_invoice(self, company, token, document_id, dry_run=False):
        url = f"{API}/c/{urllib.parse.quote(company)}/issued_documents/{document_id}/e_invoice/send"
        payload = {"options": {"dry_run": True}} if dry_run else {}
        return api_request("POST", url, token, payload)

    def process_all(self, dry_run=False):
        if not self.candidates:
            messagebox.showinfo(APP, "Prima controlla le fatture da spedire.")
            return

        if not dry_run:
            ok = messagebox.askyesno(
                APP,
                f"Stai per inviare realmente allo SDI {len(self.candidates)} fatture elettroniche.\n\n"
                "Ogni fattura verrà prima verificata. Quelle non valide saranno saltate.\n\nProcedere?",
                icon="warning",
            )
            if not ok:
                return

        def job():
            self.set_busy(True)
            sent = 0
            errors = 0
            try:
                c, t = self.creds()
                total = len(self.candidates)
                for idx, d in enumerate(self.candidates):
                    number = str(d.get("number") or d.get("id") or "")
                    self.set_status(f"{idx+1}/{total} — verifico fattura {number}...")
                    self.set_row_result(idx, "Verifica XML...")
                    try:
                        self.verify_xml(c, t, d.get("id"))
                    except Exception as e:
                        errors += 1
                        self.set_row_result(idx, "SALTATA — " + str(e))
                        continue

                    try:
                        if dry_run:
                            self.set_status(f"{idx+1}/{total} — test fattura {number}...")
                            self.set_row_result(idx, "Test invio...")
                            self.send_invoice(c, t, d.get("id"), dry_run=True)
                            self.set_row_result(idx, "✓ TEST OK — non inviata")
                        else:
                            self.set_status(f"{idx+1}/{total} — invio fattura {number} allo SDI...")
                            self.set_row_result(idx, "Invio allo SDI...")
                            self.send_invoice(c, t, d.get("id"), dry_run=False)
                            sent += 1
                            self.set_row_result(idx, "✓ RICHIESTA DI INVIO ACCETTATA")
                    except Exception as e:
                        errors += 1
                        self.set_row_result(idx, "ERRORE — " + str(e))

                self.after(0, lambda: self.count_sent.set(str(sent)))
                self.after(0, lambda: self.count_errors.set(str(errors)))
                if dry_run:
                    self.set_status(f"Test completato: {total-errors} OK, {errors} con errore. Nessuna fattura inviata.")
                else:
                    self.set_status(f"Completato: {sent} richieste di invio accettate, {errors} errori/saltate.")
            except Exception as e:
                self.set_status("Errore generale: " + str(e))
            finally:
                self.set_busy(False)
        self.run_bg(job)

    def dry_run_all(self):
        if not self.candidates:
            # Prima carica la lista, poi il pulsante può essere premuto di nuovo.
            messagebox.showinfo(APP, "Premi prima 'Controlla fatture da spedire'.")
            return
        self.process_all(dry_run=True)

    def check_and_send(self):
        def job():
            self.set_busy(True)
            try:
                c, t = self.creds()
                scanned, ready = self.scan_candidates(c, t)
                self.candidates = ready
                self.after(0, self.render)
                self.after(0, lambda: self.count_checked.set(str(scanned)))
                self.after(0, lambda: self.count_sent.set("0"))
                self.after(0, lambda: self.count_errors.set("0"))
                if not ready:
                    self.set_status("Nessuna fattura elettronica da spedire nel periodo selezionato.")
                    return
                self.set_status(f"Trovate {len(ready)} fatture elettroniche da spedire.")
                self.after(0, lambda: self.process_all(dry_run=False))
            except Exception as e:
                self.set_status("Errore: " + str(e))
            finally:
                self.set_busy(False)
        self.run_bg(job)


if __name__ == "__main__":
    App().mainloop()
