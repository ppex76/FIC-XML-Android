import json
import os
import urllib.parse
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

import NMI_Invio_SDI_PC as base

base.APP = "Invio SDI - Multi Azienda"


class App(base.App):
    """Invio elettroniche allo SDI con anagrafiche aziendali modificabili."""

    def _build_ui(self):
        super()._build_ui()
        self.title("Invio SDI - Multi Azienda PC v2")
        self.geometry("1170x780")
        self.cfg = Path(os.getenv("APPDATA", str(Path.home()))) / "Invio_SDI_MultiAzienda_PC.json"
        self.profiles = {}
        self.company_name = tk.StringVar(value="")
        self.actual_company_name = ""
        self.actual_company_id = None
        self.candidates_company_id = None

        top = self.btn_verify.master
        for child in top.winfo_children():
            if isinstance(child, ttk.Label) and child.cget("text") == "ID azienda NMI":
                child.configure(text="ID azienda")

        ttk.Label(top, text="Nome azienda").grid(row=3, column=0, sticky="w", pady=(12, 0))
        self.profile_combo = ttk.Combobox(top, textvariable=self.company_name, width=28)
        self.profile_combo.grid(row=3, column=1, sticky="ew", padx=(8, 20), pady=(12, 0))
        self.profile_combo.bind("<<ComboboxSelected>>", self.select_profile)
        ttk.Button(top, text="SALVA / AGGIORNA AZIENDA",
                   command=self.save_profile).grid(row=3, column=2, sticky="w", pady=(12, 0))
        ttk.Button(top, text="Elimina azienda salvata",
                   command=self.delete_profile).grid(row=3, column=3, sticky="w", pady=(12, 0))

        self.verified_label = tk.StringVar(value="Azienda FIC non ancora verificata")
        ttk.Label(top, textvariable=self.verified_label,
                  foreground="#344054").grid(row=4, column=0, columnspan=5,
                                              sticky="w", pady=(9, 0))
        self.company.trace_add("write", self._clear_verification)

    def _clear_verification(self, *args):
        self.actual_company_id = None
        self.actual_company_name = ""
        self.verified_label.set("Azienda FIC non ancora verificata")

    def _refresh_profiles(self):
        self.profile_combo.configure(values=sorted(self.profiles, key=str.casefold))

    def select_profile(self, event=None):
        name = self.company_name.get().strip()
        if name in self.profiles:
            self.company.set(self.profiles[name])
            self._save_cfg()

    def save_profile(self):
        name, cid = self.company_name.get().strip(), self.company.get().strip()
        if not name or not cid:
            messagebox.showwarning(base.APP, "Inserisci nome azienda e ID azienda.")
            return
        if not cid.isdigit():
            messagebox.showwarning(base.APP, "L'ID azienda deve essere numerico.")
            return
        self.profiles[name] = cid
        self._refresh_profiles()
        self._save_cfg()
        self.set_status(f"Salvata azienda: {name} (ID {cid}). Token non salvato.")

    def delete_profile(self):
        name = self.company_name.get().strip()
        if name not in self.profiles:
            messagebox.showinfo(base.APP, "Questa azienda non risulta salvata.")
            return
        if not messagebox.askyesno(base.APP, f"Eliminare il profilo locale '{name}'?"):
            return
        del self.profiles[name]
        self.company_name.set("")
        self.company.set("")
        self._refresh_profiles()
        self._save_cfg()
        self.set_status("Profilo locale eliminato. Nessuna fattura modificata.")

    def _load_cfg(self):
        try:
            if self.cfg.is_file():
                data = json.loads(self.cfg.read_text("utf-8"))
            else:
                old = self.cfg.with_name("NMI_Invio_SDI_PC.json")
                data = json.loads(old.read_text("utf-8")) if old.is_file() else {}
            p = data.get("profiles")
            if isinstance(p, dict):
                self.profiles = {
                    str(k): str(v) for k, v in p.items()
                    if str(k).strip() and str(v).strip()
                }
            cid = str(data.get("company_id") or "")
            name = str(data.get("company_name") or ("NMI" if cid else ""))
            self.company_name.set(name)
            self.company.set(cid)
            self.date_from.set(data.get("date_from") or self.date_from.get())
            self.date_to.set(data.get("date_to") or self.date_to.get())
            self._refresh_profiles()
        except Exception:
            pass

    def _save_cfg(self):
        try:
            self.cfg.write_text(json.dumps({
                "company_name": self.company_name.get().strip(),
                "company_id": self.company.get().strip(),
                "profiles": self.profiles,
                "date_from": self.date_from.get().strip(),
                "date_to": self.date_to.get().strip()
            }, ensure_ascii=False, indent=2), "utf-8")
        except Exception:
            pass

    def creds(self):
        if not self.company_name.get().strip():
            raise base.ApiError("Inserisci il nome dell'azienda.")
        return super().creds()

    def _set_real_company(self, company_id, real_name):
        self.actual_company_id = company_id
        self.actual_company_name = real_name
        self.after(0, lambda: self.verified_label.set(
            f"Azienda realmente collegata su FIC: {real_name} | ID {company_id}"
        ))

    def verify_connection(self):
        def job():
            self.set_busy(True)
            try:
                c, t = self.creds()
                _, js = base.api_request(
                    "GET", f"{base.API}/c/{urllib.parse.quote(c)}/company/info", t)
                name = str((js.get("data") or {}).get("name") or "azienda")
                self._set_real_company(c, name)
                self.set_status(f"Collegamento verificato: {name} | ID {c}")
            except Exception as e:
                self.set_status("Errore: " + str(e))
            finally:
                self.set_busy(False)
        self.run_bg(job)

    def scan_candidates(self, company, token):
        scanned, ready = super().scan_candidates(company, token)
        self.candidates_company_id = company
        return scanned, ready

    def check_and_send(self):
        def job():
            self.set_busy(True)
            try:
                c, t = self.creds()
                _, js = base.api_request(
                    "GET", f"{base.API}/c/{urllib.parse.quote(c)}/company/info", t)
                name = str((js.get("data") or {}).get("name") or "azienda")
                self._set_real_company(c, name)
                scanned, ready = self.scan_candidates(c, t)
                self.candidates = ready
                self.after(0, self.render)
                self.after(0, lambda: self.count_checked.set(str(scanned)))
                self.after(0, lambda: self.count_sent.set("0"))
                self.after(0, lambda: self.count_errors.set("0"))
                if ready:
                    self.set_status(f"{len(ready)} fatture da spedire per {name}.")
                    self.after(0, lambda: self.process_all(dry_run=False))
                else:
                    self.set_status(f"Nessuna fattura da spedire per {name}.")
            except Exception as e:
                self.set_status("Errore: " + str(e))
            finally:
                self.set_busy(False)
        self.run_bg(job)

    def process_all(self, dry_run=False):
        if not self.candidates:
            messagebox.showinfo(base.APP, "Prima carica le fatture da spedire.")
            return
        if self.candidates_company_id != self.company.get().strip():
            messagebox.showerror(
                base.APP,
                "L'ID azienda e cambiato dopo il controllo delle fatture. "
                "Ripeti il controllo prima di procedere.")
            return

        if not dry_run:
            alias = self.company_name.get().strip()
            cid = self.company.get().strip()
            real = (self.actual_company_name
                    if self.actual_company_id == cid else "(non verificata)")
            if not messagebox.askyesno(
                base.APP,
                f"Controlla l'azienda destinataria dell'invio:\n\n"
                f"Nome inserito: {alias}\n"
                f"ID azienda: {cid}\n"
                f"Azienda registrata su FIC: {real}\n\n"
                f"Fatture elettroniche da spedire: {len(self.candidates)}\n\n"
                "Confermi questa azienda?",
                icon="warning",
            ):
                return
        # La procedura originale esegue verifica XML, seconda conferma
        # e invio effettivo; nessun invio avviene durante le verifiche.
        super().process_all(dry_run=dry_run)


if __name__ == "__main__":
    App().mainloop()
