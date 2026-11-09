"""V8 originale con il solo nome azienda modificabile (ID gia' editabile)."""
import json
import tkinter as tk
from tkinter import ttk
import urllib.parse

import Studio_Dinamica_XML_NonElettroniche as base
import Studio_Dinamica_XML_NonElettroniche_v8 as v8


class App(v8.App):
    def _build_ui(self):
        super()._build_ui()
        self.nome_azienda_input = tk.StringVar(value="")
        self.intestazione_azienda = None
        for frame in self.winfo_children():
            if isinstance(frame, ttk.Frame) and frame.grid_info().get("row") == 0:
                for widget in frame.winfo_children():
                    if isinstance(widget, ttk.Label) and "Studio Dinamica" in str(widget.cget("text")):
                        self.intestazione_azienda = widget
                        break
                break
        top = self.token_entry.master
        ttk.Label(top, text="Nome azienda").grid(
            row=4, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(top, textvariable=self.nome_azienda_input, width=38).grid(
            row=4, column=1, columnspan=2, sticky="ew", padx=(8, 18), pady=(10, 0))
        ttk.Label(top, text="Modificabile. ID azienda: campo in alto.", foreground="#667085").grid(
            row=4, column=3, columnspan=2, sticky="w", pady=(10, 0))
        self.nome_azienda_input.trace_add("write", self._aggiorna_nome)
        self._aggiorna_nome()
        # Il riepilogo occupa meno spazio; l'elenco fatture viene privilegiato.
        self.summary_tree.configure(height=4)
        self.invoice_tree.configure(height=16)
        self.after(100, self._massimizza_finestra)
        self.after(650, self._imposta_altezze_tabelle)

    def _massimizza_finestra(self):
        try:
            self.state('zoomed')
        except tk.TclError:
            try:
                self.attributes('-zoomed', True)
            except tk.TclError:
                pass

    def _imposta_altezze_tabelle(self):
        try:
            panes = next(w for w in self.grid_slaves(row=2, column=0) if isinstance(w, ttk.Panedwindow))
            height = panes.winfo_height()
            if height > 260:
                # 30% riepilogo; 70% elenco fatture. La barra resta trascinabile.
                panes.sashpos(0, min(max(110, int(height * 0.30)), height - 150))
        except (StopIteration, tk.TclError):
            pass

    def _aggiorna_nome(self, *args):
        nome = self.nome_azienda_input.get().strip()
        self.company_name = nome
        self.title(f"{nome or 'Fatture in Cloud'} — XML + Excel v8")
        if self.intestazione_azienda is not None:
            self.intestazione_azienda.configure(
                text=f"{nome or 'Fatture in Cloud'} — Fatture non elettroniche"
            )

    def _load_cfg(self):
        super()._load_cfg()
        try:
            data = json.loads(self.cfg_path.read_text("utf-8"))
            nome = data.get("nome_azienda_editabile")
            if nome is not None:
                self.nome_azienda_input.set(str(nome))
        except Exception:
            pass

    def _save_cfg(self):
        # Mantiene esattamente i campi originali (mese, trimestre, cartella, ID).
        super()._save_cfg()
        try:
            data = json.loads(self.cfg_path.read_text("utf-8"))
            data["nome_azienda_editabile"] = self.nome_azienda_input.get().strip()
            self.cfg_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
        except Exception:
            pass

    def verify(self):
        # Stessa verifica dell'originale, senza l'avviso specifico su 'Dinamica'.
        # Il nome inserito dall'utente resta editabile e non viene sovrascritto.
        def job():
            try:
                c, t = self.creds()
                _, _, body = base.api_get(
                    f"{base.API}/c/{urllib.parse.quote(c)}/company/info", t)
                data = json.loads(body.decode("utf-8"))
                nome_fic = str((data.get("data") or {}).get("name") or "azienda")
                self.set_status(f"Collegamento riuscito: {nome_fic} (ID {c})")
            except Exception as e:
                self.set_status("Errore: " + str(e))
        self.run_bg(job)


if __name__ == "__main__":
    App().mainloop()
