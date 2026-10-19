import Studio_Dinamica_XML_NonElettroniche as base
import Studio_Dinamica_XML_NonElettroniche_v5 as v5


class App(v5.App):
    def _build_ui(self):
        super()._build_ui()
        self.title("Studio Dinamica — XML non elettroniche + Excel + ordinamento")
        self.status.set(
            "Pronto. Le fatture non elettroniche vengono esportate direttamente come XML dati, senza chiamare FatturaPA/SDI."
        )

    def export_one(self, d, c, t):
        """Esporta SEMPRE una fattura non elettronica come XML dati.

        Non usa l'endpoint e_invoice/xml, perché quello è per FatturaPA/e-fatture
        e sulle fatture non elettroniche può rispondere 422.
        """
        path = self.unique_path(d)
        body = base.make_custom_xml(d)
        path.write_bytes(body)
        kind = "XML dati fattura non elettronica"
        d["_xml_status"] = f"✓ {kind}: {path.name}"
        return path, kind

    def export_batch(self, docs, label):
        if not docs:
            from tkinter import messagebox
            messagebox.showinfo(base.APP_TITLE, "Nessuna fattura da esportare.")
            return

        def job():
            ok = errors = 0
            try:
                # Chiediamo comunque le credenziali per coerenza con l'app e per eventuali operazioni successive,
                # ma l'XML viene creato dai dati già caricati e non richiama FatturaPA.
                c, t = self.creds()
                for i, d in enumerate(docs, start=1):
                    self.set_status(f"{label}: {i}/{len(docs)} — fattura {base.invoice_number(d)}")
                    try:
                        self.export_one(d, c, t)
                        ok += 1
                    except Exception as e:
                        errors += 1
                        d["_xml_status"] = "✕ " + str(e)
                    self.after(0, self.sync_visible_status)
                self.write_summary_csv(self.output_root() / f"Riepilogo_{self.year.get()}.csv")
                self.set_status(
                    f"Completato: {ok} XML dati di fatture non elettroniche salvati, {errors} errori. "
                    "Nessuna chiamata FatturaPA/SDI: il 422 non viene più generato."
                )
            except Exception as e:
                self.set_status("Errore: " + str(e))

        self.run_bg(job)


if __name__ == "__main__":
    App().mainloop()
