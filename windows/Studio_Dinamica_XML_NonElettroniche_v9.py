import Studio_Dinamica_XML_NonElettroniche as base
import Studio_Dinamica_XML_NonElettroniche_v8 as v8


class App(v8.App):
    def _build_ui(self):
        super()._build_ui()
        self.title("Studio Dinamica — SOLO fatture non elettroniche + imponibile + Excel v9")
        self.status.set(
            "Pronto. V9: carico tutte le fatture dell'anno, poi ESCLUDO SEMPRE quelle elettroniche. "
            "Pagamento e Sistema TS non vengono usati come filtro."
        )

    @staticmethod
    def _is_electronic(d):
        """Riconosce esclusivamente il flag e_invoice della fattura.

        Non usa ei_status, pagamento, incasso o Sistema TS.
        """
        v = d.get("e_invoice")
        if v is True:
            return True
        if isinstance(v, (int, float)):
            return v == 1
        if isinstance(v, str):
            return v.strip().lower() in {"1", "true", "yes", "si", "sì"}
        return False

    def load_year(self):
        """Carica tutte le fatture e mantiene SOLO quelle non elettroniche.

        Regola unica di esclusione: e_invoice == true.
        NON vengono considerati come filtro:
        - pagamento / non pagamento
        - incasso
        - Sistema Tessera Sanitaria inviato / non inviato
        - ei_status
        """
        def job():
            try:
                c, t = self.creds()
                start, end = self.year_bounds()
                all_rows = self._list_all_invoices(c, t, start, end)

                included = []
                excluded_electronic = []

                for d in all_rows:
                    if self._is_electronic(d):
                        excluded_electronic.append(d)
                        continue

                    d["_stamp_total"], d["_stamp_count"] = base.stamp_info(d)
                    d["_xml_status"] = "Da esportare"
                    included.append(d)

                included.sort(
                    key=lambda x: (
                        str(x.get("date") or ""),
                        base.invoice_number(x),
                        str(x.get("id") or ""),
                    )
                )

                self.all_invoices = included
                self.after(0, self.render_all)

                jan = [d for d in included if str(d.get("date") or "")[5:7] == "01"]
                jan_net = sum(base.to_float(d.get("amount_net"), 0.0) for d in jan)
                jan_gross = sum(base.to_float(d.get("amount_gross"), 0.0) for d in jan)

                jan_e = [d for d in excluded_electronic if str(d.get("date") or "")[5:7] == "01"]
                jan_e_net = sum(base.to_float(d.get("amount_net"), 0.0) for d in jan_e)

                self.set_status(
                    f"Caricate {len(included)} fatture NON elettroniche. "
                    f"Escluse {len(excluded_electronic)} fatture elettroniche. "
                    f"GENNAIO NON ELETTRONICHE: {len(jan)} fatture | imponibile € {base.fmt_money(jan_net)} | "
                    f"totale documenti € {base.fmt_money(jan_gross)}. "
                    f"Elettroniche escluse a gennaio: {len(jan_e)} | imponibile escluso € {base.fmt_money(jan_e_net)}."
                )
            except Exception as e:
                self.set_status("Errore: " + str(e))

        self.run_bg(job)

    def update_annual_line(self):
        net = sum(base.to_float(d.get("amount_net"), 0.0) for d in self.all_invoices)
        gross = sum(base.to_float(d.get("amount_gross"), 0.0) for d in self.all_invoices)
        stamp_total = sum(base.to_float(d.get("_stamp_total"), 0.0) for d in self.all_invoices)
        stamp_count = sum(int(d.get("_stamp_count") or 0) for d in self.all_invoices)
        self.annual.set(
            f"ANNO {self.year.get()} — SOLO fatture NON elettroniche: {len(self.all_invoices)} | "
            f"Imponibile: € {base.fmt_money(net)} | Totale documenti: € {base.fmt_money(gross)} | "
            f"Marche da 2 €: {stamp_count} | Totale bolli: € {base.fmt_money(stamp_total)}"
        )


if __name__ == "__main__":
    App().mainloop()
