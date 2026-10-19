import json
import urllib.parse

import Studio_Dinamica_XML_NonElettroniche as base
import Studio_Dinamica_XML_NonElettroniche_v7 as v7


class App(v7.App):
    def _build_ui(self):
        super()._build_ui()
        self.title("Studio Dinamica — tutte le fatture + imponibile + Excel v8")
        self.geometry("1450x930")
        self.minsize(1150, 780)
        self.status.set(
            "Pronto. V8: nessuna fattura viene scartata per e_invoice, pagamento o Sistema TS. "
            "Il riepilogo mostra anche l'IMPONIBILE (amount_net)."
        )

        # Riepilogo: aggiungiamo l'imponibile senza togliere il totale documento.
        self._summary_labels = {
            "anno": "Anno",
            "mese": "Mese",
            "fatture": "N. fatture",
            "imponibile": "Imponibile €",
            "totale": "Totale documenti €",
            "bolli_n": "Marche da 2 €",
            "bolli_tot": "Totale bolli €",
        }
        summary_cols = tuple(self._summary_labels.keys())
        self.summary_tree.configure(columns=summary_cols)
        summary_widths = {
            "anno": (75, "center"),
            "mese": (135, "w"),
            "fatture": (95, "e"),
            "imponibile": (145, "e"),
            "totale": (155, "e"),
            "bolli_n": (120, "e"),
            "bolli_tot": (135, "e"),
        }
        for col in summary_cols:
            width, anchor = summary_widths[col]
            self.summary_tree.column(col, width=width, anchor=anchor)
            self.summary_tree.heading(
                col,
                text=self._summary_labels[col],
                command=lambda c=col: self.sort_summary_by(c),
            )

        # Dettaglio: imponibile + flag elettronica + stato SDI, solo come informazione.
        self._invoice_labels = {
            "anno": "Anno",
            "mese": "Mese",
            "data": "Data",
            "numero": "N. fattura",
            "cliente": "Cliente",
            "imponibile": "Imponibile €",
            "totale": "Totale €",
            "bollo": "Bollo €",
            "elettronica": "e_invoice",
            "sdi": "Stato SDI",
            "xml": "XML / stato",
        }
        invoice_cols = tuple(self._invoice_labels.keys())
        self.invoice_tree.configure(columns=invoice_cols)
        invoice_widths = {
            "anno": (68, "center"),
            "mese": (100, "w"),
            "data": (100, "center"),
            "numero": (100, "w"),
            "cliente": (250, "w"),
            "imponibile": (110, "e"),
            "totale": (105, "e"),
            "bollo": (80, "e"),
            "elettronica": (85, "center"),
            "sdi": (100, "w"),
            "xml": (280, "w"),
        }
        for col in invoice_cols:
            width, anchor = invoice_widths[col]
            self.invoice_tree.column(col, width=width, anchor=anchor)
            self.invoice_tree.heading(
                col,
                text=self._invoice_labels[col],
                command=lambda c=col: self.sort_invoices_by(c),
            )

        self._update_summary_headers()
        self._update_invoice_headers()

    def _tree_value_key(self, col, value):
        if col == "imponibile":
            return self._money_key(value)
        if col == "elettronica":
            return str(value or "").casefold()
        if col == "sdi":
            return str(value or "").casefold()
        return super()._tree_value_key(col, value)

    def _list_all_invoices(self, company, token, start, end):
        """Segue next_page_url dell'API e NON applica filtri locali su e_invoice/TS/pagamenti."""
        q = f"date >= '{start}' and date <= '{end}'"
        params = {
            "type": "invoice",
            "q": q,
            "per_page": "100",
            "fieldset": "detailed",
            "sort": "date",
        }
        url = (
            f"{base.API}/c/{urllib.parse.quote(str(company))}/issued_documents?"
            f"{urllib.parse.urlencode(params)}"
        )
        rows = []
        seen = set()
        page_no = 0

        while url:
            page_no += 1
            self.set_status(
                f"Carico TUTTE le fatture {self.year.get()} — pagina {page_no}… "
                "nessun filtro e_invoice / pagamento / Sistema TS"
            )
            _, _, body = base.api_get(url, token)
            js = json.loads(body.decode("utf-8"))
            for d in js.get("data") or []:
                did = str(d.get("id") or "")
                if did and did in seen:
                    continue
                if did:
                    seen.add(did)
                rows.append(d)
            url = js.get("next_page_url") or None

        return rows

    def load_year(self):
        """V8: prende TUTTE le fatture del periodo e non ne scarta nessuna.

        Non filtra:
        - e_invoice
        - ei_status
        - Sistema Tessera Sanitaria
        - pagata/non pagata
        - incassata/non incassata

        Il solo filtro API e' type=invoice, cioe' chiediamo all'API le fatture e non altri documenti.
        """
        def job():
            try:
                c, t = self.creds()
                start, end = self.year_bounds()
                rows = self._list_all_invoices(c, t, start, end)

                for d in rows:
                    d["_stamp_total"], d["_stamp_count"] = base.stamp_info(d)
                    d["_xml_status"] = "Da esportare"

                rows.sort(
                    key=lambda x: (
                        str(x.get("date") or ""),
                        base.invoice_number(x),
                        str(x.get("id") or ""),
                    )
                )
                self.all_invoices = rows
                self.after(0, self.render_all)

                jan = [d for d in rows if str(d.get("date") or "")[5:7] == "01"]
                jan_net = sum(base.to_float(d.get("amount_net"), 0.0) for d in jan)
                jan_gross = sum(base.to_float(d.get("amount_gross"), 0.0) for d in jan)
                jan_e_no = sum(
                    base.to_float(d.get("amount_net"), 0.0)
                    for d in jan if not base.boolish(d.get("e_invoice"))
                )
                jan_e_yes = sum(
                    base.to_float(d.get("amount_net"), 0.0)
                    for d in jan if base.boolish(d.get("e_invoice"))
                )

                self.set_status(
                    f"Caricate {len(rows)} fatture SENZA ESCLUSIONI. "
                    f"GENNAIO: {len(jan)} fatture | imponibile totale € {base.fmt_money(jan_net)} | "
                    f"totale documenti € {base.fmt_money(jan_gross)} | "
                    f"imponibile e_invoice=NO € {base.fmt_money(jan_e_no)} | "
                    f"e_invoice=SI € {base.fmt_money(jan_e_yes)}."
                )
            except Exception as e:
                self.set_status("Errore: " + str(e))

        self.run_bg(job)

    def build_summary(self):
        rows = {
            m: {
                "count": 0,
                "net": 0.0,
                "total": 0.0,
                "stamp_count": 0,
                "stamp_total": 0.0,
            }
            for m in range(1, 13)
        }
        for d in self.all_invoices:
            try:
                m = int(str(d.get("date") or "")[5:7])
            except Exception:
                continue
            if m not in rows:
                continue
            r = rows[m]
            r["count"] += 1
            r["net"] += base.to_float(d.get("amount_net"), 0.0)
            r["total"] += base.to_float(d.get("amount_gross"), 0.0)
            r["stamp_total"] += base.to_float(d.get("_stamp_total"), 0.0)
            r["stamp_count"] += int(d.get("_stamp_count") or 0)
        self.summary_rows = rows

    def render_summary(self):
        for iid in self.summary_tree.get_children():
            self.summary_tree.delete(iid)
        year = self.year.get()
        for m in range(1, 13):
            r = self.summary_rows[m]
            self.summary_tree.insert(
                "",
                "end",
                iid=str(m),
                values=(
                    year,
                    base.MONTHS[m - 1],
                    r["count"],
                    base.fmt_money(r["net"]),
                    base.fmt_money(r["total"]),
                    r["stamp_count"],
                    base.fmt_money(r["stamp_total"]),
                ),
            )
        self._sort_tree(self.summary_tree, self._summary_sort_col, self._summary_sort_reverse)
        self._update_summary_headers()

    def render_invoices(self):
        for iid in self.invoice_tree.get_children():
            self.invoice_tree.delete(iid)

        if self.selected_month:
            key = f"{self.selected_month:02d}"
            rows = [d for d in self.all_invoices if str(d.get("date") or "")[5:7] == key]
        else:
            rows = list(self.all_invoices)

        self.visible_invoices = rows
        for i, d in enumerate(rows):
            date_s = str(d.get("date") or "")
            try:
                year_s = date_s[:4]
                month_n = int(date_s[5:7])
                month_s = base.MONTHS[month_n - 1] if 1 <= month_n <= 12 else ""
            except Exception:
                year_s = ""
                month_s = ""
            ent = d.get("entity") or {}
            self.invoice_tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    year_s,
                    month_s,
                    date_s,
                    base.invoice_number(d),
                    str(ent.get("name") or ""),
                    base.fmt_money(d.get("amount_net")),
                    base.fmt_money(d.get("amount_gross")),
                    base.fmt_money(d.get("_stamp_total")),
                    "SI" if base.boolish(d.get("e_invoice")) else "NO",
                    str(d.get("ei_status") or ""),
                    str(d.get("_xml_status") or ""),
                ),
            )

        self._sort_tree(self.invoice_tree, self._invoice_sort_col, self._invoice_sort_reverse)
        self._update_invoice_headers()

    def update_annual_line(self):
        net = sum(base.to_float(d.get("amount_net"), 0.0) for d in self.all_invoices)
        gross = sum(base.to_float(d.get("amount_gross"), 0.0) for d in self.all_invoices)
        stamp_total = sum(base.to_float(d.get("_stamp_total"), 0.0) for d in self.all_invoices)
        stamp_count = sum(int(d.get("_stamp_count") or 0) for d in self.all_invoices)
        self.annual.set(
            f"ANNO {self.year.get()} — Fatture: {len(self.all_invoices)} | "
            f"Imponibile: € {base.fmt_money(net)} | Totale documenti: € {base.fmt_money(gross)} | "
            f"Marche da 2 €: {stamp_count} | Totale bolli: € {base.fmt_money(stamp_total)}"
        )


if __name__ == "__main__":
    App().mainloop()
