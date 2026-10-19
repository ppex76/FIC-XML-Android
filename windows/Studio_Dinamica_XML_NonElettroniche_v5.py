import re
import tkinter as tk
from tkinter import ttk

import Studio_Dinamica_XML_NonElettroniche as base
import Studio_Dinamica_XML_NonElettroniche_v4 as v4


class App(v4.App):
    def _build_ui(self):
        # Impostazioni di ordinamento iniziali.
        self._invoice_sort_col = "data"
        self._invoice_sort_reverse = False
        self._summary_sort_col = "mese"
        self._summary_sort_reverse = False

        super()._build_ui()
        self.title("Studio Dinamica — XML + Excel + tabelle ordinabili")

        # Riepilogo: aggiunge anche l'anno e rende cliccabili tutte le intestazioni.
        self._summary_labels = {
            "anno": "Anno",
            "mese": "Mese",
            "fatture": "N. fatture",
            "totale": "Totale fatture €",
            "bolli_n": "Marche da 2 €",
            "bolli_tot": "Totale bolli €",
        }
        summary_cols = tuple(self._summary_labels.keys())
        self.summary_tree.configure(columns=summary_cols)
        summary_widths = {
            "anno": (80, "center"),
            "mese": (150, "w"),
            "fatture": (100, "e"),
            "totale": (155, "e"),
            "bolli_n": (125, "e"),
            "bolli_tot": (145, "e"),
        }
        for col in summary_cols:
            width, anchor = summary_widths[col]
            self.summary_tree.column(col, width=width, anchor=anchor)
            self.summary_tree.heading(
                col,
                text=self._summary_labels[col],
                command=lambda c=col: self.sort_summary_by(c),
            )

        # Dettaglio fatture: anno + mese + data + numero e tutte le altre colonne ordinabili.
        self._invoice_labels = {
            "anno": "Anno",
            "mese": "Mese",
            "data": "Data",
            "numero": "N. fattura",
            "cliente": "Cliente",
            "totale": "Totale €",
            "bollo": "Bollo €",
            "xml": "XML / stato",
        }
        invoice_cols = tuple(self._invoice_labels.keys())
        self.invoice_tree.configure(columns=invoice_cols)
        invoice_widths = {
            "anno": (72, "center"),
            "mese": (110, "w"),
            "data": (105, "center"),
            "numero": (110, "w"),
            "cliente": (280, "w"),
            "totale": (105, "e"),
            "bollo": (85, "e"),
            "xml": (330, "w"),
        }
        for col in invoice_cols:
            width, anchor = invoice_widths[col]
            self.invoice_tree.column(col, width=width, anchor=anchor)
            self.invoice_tree.heading(
                col,
                text=self._invoice_labels[col],
                command=lambda c=col: self.sort_invoices_by(c),
            )

        # Scroll orizzontale utile con tutte le nuove colonne.
        try:
            parent = self.invoice_tree.master
            xbar = ttk.Scrollbar(parent, orient="horizontal", command=self.invoice_tree.xview)
            self.invoice_tree.configure(xscrollcommand=xbar.set)
            xbar.grid(row=1, column=0, sticky="ew")
        except Exception:
            pass

        hint = ttk.Label(
            self,
            text=(
                "ORDINAMENTO: clicca sull'intestazione di una colonna (Anno, Mese, Data, N. fattura, "
                "Cliente, Totale, Bollo, ecc.). Un secondo clic inverte crescente/decrescente."
            ),
            foreground="#475467",
            font=("Segoe UI", 9, "bold"),
        )
        hint.grid(row=6, column=0, sticky="w", padx=14, pady=(0, 10))

        self._update_summary_headers()
        self._update_invoice_headers()

    @staticmethod
    def _money_key(value):
        s = str(value or "").replace("€", "").replace(" ", "").strip()
        if not s:
            return 0.0
        # Formato italiano: 1.234,56
        try:
            return float(s.replace(".", "").replace(",", "."))
        except Exception:
            return 0.0

    @staticmethod
    def _natural_key(value):
        # 2, 10, 10/A, 11 vengono ordinati in modo naturale e non alfabetico puro.
        parts = re.split(r"(\d+)", str(value or "").strip().lower())
        return tuple(int(p) if p.isdigit() else p for p in parts)

    @staticmethod
    def _month_key(value):
        s = str(value or "").strip()
        try:
            return base.MONTHS.index(s) + 1
        except ValueError:
            try:
                return int(s)
            except Exception:
                return 0

    def _tree_value_key(self, col, value):
        if col in {"anno", "fatture", "bolli_n"}:
            try:
                return int(str(value or "0").strip())
            except Exception:
                return 0
        if col == "mese":
            return self._month_key(value)
        if col in {"totale", "bolli_tot", "bollo"}:
            return self._money_key(value)
        if col == "numero":
            return self._natural_key(value)
        if col == "data":
            # Le date FIC sono YYYY-MM-DD: l'ordinamento stringa coincide con quello cronologico.
            return str(value or "")
        return str(value or "").casefold()

    def _sort_tree(self, tree, col, reverse):
        rows = list(tree.get_children(""))
        rows.sort(
            key=lambda iid: self._tree_value_key(col, tree.set(iid, col)),
            reverse=reverse,
        )
        for pos, iid in enumerate(rows):
            tree.move(iid, "", pos)

    def _update_summary_headers(self):
        for col, label in self._summary_labels.items():
            arrow = ""
            if col == self._summary_sort_col:
                arrow = " ▼" if self._summary_sort_reverse else " ▲"
            self.summary_tree.heading(
                col,
                text=label + arrow,
                command=lambda c=col: self.sort_summary_by(c),
            )

    def _update_invoice_headers(self):
        for col, label in self._invoice_labels.items():
            arrow = ""
            if col == self._invoice_sort_col:
                arrow = " ▼" if self._invoice_sort_reverse else " ▲"
            self.invoice_tree.heading(
                col,
                text=label + arrow,
                command=lambda c=col: self.sort_invoices_by(c),
            )

    def sort_summary_by(self, col):
        if col == self._summary_sort_col:
            self._summary_sort_reverse = not self._summary_sort_reverse
        else:
            self._summary_sort_col = col
            self._summary_sort_reverse = False
        self._sort_tree(self.summary_tree, self._summary_sort_col, self._summary_sort_reverse)
        self._update_summary_headers()

    def sort_invoices_by(self, col):
        if col == self._invoice_sort_col:
            self._invoice_sort_reverse = not self._invoice_sort_reverse
        else:
            self._invoice_sort_col = col
            self._invoice_sort_reverse = False
        self._sort_tree(self.invoice_tree, self._invoice_sort_col, self._invoice_sort_reverse)
        self._update_invoice_headers()

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
                    base.fmt_money(d.get("amount_gross")),
                    base.fmt_money(d.get("_stamp_total")),
                    str(d.get("_xml_status") or ""),
                ),
            )

        self._sort_tree(self.invoice_tree, self._invoice_sort_col, self._invoice_sort_reverse)
        self._update_invoice_headers()


if __name__ == "__main__":
    App().mainloop()
