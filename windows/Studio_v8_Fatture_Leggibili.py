"""Solo correzione dell'area fatture: 30% riepilogo, 70% elenco, splitter trascinabile."""
import tkinter as tk
from tkinter import ttk

import Studio_Dinamica_XML_v8_Azienda_Editabile as precedente


class App(precedente.App):
    def __init__(self):
        super().__init__()
        self._splitter_fatture = next(
            (widget for widget in self.winfo_children()
             if isinstance(widget, ttk.Panedwindow)),
            None,
        )
        if self._splitter_fatture is not None:
            pannelli = self._splitter_fatture.panes()
            if len(pannelli) >= 2:
                # Distribuisce lo spazio disponibile, anche dopo un ridimensionamento.
                self._splitter_fatture.pane(pannelli[0], weight=3)
                self._splitter_fatture.pane(pannelli[1], weight=7)

        # L'interfaccia di v8, tutti i filtri e le esportazioni sono invariati.
        try:
            self.state("zoomed")  # Massimizzata su Windows.
        except tk.TclError:
            self.geometry("1450x930")

        self.after(350, self._imposta_altezze)

    def _imposta_altezze(self):
        # Esegue dopo che Windows ha assegnato l'altezza effettiva ai pannelli.
        pane = self._splitter_fatture
        if pane is None or not pane.winfo_exists():
            return
        height = pane.winfo_height()
        if height > 130 and len(pane.panes()) >= 2:
            try:
                pane.sashpos(0, round(height * 0.30))
            except tk.TclError:
                pass  # Il divisorio resta comunque trascinabile.


if __name__ == "__main__":
    App().mainloop()
