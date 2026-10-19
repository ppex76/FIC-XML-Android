import json
import urllib.parse

import Studio_Dinamica_XML_NonElettroniche as base
import Studio_Dinamica_XML_NonElettroniche_v3 as v3


class App(v3.App):
    def _build_ui(self):
        super()._build_ui()
        self.title("Studio Dinamica — tutte le fatture non elettroniche + Excel")
        self.status.set(
            "Pronto. Verranno incluse TUTTE le fatture non elettroniche, indipendentemente da pagamento e Sistema TS."
        )

    def load_year(self):
        """Carica tutti i documenti emessi dell'anno e filtra SOLO lato client.

        Regola di inclusione:
        - type == invoice
        - e_invoice != true

        NON vengono mai usati come filtro:
        - stato pagamento / paid / payment_status
        - stato Sistema Tessera Sanitaria / TS
        - stato incasso
        - stato invio TS
        """
        def job():
            try:
                c, t = self.creds()
                start, end = self.year_bounds()
                page, last_page, rows = 1, 1, []

                while True:
                    q = f"date >= '{start}' and date <= '{end}'"
                    # Niente filtro API su pagamento, TS, stato, incasso o tipo.
                    # Recuperiamo tutti i documenti emessi dell'anno e decidiamo noi localmente.
                    params = {
                        "q": q,
                        "page": str(page),
                        "per_page": "100",
                        "fieldset": "detailed",
                    }
                    url = (
                        f"{base.API}/c/{urllib.parse.quote(c)}/issued_documents?"
                        f"{urllib.parse.urlencode(params)}"
                    )
                    self.set_status(
                        f"Carico TUTTI i documenti emessi {self.year.get()} — pagina {page}…"
                    )
                    _, _, body = base.api_get(url, t)
                    js = json.loads(body.decode("utf-8"))
                    data = js.get("data") or []
                    if page == 1:
                        last_page = int(js.get("last_page") or 1)
                    rows.extend(data)
                    if page >= last_page:
                        break
                    page += 1

                included = []
                skipped_electronic = 0
                skipped_other_type = 0
                seen = set()

                for d in rows:
                    did = str(d.get("id") or "")
                    if did and did in seen:
                        continue
                    if did:
                        seen.add(did)

                    # Unici due criteri: deve essere una fattura e NON deve essere elettronica.
                    if str(d.get("type") or "").lower() != "invoice":
                        skipped_other_type += 1
                        continue
                    if base.boolish(d.get("e_invoice")):
                        skipped_electronic += 1
                        continue

                    # Nessun controllo su pagata/non pagata, TS inviato/non inviato, incassata/non incassata.
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
                self.set_status(
                    f"Caricate {len(included)} fatture NON elettroniche {self.year.get()}. "
                    f"Incluse anche NON PAGATE e TS NON INVIATE. "
                    f"Escluse solo {skipped_electronic} elettroniche e {skipped_other_type} documenti non-fattura."
                )
            except Exception as e:
                self.set_status("Errore: " + str(e))

        self.run_bg(job)


if __name__ == "__main__":
    App().mainloop()
