import json
import urllib.error
import urllib.parse
import urllib.request

import Studio_Dinamica_XML_NonElettroniche as base
import Studio_Dinamica_XML_NonElettroniche_v6 as v6


class App(v6.App):
    def _build_ui(self):
        super()._build_ui()
        self.title("Studio Dinamica — fatture non elettroniche + Excel + ordinamento v7")
        self.status.set(
            "Pronto. Carico TUTTE le fatture (anche non pagate e TS non inviate), poi escludo solo quelle elettroniche."
        )

    def _list_invoices_page(self, company, token, start, end, page):
        """Lista fatture senza usare endpoint FatturaPA.

        Il parametro type=invoice è richiesto/atteso dal metodo List Issued Documents.
        Nessun filtro su pagamento, incasso o Sistema TS.
        In caso di 422 mostriamo il vero messaggio API invece del vecchio testo FatturaPA.
        """
        q = f"date >= '{start}' and date <= '{end}'"
        params = {
            "type": "invoice",
            "q": q,
            "page": str(page),
            "per_page": "100",
            "fieldset": "detailed",
        }
        url = (
            f"{base.API}/c/{urllib.parse.quote(str(company))}/issued_documents?"
            f"{urllib.parse.urlencode(params)}"
        )
        req = urllib.request.Request(
            url,
            headers={
                "Authorization": "Bearer " + base.clean_token(token),
                "Accept": "application/json",
                "User-Agent": "Studio-Dinamica-NonElettroniche-v7",
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=90) as r:
                body = r.read().decode("utf-8", "replace")
                return json.loads(body)
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            detail = body[:1200]
            try:
                js = json.loads(body)
                err = js.get("error") or {}
                msg = str(err.get("message") or "").strip()
                validation = err.get("validation_result")
                if msg:
                    detail = msg
                if validation:
                    detail += " — " + json.dumps(validation, ensure_ascii=False)
            except Exception:
                pass
            if e.code == 401:
                raise base.ApiError("401 — token non valido, scaduto o revocato")
            if e.code == 403:
                raise base.ApiError("403 — permessi insufficienti sul token per leggere le fatture")
            if e.code == 422:
                raise base.ApiError("422 — richiesta elenco fatture non valida: " + detail)
            raise base.ApiError(f"HTTP {e.code} — {detail}")
        except urllib.error.URLError as e:
            raise base.ApiError("Errore di rete: " + str(e))

    def load_year(self):
        """Carica tutte le fatture dell'anno, indipendentemente da pagamento e TS.

        Inclusione:
        - type=invoice (serve solo a chiedere all'API il tipo corretto di documenti)
        - poi lato client: e_invoice NON deve essere true

        NON filtriamo mai paid, payment_status, TS, system_ts, incasso o invio TS.
        """
        def job():
            try:
                c, t = self.creds()
                start, end = self.year_bounds()
                page = 1
                last_page = 1
                rows = []

                while True:
                    self.set_status(
                        f"Carico TUTTE le fatture {self.year.get()} — pagina {page}… "
                        "(nessun filtro pagamento/TS)"
                    )
                    js = self._list_invoices_page(c, t, start, end, page)
                    data = js.get("data") or []
                    if page == 1:
                        last_page = int(js.get("last_page") or 1)
                    rows.extend(data)
                    if page >= last_page:
                        break
                    page += 1

                included = []
                skipped_electronic = 0
                seen = set()

                for d in rows:
                    did = str(d.get("id") or "")
                    if did and did in seen:
                        continue
                    if did:
                        seen.add(did)

                    # L'API ha già restituito solo type=invoice.
                    # Qui escludiamo esclusivamente le fatture elettroniche.
                    if base.boolish(d.get("e_invoice")):
                        skipped_electronic += 1
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
                self.set_status(
                    f"Caricate {len(included)} fatture NON elettroniche {self.year.get()}. "
                    f"Incluse NON PAGATE e TS NON INVIATE. Escluse solo {skipped_electronic} fatture elettroniche."
                )
            except Exception as e:
                self.set_status("Errore: " + str(e))

        self.run_bg(job)


if __name__ == "__main__":
    App().mainloop()
