# FIC XML Mese — Android

App Android per esportare gli XML da Fatture in Cloud una fattura alla volta.

## Build automatico con GitHub Actions

Ogni push sul branch `main` avvia automaticamente la compilazione.

Per scaricare l'APK:

1. Apri il repository su GitHub.
2. Vai su **Actions**.
3. Apri l'esecuzione **Build Android APK** con spunta verde.
4. Scorri fino a **Artifacts**.
5. Scarica **FIC-XML-Mese-APK**.
6. Estrai lo ZIP: dentro trovi `FIC-XML-Mese.apk`.
7. Copia/installalo sul telefono Android.

L'APK generato è una build debug, sufficiente per provare l'app sul proprio telefono.

## Funzionamento

- ID azienda
- token personale Fatture in Cloud
- selezione mese/anno
- elenco fatture del mese
- pulsante `Esporta XML` per ogni fattura
- `Esporta tutte — una alla volta` esegue le chiamate in sequenza
- salvataggio in `Download/FattureInCloud/AAAA-MM`

L'app prova sull'API pubblica:
`/e_invoice/xml?avoid_einvoice_check=1&include_attachment=false`

Se una singola fattura non è esportabile tramite API pubblica, viene segnalata e l'app continua con la successiva.
