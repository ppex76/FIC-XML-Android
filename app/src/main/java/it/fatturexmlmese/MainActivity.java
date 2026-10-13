package it.fatturexmlmese;

import android.app.Activity;
import android.content.ContentValues;
import android.graphics.Color;
import android.graphics.Typeface;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.text.InputType;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.TextView;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainActivity extends Activity {

    private static final String API = "https://api-v2.fattureincloud.it";
    private final ExecutorService executor = Executors.newSingleThreadExecutor();

    private EditText companyInput;
    private EditText tokenInput;
    private EditText monthInput;
    private TextView status;
    private LinearLayout invoiceList;
    private Button exportAllButton;

    private String currentCompany = "";
    private String currentToken = "";
    private String currentMonth = "";
    private final List<Invoice> invoices = new ArrayList<>();

    static class Invoice {
        long id;
        String number;
        String date;
        String client;
        TextView state;
        Button button;
    }

    static class HttpResult {
        int code;
        byte[] body;
        String disposition;
        String text() { return new String(body, StandardCharsets.UTF_8); }
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        buildUi();
    }

    private int dp(int n) {
        return Math.round(n * getResources().getDisplayMetrics().density);
    }

    private TextView title(String text, int size) {
        TextView t = new TextView(this);
        t.setText(text);
        t.setTextSize(size);
        t.setTypeface(null, Typeface.BOLD);
        t.setTextColor(Color.rgb(25, 35, 50));
        t.setPadding(0, dp(8), 0, dp(6));
        return t;
    }

    private EditText field(String hint) {
        EditText e = new EditText(this);
        e.setHint(hint);
        e.setSingleLine(true);
        return e;
    }

    private Button button(String text) {
        Button b = new Button(this);
        b.setText(text);
        b.setAllCaps(false);
        return b;
    }

    private void buildUi() {
        ScrollView scroll = new ScrollView(this);
        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setPadding(dp(18), dp(18), dp(18), dp(30));
        root.setBackgroundColor(Color.rgb(246, 248, 251));
        scroll.addView(root);

        root.addView(title("Fatture in Cloud — XML", 24));
        TextView sub = new TextView(this);
        sub.setText("Carica un mese e scarica ogni XML singolarmente.");
        sub.setTextColor(Color.DKGRAY);
        root.addView(sub);

        root.addView(title("ID azienda", 13));
        companyInput = field("es. 1498046");
        companyInput.setInputType(InputType.TYPE_CLASS_NUMBER);
        companyInput.setText(getPreferences(MODE_PRIVATE).getString("company", ""));
        root.addView(companyInput);

        root.addView(title("Token personale", 13));
        tokenInput = field("Incolla qui il token");
        tokenInput.setInputType(InputType.TYPE_CLASS_TEXT | InputType.TYPE_TEXT_VARIATION_PASSWORD);
        root.addView(tokenInput);

        root.addView(title("Mese", 13));
        monthInput = field("AAAA-MM, es. 2026-09");
        monthInput.setText(YearMonth.now().toString());
        root.addView(monthInput);

        Button verify = button("Verifica collegamento");
        verify.setOnClickListener(v -> verify());
        root.addView(verify);

        Button load = button("Carica fatture del mese");
        load.setOnClickListener(v -> loadInvoices());
        root.addView(load);

        exportAllButton = button("Esporta tutte — una alla volta");
        exportAllButton.setEnabled(false);
        exportAllButton.setOnClickListener(v -> exportAll());
        root.addView(exportAllButton);

        status = new TextView(this);
        status.setText("Pronto.");
        status.setPadding(0, dp(12), 0, dp(12));
        status.setTextColor(Color.rgb(60, 70, 85));
        root.addView(status);

        root.addView(title("Fatture", 18));
        invoiceList = new LinearLayout(this);
        invoiceList.setOrientation(LinearLayout.VERTICAL);
        root.addView(invoiceList);

        TextView footer = new TextView(this);
        footer.setText("Salvataggio: Download/FattureInCloud/AAAA-MM. Il token non viene memorizzato.");
        footer.setTextSize(12);
        footer.setTextColor(Color.GRAY);
        footer.setPadding(0, dp(14), 0, 0);
        root.addView(footer);

        setContentView(scroll);
    }

    private boolean captureInputs() {
        currentCompany = companyInput.getText().toString().trim();
        currentToken = tokenInput.getText().toString().trim();
        currentMonth = monthInput.getText().toString().trim();

        if (currentToken.toLowerCase(Locale.ROOT).startsWith("bearer ")) {
            currentToken = currentToken.substring(7).trim();
        }
        if (currentCompany.isEmpty()) return fail("Inserisci l'ID azienda.");
        if (currentToken.isEmpty()) return fail("Inserisci il token personale.");
        if (!currentMonth.matches("\\d{4}-(0[1-9]|1[0-2])")) return fail("Mese non valido. Usa AAAA-MM.");

        getPreferences(MODE_PRIVATE).edit().putString("company", currentCompany).apply();
        return true;
    }

    private boolean fail(String text) {
        status.setText("Errore: " + text);
        return false;
    }

    private void setStatus(String text) {
        runOnUiThread(() -> status.setText(text));
    }

    private void verify() {
        if (!captureInputs()) return;
        final String company = currentCompany;
        final String token = currentToken;
        setStatus("Verifica collegamento…");

        executor.execute(() -> {
            try {
                HttpResult r = get(API + "/c/" + enc(company) + "/company/info", token, "application/json");
                if (r.code < 200 || r.code >= 300) throw new Exception(error(r));
                JSONObject j = new JSONObject(r.text());
                JSONObject d = j.optJSONObject("data");
                String name = d == null ? "azienda" : d.optString("name", "azienda");
                setStatus("Collegamento riuscito: " + name);
            } catch (Exception e) {
                setStatus("Errore: " + e.getMessage());
            }
        });
    }

    private void loadInvoices() {
        if (!captureInputs()) return;
        final String company = currentCompany;
        final String token = currentToken;
        final String ym = currentMonth;

        invoices.clear();
        invoiceList.removeAllViews();
        exportAllButton.setEnabled(false);
        setStatus("Carico le fatture…");

        executor.execute(() -> {
            try {
                YearMonth month = YearMonth.parse(ym);
                String first = month.atDay(1).toString();
                String last = month.atEndOfMonth().toString();
                int page = 1;

                while (true) {
                    String q = "type = 'invoice' and date >= '" + first + "' and date <= '" + last + "'";
                    String url = API + "/c/" + enc(company) + "/issued_documents"
                            + "?type=invoice&q=" + enc(q)
                            + "&page=" + page + "&per_page=100&fieldset=basic";

                    HttpResult r = get(url, token, "application/json");
                    if (r.code < 200 || r.code >= 300) throw new Exception(error(r));
                    JSONObject j = new JSONObject(r.text());
                    JSONArray data = j.optJSONArray("data");
                    if (data == null || data.length() == 0) break;

                    for (int i = 0; i < data.length(); i++) {
                        JSONObject d = data.getJSONObject(i);
                        Invoice inv = new Invoice();
                        inv.id = d.optLong("id");
                        inv.number = d.optString("number", String.valueOf(inv.id));
                        inv.date = d.optString("date", "");
                        JSONObject entity = d.optJSONObject("entity");
                        inv.client = entity == null ? "" : entity.optString("name", "");
                        invoices.add(inv);
                    }

                    String next = j.optString("next_page_url", "");
                    if (next.isEmpty() || "null".equals(next)) break;
                    page++;
                    if (page > 1000) throw new Exception("Troppe pagine.");
                }

                runOnUiThread(() -> {
                    renderInvoices();
                    exportAllButton.setEnabled(!invoices.isEmpty());
                    status.setText("Trovate " + invoices.size() + " fatture per " + ym + ".");
                });
            } catch (Exception e) {
                setStatus("Errore: " + e.getMessage());
            }
        });
    }

    private void renderInvoices() {
        invoiceList.removeAllViews();
        for (Invoice inv : invoices) {
            LinearLayout row = new LinearLayout(this);
            row.setOrientation(LinearLayout.VERTICAL);
            row.setPadding(dp(12), dp(10), dp(12), dp(10));
            row.setBackgroundColor(Color.WHITE);

            TextView t = title("Fattura " + inv.number + "  •  " + inv.date, 16);
            row.addView(t);

            TextView client = new TextView(this);
            client.setText(inv.client.isEmpty() ? "Cliente non indicato" : inv.client);
            client.setTextColor(Color.DKGRAY);
            row.addView(client);

            inv.state = new TextView(this);
            inv.state.setText("Non esportata");
            inv.state.setTextColor(Color.GRAY);
            inv.state.setPadding(0, dp(4), 0, dp(4));
            row.addView(inv.state);

            inv.button = button("Esporta XML");
            inv.button.setOnClickListener(v -> exportOne(inv));
            row.addView(inv.button);

            LinearLayout.LayoutParams p = new LinearLayout.LayoutParams(
                    LinearLayout.LayoutParams.MATCH_PARENT,
                    LinearLayout.LayoutParams.WRAP_CONTENT);
            p.setMargins(0, 0, 0, dp(10));
            invoiceList.addView(row, p);
        }
    }

    private void exportOne(Invoice inv) {
        if (!captureInputs()) return;
        final String company = currentCompany;
        final String token = currentToken;
        final String ym = currentMonth;
        inv.button.setEnabled(false);
        inv.state.setText("Esporto…");

        executor.execute(() -> {
            try {
                String file = downloadXml(company, token, ym, inv);
                runOnUiThread(() -> {
                    inv.state.setText("✓ Salvata: " + file);
                    inv.state.setTextColor(Color.rgb(0, 120, 75));
                    inv.button.setEnabled(true);
                });
                setStatus("Fattura " + inv.number + " esportata.");
            } catch (Exception e) {
                runOnUiThread(() -> {
                    inv.state.setText("✕ " + e.getMessage());
                    inv.state.setTextColor(Color.rgb(180, 40, 40));
                    inv.button.setEnabled(true);
                });
            }
        });
    }

    private void exportAll() {
        if (!captureInputs()) return;
        final String company = currentCompany;
        final String token = currentToken;
        final String ym = currentMonth;
        exportAllButton.setEnabled(false);

        executor.execute(() -> {
            int ok = 0;
            int ko = 0;
            for (int i = 0; i < invoices.size(); i++) {
                Invoice inv = invoices.get(i);
                int pos = i + 1;
                setStatus("Esporto " + pos + " / " + invoices.size() + " — fattura " + inv.number);
                runOnUiThread(() -> inv.state.setText("Esporto…"));
                try {
                    String file = downloadXml(company, token, ym, inv);
                    ok++;
                    runOnUiThread(() -> {
                        inv.state.setText("✓ Salvata: " + file);
                        inv.state.setTextColor(Color.rgb(0, 120, 75));
                    });
                } catch (Exception e) {
                    ko++;
                    runOnUiThread(() -> {
                        inv.state.setText("✕ " + e.getMessage());
                        inv.state.setTextColor(Color.rgb(180, 40, 40));
                    });
                }
            }
            int done = ok;
            int errors = ko;
            runOnUiThread(() -> {
                status.setText("Completato: " + done + " XML salvati, " + errors + " non esportati.");
                exportAllButton.setEnabled(true);
            });
        });
    }

    private String downloadXml(String company, String token, String ym, Invoice inv) throws Exception {
        String url = API + "/c/" + enc(company)
                + "/issued_documents/" + inv.id
                + "/e_invoice/xml?avoid_einvoice_check=1&include_attachment=false";

        HttpResult r = get(url, token, "text/xml,application/xml,*/*");
        if (r.code < 200 || r.code >= 300) throw new Exception(shortError(r));

        String filename = filenameFromDisposition(r.disposition);
        if (filename.isEmpty()) {
            filename = safe(inv.date) + "_fattura_" + safe(inv.number)
                    + (inv.client.isEmpty() ? "" : "_" + safe(inv.client)) + ".xml";
        }
        if (!filename.toLowerCase(Locale.ROOT).endsWith(".xml")) filename += ".xml";
        saveDownload(ym, filename, r.body);
        return filename;
    }

    private void saveDownload(String ym, String filename, byte[] data) throws Exception {
        ContentValues values = new ContentValues();
        values.put(MediaStore.Downloads.DISPLAY_NAME, filename);
        values.put(MediaStore.Downloads.MIME_TYPE, "text/xml");
        values.put(MediaStore.Downloads.RELATIVE_PATH,
                Environment.DIRECTORY_DOWNLOADS + "/FattureInCloud/" + ym);

        Uri uri = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
        if (uri == null) throw new Exception("Impossibile creare il file nei Download.");

        try (OutputStream out = getContentResolver().openOutputStream(uri)) {
            if (out == null) throw new Exception("Impossibile aprire il file.");
            out.write(data);
        } catch (Exception e) {
            getContentResolver().delete(uri, null, null);
            throw e;
        }
    }

    private HttpResult get(String url, String token, String accept) throws Exception {
        HttpURLConnection c = (HttpURLConnection) new URL(url).openConnection();
        c.setRequestMethod("GET");
        c.setConnectTimeout(20000);
        c.setReadTimeout(60000);
        c.setRequestProperty("Authorization", "Bearer " + token);
        c.setRequestProperty("Accept", accept);
        c.setRequestProperty("User-Agent", "FIC-XML-Android/1.0");

        HttpResult r = new HttpResult();
        r.code = c.getResponseCode();
        r.disposition = c.getHeaderField("Content-Disposition");
        InputStream in = r.code >= 400 ? c.getErrorStream() : c.getInputStream();
        r.body = in == null ? new byte[0] : readAll(in);
        if (in != null) in.close();
        c.disconnect();
        return r;
    }

    private byte[] readAll(InputStream in) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        byte[] buf = new byte[8192];
        int n;
        while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
        return out.toByteArray();
    }

    private String shortError(HttpResult r) {
        String t = r.text();
        if (r.code == 401) return "401 — token non valido o scaduto";
        if (r.code == 403) return "403 — permesso insufficiente";
        if (r.code == 422 && t.toLowerCase(Locale.ROOT).contains("fattura elettronica")) {
            return "422 — XML non disponibile tramite API pubblica";
        }
        if (t.length() > 180) t = t.substring(0, 180);
        return r.code + " — " + t;
    }

    private String error(HttpResult r) {
        String t = r.text();
        if (t.length() > 250) t = t.substring(0, 250);
        return "HTTP " + r.code + ": " + t;
    }

    private String filenameFromDisposition(String cd) {
        if (cd == null) return "";
        int p = cd.toLowerCase(Locale.ROOT).indexOf("filename=");
        if (p < 0) return "";
        String s = cd.substring(p + 9).trim();
        int semi = s.indexOf(';');
        if (semi >= 0) s = s.substring(0, semi);
        s = s.replace("\"", "");
        return safe(s);
    }

    private String safe(String s) {
        if (s == null) return "";
        String out = s.trim().replaceAll("[^A-Za-z0-9._-]+", "_");
        if (out.length() > 100) out = out.substring(0, 100);
        return out;
    }

    private String enc(String s) throws Exception {
        return URLEncoder.encode(s, "UTF-8").replace("+", "%20");
    }

    @Override
    protected void onDestroy() {
        executor.shutdownNow();
        super.onDestroy();
    }
}
