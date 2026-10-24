package com.guadalupenhs.pptremote;

import android.graphics.Color;
import android.graphics.Typeface;
import android.graphics.drawable.GradientDrawable;
import android.os.Bundle;
import android.text.InputType;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.view.inputmethod.EditorInfo;
import android.view.inputmethod.InputMethodManager;
import android.webkit.CookieManager;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Button;
import android.widget.EditText;
import android.widget.FrameLayout;
import android.widget.ImageView;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.TextView;
import android.widget.Toast;

import androidx.activity.OnBackPressedCallback;
import androidx.activity.result.ActivityResultLauncher;
import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;

import com.journeyapps.barcodescanner.ScanContract;
import com.journeyapps.barcodescanner.ScanOptions;

import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;

/**
 * Guadalupe NHS Remote
 * Screen 1: scan the QR code shown by the computer app
 * Screen 2: enter the PIN
 * Screen 3: the computer's own remote pages (4 permission options, PowerPoint remote, desktop control)
 */
public class MainActivity extends AppCompatActivity {

    private static final int BG = Color.parseColor("#E4F1FF");
    private static final int BLUE = Color.parseColor("#2F6DF6");
    private static final int INK = Color.parseColor("#12306B");
    private static final int MUTED = Color.parseColor("#5F78A8");
    private static final int GREEN = Color.parseColor("#1FC76A");
    private static final int RED = Color.parseColor("#FF5A5F");
    private static final int LINE = Color.parseColor("#C6DBFF");

    private static final int MATCH = ViewGroup.LayoutParams.MATCH_PARENT;
    private static final int WRAP = ViewGroup.LayoutParams.WRAP_CONTENT;

    private FrameLayout root;
    private ScrollView homeView;
    private ScrollView pinView;
    private WebView web;
    private EditText pinInput;
    private EditText manualInput;
    private TextView pinHost;
    private TextView pinError;
    private Button connectBtn;

    private String baseUrl;
    private boolean lostShown = false;
    private ActivityResultLauncher<ScanOptions> scanner;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        root = new FrameLayout(this);
        root.setBackgroundColor(BG);

        buildHome();
        buildPin();
        buildWeb();
        setContentView(root);
        show(homeView);

        scanner = registerForActivityResult(new ScanContract(), result -> {
            if (result.getContents() != null) {
                onScanned(result.getContents());
            }
        });

        getOnBackPressedDispatcher().addCallback(this, new OnBackPressedCallback(true) {
            @Override
            public void handleOnBackPressed() {
                onBack();
            }
        });
    }

    // ------------------------------------------------------------ UI helpers
    private int dp(float v) {
        return Math.round(v * getResources().getDisplayMetrics().density);
    }

    private GradientDrawable box(int fill, float radiusDp, int stroke, float strokeDp) {
        GradientDrawable g = new GradientDrawable();
        g.setColor(fill);
        g.setCornerRadius(dp(radiusDp));
        if (stroke != 0) {
            g.setStroke(dp(strokeDp), stroke);
        }
        return g;
    }

    private Button button(String label, int color) {
        Button b = new Button(this);
        b.setText(label);
        b.setAllCaps(false);
        b.setTextColor(Color.WHITE);
        b.setTextSize(18);
        b.setTypeface(Typeface.DEFAULT_BOLD);
        b.setBackground(box(color, 16, 0, 0));
        b.setStateListAnimator(null);
        b.setPadding(dp(24), dp(14), dp(24), dp(14));
        return b;
    }

    private TextView label(String s, float sp, int color, boolean bold) {
        TextView t = new TextView(this);
        t.setText(s);
        t.setTextSize(sp);
        t.setTextColor(color);
        t.setGravity(Gravity.CENTER);
        if (bold) {
            t.setTypeface(Typeface.DEFAULT_BOLD);
        }
        return t;
    }

    private LinearLayout.LayoutParams lp(int w, int h, int topPx) {
        LinearLayout.LayoutParams p = new LinearLayout.LayoutParams(w, h);
        p.topMargin = topPx;
        p.gravity = Gravity.CENTER_HORIZONTAL;
        return p;
    }

    private ImageView logo(int sizeDp) {
        ImageView i = new ImageView(this);
        i.setImageResource(R.drawable.logo);
        i.setScaleType(ImageView.ScaleType.FIT_CENTER);
        i.setLayoutParams(lp(dp(sizeDp), dp(sizeDp), 0));
        return i;
    }

    private LinearLayout column() {
        LinearLayout c = new LinearLayout(this);
        c.setOrientation(LinearLayout.VERTICAL);
        c.setGravity(Gravity.CENTER);
        c.setPadding(dp(28), dp(28), dp(28), dp(28));
        return c;
    }

    private ScrollView scroller(LinearLayout child) {
        ScrollView sv = new ScrollView(this);
        sv.setFillViewport(true);
        sv.addView(child, new FrameLayout.LayoutParams(MATCH, WRAP));
        return sv;
    }

    // ------------------------------------------------------------ screen 1: scan
    private void buildHome() {
        LinearLayout c = column();
        c.addView(logo(140));
        c.addView(label("Guadalupe NHS", 28, BLUE, true), lp(WRAP, WRAP, dp(16)));
        c.addView(label("PPT Remote", 16, MUTED, false), lp(WRAP, WRAP, dp(2)));
        c.addView(label("Open the Guadalupe NHS Remote app on the computer, then connect:",
                15, INK, false), lp(MATCH, WRAP, dp(28)));

        // Option 1: scan the QR code
        Button scan = button("\uD83D\uDCF7  Scan QR code", BLUE);
        scan.setOnClickListener(v -> startScan());
        c.addView(scan, lp(MATCH, WRAP, dp(22)));

        // Option 2: type the IP address
        c.addView(label("\u2014  or type the address  \u2014", 14, MUTED, false), lp(MATCH, WRAP, dp(24)));
        manualInput = new EditText(this);
        manualInput.setInputType(InputType.TYPE_CLASS_TEXT | InputType.TYPE_TEXT_VARIATION_URI);
        manualInput.setImeOptions(EditorInfo.IME_ACTION_GO);
        manualInput.setSingleLine(true);
        manualInput.setHint("e.g. 192.168.1.23");
        manualInput.setGravity(Gravity.CENTER);
        manualInput.setTextSize(20);
        manualInput.setTextColor(INK);
        manualInput.setBackground(box(Color.WHITE, 16, LINE, 2));
        manualInput.setPadding(dp(16), dp(12), dp(16), dp(12));
        manualInput.setOnEditorActionListener((v, actionId, event) -> {
            if (actionId == EditorInfo.IME_ACTION_GO) {
                connectManual();
                return true;
            }
            return false;
        });
        c.addView(manualInput, lp(MATCH, WRAP, dp(10)));

        Button manual = button("Connect with this address", BLUE);
        manual.setTextColor(BLUE);
        manual.setBackground(box(Color.WHITE, 16, BLUE, 2));
        manual.setOnClickListener(v -> connectManual());
        c.addView(manual, lp(MATCH, WRAP, dp(10)));

        homeView = scroller(c);
        root.addView(homeView, new FrameLayout.LayoutParams(MATCH, MATCH));
    }

    private void startScan() {
        ScanOptions o = new ScanOptions();
        o.setDesiredBarcodeFormats(ScanOptions.QR_CODE);
        o.setPrompt("Point the camera at the QR code on the computer");
        o.setBeepEnabled(false);
        scanner.launch(o);
    }

    private void onScanned(String text) {
        String t = text.trim();
        if (!(t.startsWith("http://") || t.startsWith("https://"))) {
            Toast.makeText(this, "That QR code is not from the PPT Remote app.", Toast.LENGTH_LONG).show();
            return;
        }
        while (t.endsWith("/")) {
            t = t.substring(0, t.length() - 1);
        }
        goToPin(t);
    }

    /** Typed address: accepts 192.168.1.23, 192.168.1.23:5001 or http://192.168.1.23:5001 */
    private void connectManual() {
        String url = normalizeAddress(manualInput.getText().toString());
        if (url == null) {
            Toast.makeText(this, "Please type the address shown on the computer, e.g. 192.168.1.23",
                    Toast.LENGTH_LONG).show();
            return;
        }
        InputMethodManager m = (InputMethodManager) getSystemService(INPUT_METHOD_SERVICE);
        if (m != null) {
            m.hideSoftInputFromWindow(manualInput.getWindowToken(), 0);
        }
        goToPin(url);
    }

    private String normalizeAddress(String input) {
        String rest = input.trim().replaceFirst("(?i)^https?://", "");
        int slash = rest.indexOf('/');
        if (slash >= 0) {
            rest = rest.substring(0, slash);
        }
        if (!rest.matches("[A-Za-z0-9.\\-]+(:[0-9]{1,5})?")) {
            return null;
        }
        if (!rest.contains(":")) {
            rest = rest + ":5001";   // the computer app's default port
        }
        return "http://" + rest;
    }

    private void goToPin(String t) {
        baseUrl = t;
        pinHost.setText("Computer: " + t.replaceFirst("^https?://", ""));
        pinInput.setText("");
        pinError.setText("");
        setBusy(false);
        show(pinView);
        pinInput.postDelayed(() -> {
            pinInput.requestFocus();
            InputMethodManager m = (InputMethodManager) getSystemService(INPUT_METHOD_SERVICE);
            if (m != null) {
                m.showSoftInput(pinInput, InputMethodManager.SHOW_IMPLICIT);
            }
        }, 250);
    }

    // ------------------------------------------------------------ screen 2: PIN
    private void buildPin() {
        LinearLayout c = column();
        c.addView(logo(96));
        c.addView(label("Enter the PIN", 26, BLUE, true), lp(WRAP, WRAP, dp(14)));
        pinHost = label("", 14, MUTED, false);
        c.addView(pinHost, lp(WRAP, WRAP, dp(4)));
        c.addView(label("The PIN is shown in the window on the computer.", 14, INK, false),
                lp(MATCH, WRAP, dp(18)));

        pinInput = new EditText(this);
        pinInput.setInputType(InputType.TYPE_CLASS_NUMBER | InputType.TYPE_NUMBER_VARIATION_PASSWORD);
        pinInput.setImeOptions(EditorInfo.IME_ACTION_DONE);
        pinInput.setSingleLine(true);
        pinInput.setGravity(Gravity.CENTER);
        pinInput.setTextSize(30);
        pinInput.setTextColor(INK);
        pinInput.setLetterSpacing(0.3f);
        pinInput.setBackground(box(Color.WHITE, 16, LINE, 2));
        pinInput.setPadding(dp(16), dp(12), dp(16), dp(12));
        pinInput.setOnEditorActionListener((v, actionId, event) -> {
            if (actionId == EditorInfo.IME_ACTION_DONE) {
                submitPin();
                return true;
            }
            return false;
        });
        c.addView(pinInput, lp(MATCH, WRAP, dp(12)));

        pinError = label("", 15, RED, true);
        c.addView(pinError, lp(MATCH, WRAP, dp(10)));

        connectBtn = button("Connect", GREEN);
        connectBtn.setOnClickListener(v -> submitPin());
        c.addView(connectBtn, lp(MATCH, WRAP, dp(14)));

        TextView again = label("Back", 15, MUTED, true);
        again.setPadding(dp(16), dp(14), dp(16), dp(14));
        again.setOnClickListener(v -> show(homeView));
        c.addView(again, lp(WRAP, WRAP, dp(8)));

        pinView = scroller(c);
        root.addView(pinView, new FrameLayout.LayoutParams(MATCH, MATCH));
    }

    private void setBusy(boolean busy) {
        connectBtn.setEnabled(!busy);
        connectBtn.setText(busy ? "Connecting..." : "Connect");
    }

    private void submitPin() {
        final String pin = pinInput.getText().toString().trim();
        if (pin.isEmpty()) {
            pinError.setText("Please type the PIN.");
            return;
        }
        InputMethodManager m = (InputMethodManager) getSystemService(INPUT_METHOD_SERVICE);
        if (m != null) {
            m.hideSoftInputFromWindow(pinInput.getWindowToken(), 0);
        }
        pinError.setText("");
        setBusy(true);
        final String target = baseUrl;

        new Thread(() -> {
            int code = -1;
            boolean netError = false;
            try {
                byte[] body = ("pin=" + URLEncoder.encode(pin, "UTF-8")).getBytes("UTF-8");
                HttpURLConnection c = (HttpURLConnection) new URL(target + "/login").openConnection();
                c.setInstanceFollowRedirects(false);
                c.setConnectTimeout(5000);
                c.setReadTimeout(8000);
                c.setRequestMethod("POST");
                c.setDoOutput(true);
                c.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
                c.setFixedLengthStreamingMode(body.length);
                try (OutputStream os = c.getOutputStream()) {
                    os.write(body);
                }
                code = c.getResponseCode();
                c.disconnect();
            } catch (Exception e) {
                netError = true;
            }
            final int fc = code;
            final boolean fe = netError;
            runOnUiThread(() -> {
                setBusy(false);
                if (fe) {
                    pinError.setText("Can't reach the computer. Are both on the same Wi-Fi?");
                } else if (fc == 301 || fc == 302 || fc == 303) {
                    openRemote(pin);
                } else {
                    pinError.setText("Wrong PIN. Please try again.");
                    pinInput.setText("");
                }
            });
        }).start();
    }

    // ------------------------------------------------------------ screen 3: the remote
    private void buildWeb() {
        web = new WebView(this);
        web.setBackgroundColor(BG);
        WebSettings ws = web.getSettings();
        ws.setJavaScriptEnabled(true);
        ws.setDomStorageEnabled(true);
        web.setWebViewClient(new WebViewClient() {
            @Override
            public void onReceivedError(WebView view, WebResourceRequest request, WebResourceError error) {
                if (request.isForMainFrame()) {
                    runOnUiThread(() -> lost());
                }
            }
        });
        web.setVisibility(View.GONE);
        root.addView(web, new FrameLayout.LayoutParams(MATCH, MATCH));
    }

    private void openRemote(String pin) {
        try {
            lostShown = false;
            byte[] body = ("pin=" + URLEncoder.encode(pin, "UTF-8")).getBytes("UTF-8");
            // Logs the page in (cookie), then the computer redirects to the permissions page.
            web.postUrl(baseUrl + "/login", body);
            show(web);
            web.setKeepScreenOn(true);
        } catch (Exception e) {
            pinError.setText("Something went wrong. Please try again.");
        }
    }

    private void lost() {
        if (lostShown || web.getVisibility() != View.VISIBLE) {
            return;
        }
        lostShown = true;
        new AlertDialog.Builder(this)
                .setTitle("Connection lost")
                .setMessage("Can't reach the computer. Check that the app is still running and "
                        + "that both are on the same Wi-Fi.")
                .setCancelable(false)
                .setPositiveButton("OK", (d, w) -> {
                    lostShown = false;
                    disconnect();
                })
                .show();
    }

    private void disconnect() {
        web.loadUrl("about:blank");
        web.clearHistory();
        web.setKeepScreenOn(false);
        CookieManager.getInstance().removeAllCookies(null);
        CookieManager.getInstance().flush();
        baseUrl = null;
        show(homeView);
    }

    // ------------------------------------------------------------ navigation
    private void show(View v) {
        homeView.setVisibility(v == homeView ? View.VISIBLE : View.GONE);
        pinView.setVisibility(v == pinView ? View.VISIBLE : View.GONE);
        web.setVisibility(v == web ? View.VISIBLE : View.GONE);
    }

    private void onBack() {
        if (web.getVisibility() == View.VISIBLE) {
            new AlertDialog.Builder(this)
                    .setTitle("Disconnect?")
                    .setMessage("Stop controlling the computer and go back to the scan screen?")
                    .setPositiveButton("Disconnect", (d, w) -> disconnect())
                    .setNegativeButton("Stay", null)
                    .show();
        } else if (pinView.getVisibility() == View.VISIBLE) {
            show(homeView);
        } else {
            finish();
        }
    }
}
