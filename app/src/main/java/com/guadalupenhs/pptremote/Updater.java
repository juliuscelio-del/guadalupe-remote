package com.guadalupenhs.pptremote;

import android.app.Activity;
import android.content.Intent;
import android.content.res.ColorStateList;
import android.graphics.Color;
import android.graphics.Typeface;
import android.graphics.drawable.ColorDrawable;
import android.graphics.drawable.GradientDrawable;
import android.net.Uri;
import android.os.Build;
import android.os.Handler;
import android.os.Looper;
import android.provider.Settings;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.widget.LinearLayout;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AlertDialog;
import androidx.core.content.FileProvider;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;

/**
 * On app start (one small card that changes color and text as it goes):
 *  1. "searching for updates" (blue) for at least SEARCH_MS
 *  2. no update  -> "you're up to date" (green) for NO_UPDATE_MS, then closes
 *  3. update     -> "downloading" (orange) and installs automatically
 *  4. offline    -> just closes the card and lets the app carry on
 */
public class Updater {

    private static final String API =
            "https://api.github.com/repos/juliuscelio-del/guadalupe-remote/releases/latest";

    // ---- change the texts / timings here ----
    private static final String MSG_SEARCHING = "search sa tah kadali ha ug naa ba updates";
    private static final String MSG_NO_UPDATE = "Pwed pa ka mopadajun gamit ane kai Updated pa";
    private static final String MSG_DOWNLOADING = "downloading sa ta ha sa updates...paabot lang";
    private static final String MSG_FAILED = "Dili ma-download ang update. Sulayi pag usab sa ulahi.";
    private static final long SEARCH_MS = 10000;     // 10 seconds
    private static final long NO_UPDATE_MS = 5000;   // 5 seconds
    // -----------------------------------------

    // ---- vibrant palette ----
    private static final int C_BLUE = Color.parseColor("#2F6DF6");
    private static final int C_ORANGE = Color.parseColor("#FFA31A");
    private static final int C_GREEN = Color.parseColor("#1FC76A");
    private static final int C_RED = Color.parseColor("#FF5A5F");
    private static final int C_TRACK = Color.parseColor("#DBE8FF");
    private static final int[] STRIPE = {
            Color.parseColor("#2F6DF6"), Color.parseColor("#4AA3FF"), Color.parseColor("#1FC76A"),
            Color.parseColor("#FFA31A"), Color.parseColor("#8B5CF6")};

    private static final int NONE = 0, UPDATE = 1, ERROR = 2;

    private static final Handler UI = new Handler(Looper.getMainLooper());

    // ------------------------------------------------------------ the card
    private static class Card {
        final AlertDialog dialog;
        final TextView text;
        final ProgressBar bar;

        Card(Activity a) {
            LinearLayout root = new LinearLayout(a);
            root.setOrientation(LinearLayout.VERTICAL);
            root.setMinimumWidth(dp(a, 290));

            GradientDrawable bg = new GradientDrawable();
            bg.setColor(Color.WHITE);
            bg.setCornerRadius(dp(a, 24));
            root.setBackground(bg);
            root.setClipToOutline(true);           // rounds the colored stripe too

            // colorful top stripe
            LinearLayout stripe = new LinearLayout(a);
            stripe.setOrientation(LinearLayout.HORIZONTAL);
            for (int c : STRIPE) {
                View v = new View(a);
                v.setBackgroundColor(c);
                stripe.addView(v, new LinearLayout.LayoutParams(0, ViewGroup.LayoutParams.MATCH_PARENT, 1f));
            }
            root.addView(stripe, new LinearLayout.LayoutParams(
                    ViewGroup.LayoutParams.MATCH_PARENT, dp(a, 5)));

            // message
            text = new TextView(a);
            text.setTextSize(16);
            text.setTypeface(Typeface.DEFAULT_BOLD);
            text.setGravity(Gravity.CENTER);
            text.setLineSpacing(0, 1.1f);
            LinearLayout.LayoutParams tp = new LinearLayout.LayoutParams(
                    ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.WRAP_CONTENT);
            tp.setMargins(dp(a, 26), dp(a, 30), dp(a, 26), dp(a, 22));
            root.addView(text, tp);

            // progress pill
            bar = new ProgressBar(a, null, android.R.attr.progressBarStyleHorizontal);
            bar.setMax(100);
            bar.setIndeterminate(true);
            LinearLayout.LayoutParams bp = new LinearLayout.LayoutParams(
                    ViewGroup.LayoutParams.MATCH_PARENT, dp(a, 6));
            bp.setMargins(dp(a, 26), 0, dp(a, 26), dp(a, 28));
            root.addView(bar, bp);

            dialog = new AlertDialog.Builder(a).setView(root).setCancelable(false).create();
            if (dialog.getWindow() != null) {
                dialog.getWindow().setBackgroundDrawable(new ColorDrawable(Color.TRANSPARENT));
            }
        }

        void show(String msg, int color, boolean busy) {
            set(msg, color, busy);
            dialog.show();
        }

        /** Change text + color. busy = animated bar, otherwise a full solid bar. */
        void set(String msg, int color, boolean busy) {
            text.setText(msg);
            text.setTextColor(darker(color));
            ColorStateList tint = ColorStateList.valueOf(color);
            bar.setIndeterminateTintList(tint);
            bar.setProgressTintList(tint);
            bar.setProgressBackgroundTintList(ColorStateList.valueOf(C_TRACK));
            if (busy) {
                bar.setIndeterminate(true);
            } else {
                bar.setIndeterminate(false);
                bar.setProgress(100);
            }
        }

        void dismiss() {
            try {
                if (dialog.isShowing()) dialog.dismiss();
            } catch (Exception ignored) {
            }
        }
    }

    private static int dp(Activity a, int v) {
        return Math.round(v * a.getResources().getDisplayMetrics().density);
    }

    private static int darker(int c) {
        float[] hsv = new float[3];
        Color.colorToHSV(c, hsv);
        hsv[2] *= 0.62f;
        return Color.HSVToColor(hsv);
    }

    // ------------------------------------------------------------ the flow
    public static void check(final Activity activity) {
        final long started = System.currentTimeMillis();
        final Card card = new Card(activity);
        card.show(MSG_SEARCHING, C_BLUE, true);

        new Thread(() -> {
            int result = ERROR;
            String apkUrl = null;
            try {
                HttpURLConnection c = (HttpURLConnection) new URL(API).openConnection();
                c.setConnectTimeout(8000);
                c.setReadTimeout(8000);
                c.setRequestProperty("Accept", "application/vnd.github+json");
                c.setRequestProperty("User-Agent", "GNHS-Remote");
                if (c.getResponseCode() == 200) {
                    StringBuilder sb = new StringBuilder();
                    try (BufferedReader r = new BufferedReader(
                            new InputStreamReader(c.getInputStream(), "UTF-8"))) {
                        String line;
                        while ((line = r.readLine()) != null) sb.append(line);
                    }
                    JSONObject json = new JSONObject(sb.toString());
                    String latest = json.getString("tag_name").replaceFirst("^[vV]", "");
                    String current = activity.getPackageManager()
                            .getPackageInfo(activity.getPackageName(), 0).versionName;

                    if (current == null || !isNewer(latest, current)) {
                        result = NONE;
                    } else {
                        JSONArray assets = json.getJSONArray("assets");
                        for (int i = 0; i < assets.length(); i++) {
                            JSONObject a = assets.getJSONObject(i);
                            if (a.getString("name").endsWith(".apk")) {
                                apkUrl = a.getString("browser_download_url");
                                break;
                            }
                        }
                        result = (apkUrl != null) ? UPDATE : NONE;
                    }
                }
            } catch (Exception ignored) {
                result = ERROR;   // offline or GitHub unreachable
            }

            final int fr = result;
            final String url = apkUrl;
            long wait = Math.max(0, SEARCH_MS - (System.currentTimeMillis() - started));
            UI.postDelayed(() -> {
                if (activity.isFinishing() || activity.isDestroyed()) {
                    card.dismiss();
                    return;
                }
                if (fr == NONE) {
                    showNoUpdate(card);
                } else if (fr == UPDATE) {
                    startDownload(activity, card, url);
                } else {
                    card.dismiss();
                }
            }, wait);
        }).start();
    }

    private static void showNoUpdate(final Card card) {
        card.set(MSG_NO_UPDATE, C_GREEN, false);
        card.dialog.setCancelable(true);          // tap outside to close early
        UI.postDelayed(card::dismiss, NO_UPDATE_MS);
    }

    private static void startDownload(final Activity activity, final Card card, final String url) {
        card.set(MSG_DOWNLOADING, C_ORANGE, true);

        new Thread(() -> {
            File file = null;
            try {
                File dir = new File(activity.getCacheDir(), "updates");
                dir.mkdirs();
                file = new File(dir, "update.apk");

                HttpURLConnection c = (HttpURLConnection) new URL(url).openConnection();
                c.setConnectTimeout(10000);
                c.setReadTimeout(30000);
                c.setRequestProperty("User-Agent", "GNHS-Remote");
                try (InputStream in = c.getInputStream();
                     FileOutputStream out = new FileOutputStream(file)) {
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = in.read(buf)) > 0) out.write(buf, 0, n);
                }
                final File done = file;
                UI.post(() -> {
                    card.dismiss();
                    if (!activity.isFinishing() && !activity.isDestroyed()) install(activity, done);
                });
            } catch (Exception e) {
                UI.post(() -> {
                    card.set(MSG_FAILED, C_RED, false);
                    card.dialog.setCancelable(true);
                    UI.postDelayed(card::dismiss, 4000);
                });
            }
        }).start();
    }

    private static void install(Activity activity, File file) {
        if (Build.VERSION.SDK_INT >= 26 && !activity.getPackageManager().canRequestPackageInstalls()) {
            Toast.makeText(activity,
                    "Allow \"Install unknown apps\" for this app, then open the app again.",
                    Toast.LENGTH_LONG).show();
            activity.startActivity(new Intent(Settings.ACTION_MANAGE_UNKNOWN_APP_SOURCES,
                    Uri.parse("package:" + activity.getPackageName())));
            return;
        }
        Uri uri = FileProvider.getUriForFile(activity,
                activity.getPackageName() + ".fileprovider", file);
        Intent i = new Intent(Intent.ACTION_VIEW);
        i.setDataAndType(uri, "application/vnd.android.package-archive");
        i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        activity.startActivity(i);
    }

    private static boolean isNewer(String a, String b) {
        String[] x = a.split("\\.");
        String[] y = b.split("\\.");
        for (int i = 0; i < Math.max(x.length, y.length); i++) {
            int p = i < x.length ? num(x[i]) : 0;
            int q = i < y.length ? num(y[i]) : 0;
            if (p != q) return p > q;
        }
        return false;
    }

    private static int num(String s) {
        try {
            return Integer.parseInt(s.replaceAll("[^0-9]", ""));
        } catch (Exception e) {
            return 0;
        }
    }
}
