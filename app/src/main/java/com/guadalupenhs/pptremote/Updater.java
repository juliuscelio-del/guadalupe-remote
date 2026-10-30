package com.guadalupenhs.pptremote;

import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Handler;
import android.os.Looper;
import android.provider.Settings;
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
 * On app start:
 *  1. shows "searching for updates" for at least SEARCH_MS
 *  2. no update  -> shows "you're up to date" for NO_UPDATE_MS, then closes
 *  3. update     -> shows "downloading" and installs automatically
 *  4. offline    -> just closes the dialog and lets the app carry on
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

    private static final int NONE = 0, UPDATE = 1, ERROR = 2;

    private static final Handler UI = new Handler(Looper.getMainLooper());

    public static void check(final Activity activity) {
        final long started = System.currentTimeMillis();
        final AlertDialog searching = new AlertDialog.Builder(activity)
                .setMessage(MSG_SEARCHING)
                .setCancelable(false)
                .create();
        searching.show();

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
                dismiss(searching);
                if (activity.isFinishing() || activity.isDestroyed()) return;
                if (fr == NONE) {
                    showNoUpdate(activity);
                } else if (fr == UPDATE) {
                    startDownload(activity, url);
                }
            }, wait);
        }).start();
    }

    private static void showNoUpdate(Activity activity) {
        final AlertDialog d = new AlertDialog.Builder(activity)
                .setMessage(MSG_NO_UPDATE)
                .setCancelable(true)
                .create();
        d.show();
        UI.postDelayed(() -> dismiss(d), NO_UPDATE_MS);
    }

    private static void startDownload(final Activity activity, final String url) {
        final AlertDialog d = new AlertDialog.Builder(activity)
                .setMessage(MSG_DOWNLOADING)
                .setCancelable(false)
                .create();
        d.show();

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
                    dismiss(d);
                    if (!activity.isFinishing() && !activity.isDestroyed()) install(activity, done);
                });
            } catch (Exception e) {
                UI.post(() -> {
                    dismiss(d);
                    Toast.makeText(activity, MSG_FAILED,
                            Toast.LENGTH_LONG).show();
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

    private static void dismiss(AlertDialog d) {
        try {
            if (d != null && d.isShowing()) d.dismiss();
        } catch (Exception ignored) {
        }
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
