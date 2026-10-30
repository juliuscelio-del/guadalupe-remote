package com.guadalupenhs.pptremote;

import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
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

/** Checks GitHub Releases for a newer APK and installs it. Silent if offline. */
public class Updater {

    private static final String API =
            "https://api.github.com/repos/juliuscelio-del/guadalupe-remote/releases/latest";

    public static void check(final Activity activity) {
        new Thread(() -> {
            try {
                HttpURLConnection c = (HttpURLConnection) new URL(API).openConnection();
                c.setConnectTimeout(8000);
                c.setReadTimeout(8000);
                c.setRequestProperty("Accept", "application/vnd.github+json");
                c.setRequestProperty("User-Agent", "GNHS-Remote");
                if (c.getResponseCode() != 200) return;

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
                if (current == null || !isNewer(latest, current)) return;

                String apkUrl = null;
                JSONArray assets = json.getJSONArray("assets");
                for (int i = 0; i < assets.length(); i++) {
                    JSONObject a = assets.getJSONObject(i);
                    if (a.getString("name").endsWith(".apk")) {
                        apkUrl = a.getString("browser_download_url");
                        break;
                    }
                }
                if (apkUrl == null) return;

                final String url = apkUrl;
                final String ver = latest;
                activity.runOnUiThread(() -> {
                    if (activity.isFinishing()) return;
                    new AlertDialog.Builder(activity)
                            .setTitle("Update available")
                            .setMessage("Version " + ver + " is ready. Download and install it now?")
                            .setPositiveButton("Update", (d, w) -> download(activity, url))
                            .setNegativeButton("Later", null)
                            .show();
                });
            } catch (Exception ignored) {
                // offline or GitHub unreachable: just carry on
            }
        }).start();
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

    private static void download(final Activity activity, final String url) {
        Toast.makeText(activity, "Downloading update...", Toast.LENGTH_SHORT).show();
        new Thread(() -> {
            try {
                File dir = new File(activity.getCacheDir(), "updates");
                dir.mkdirs();
                File file = new File(dir, "update.apk");

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
                activity.runOnUiThread(() -> install(activity, file));
            } catch (Exception e) {
                activity.runOnUiThread(() -> Toast.makeText(activity,
                        "Download failed. Please try again later.", Toast.LENGTH_LONG).show());
            }
        }).start();
    }

    private static void install(Activity activity, File file) {
        if (Build.VERSION.SDK_INT >= 26 && !activity.getPackageManager().canRequestPackageInstalls()) {
            Toast.makeText(activity, "Allow \"Install unknown apps\" for this app, then tap Update again.",
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
}
