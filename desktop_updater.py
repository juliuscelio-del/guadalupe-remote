"""
Auto-updater for the Windows and Mac apps (checks GitHub Releases).

Called from Remote.py AFTER the Remote window is already open:
  1. a small dialog overlaps the Remote window: "searching for updates" (at least SEARCH_MS)
  2. no update  -> "you're up to date" for NO_UPDATE_MS, then the dialog closes (Remote stays open)
  3. update     -> "downloading", replaces the app, and restarts it (the Remote window is locked meanwhile)
  4. offline    -> the dialog just closes
It does nothing when Remote.py is run from source (only in the built app).
To preview the look from source (never downloads anything):  GNHS_PREVIEW_UPDATER=1 python3 Remote.py
"""
import os
import sys
import json
import time
import shutil
import tempfile
import threading
import subprocess
import urllib.request

VERSION = "0.0.0"    # replaced by GitHub Actions with 1.0.<build number>

API = "https://api.github.com/repos/juliuscelio-del/guadalupe-remote/releases/latest"

# ---- change the texts / timings here ----
MSG_SEARCHING = "search sa tah kadali ha ug naa ba updates"
MSG_NO_UPDATE = "Pwed pa ka mopadajun gamit ane kai Updated pa"
MSG_DOWNLOADING = "downloading sa ta ha sa updates...paabot lang"
MSG_FAILED = "Dili ma-download ang update. Sulayi pag usab sa ulahi."
SEARCH_MS = 10000      # 10 seconds
NO_UPDATE_MS = 5000    # 5 seconds
# -----------------------------------------

WIN = sys.platform == "win32"
MAC = sys.platform == "darwin"


def _num(s):
    digits = "".join(ch for ch in s if ch.isdigit())
    return int(digits) if digits else 0


def _is_newer(a, b):
    x, y = a.split("."), b.split(".")
    for i in range(max(len(x), len(y))):
        p = _num(x[i]) if i < len(x) else 0
        q = _num(y[i]) if i < len(y) else 0
        if p != q:
            return p > q
    return False


def _ssl_context():
    import ssl
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def _open(url, timeout):
    req = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github+json", "User-Agent": "GNHS-Remote"})
    return urllib.request.urlopen(req, timeout=timeout, context=_ssl_context())


def _fetch_latest():
    """-> ('none',) | ('update', version, url) | ('error',)"""
    try:
        with _open(API, 8) as r:
            data = json.load(r)
        latest = data["tag_name"].lstrip("vV")
        if not _is_newer(latest, VERSION):
            return ("none",)
        want = ".exe" if WIN else "-Mac.zip"
        for a in data.get("assets", []):
            if a["name"].endswith(want):
                return ("update", latest, a["browser_download_url"])
        return ("none",)
    except Exception:
        return ("error",)


def _clean_env():
    """Environment for relaunching the app (PyInstaller must start fresh)."""
    env = dict(os.environ)
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    for k in list(env):
        if k.startswith("_PYI") or k == "_MEIPASS2":
            env.pop(k, None)
    return env


# ------------------------------------------------------------ applying the update
def _apply_windows(new_exe):
    exe = sys.executable
    bat = os.path.join(tempfile.gettempdir(), "gnhs_update.bat")
    lines = [
        "@echo off",
        "set N=0",
        ":retry",
        "ping 127.0.0.1 -n 3 >nul",
        'move /y "' + new_exe + '" "' + exe + '" >nul 2>&1',
        "if not errorlevel 1 goto launch",
        "set /a N+=1",
        "if %N% lss 30 goto retry",
        ":launch",
        'start "" "' + exe + '"',
        '(goto) 2>nul & del "%~f0"',
    ]
    with open(bat, "w", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")
    # CREATE_NO_WINDOW | CREATE_NEW_PROCESS_GROUP
    subprocess.Popen(["cmd", "/c", bat], creationflags=0x08000000 | 0x00000200,
                     close_fds=True, env=_clean_env())


def _apply_mac(zip_path):
    exe = sys.executable
    if ".app/" not in exe:
        raise RuntimeError("not running from an .app bundle")
    old_app = exe.split(".app/")[0] + ".app"
    if not os.access(os.path.dirname(old_app), os.W_OK):
        raise RuntimeError("app folder is not writable")

    work = os.path.dirname(zip_path)
    subprocess.run(["ditto", "-x", "-k", zip_path, work], check=True)
    new_app = None
    for name in os.listdir(work):
        if name.endswith(".app"):
            new_app = os.path.join(work, name)
            break
    if not new_app:
        raise RuntimeError("no .app inside the update")

    sh = os.path.join(work, "swap.sh")
    script = "\n".join([
        "#!/bin/sh",
        'OLD="' + old_app + '"',
        'NEW="' + new_app + '"',
        'BAK="' + old_app + '.old"',
        "while kill -0 " + str(os.getpid()) + " 2>/dev/null; do sleep 0.5; done",
        'rm -rf "$BAK"',
        'mv "$OLD" "$BAK" || { open "$OLD"; exit 1; }',
        'if mv "$NEW" "$OLD"; then rm -rf "$BAK"; else mv "$BAK" "$OLD"; fi',
        'xattr -dr com.apple.quarantine "$OLD" 2>/dev/null',
        'open "$OLD"',
        "",
    ])
    with open(sh, "w") as f:
        f.write(script)
    devnull = open(os.devnull, "w")
    subprocess.Popen(["/bin/sh", sh], start_new_session=True, env=_clean_env(),
                     stdout=devnull, stderr=devnull)


# ------------------------------------------------------------ the dialog
# vibrant palette (same family as the phone remote)
C_CARD = "#ffffff"
C_INK = "#12306b"
C_TRACK = "#dbe8ff"
STRIPE = ["#2f6df6", "#4aa3ff", "#1fc76a", "#ffa31a", "#8b5cf6"]
BAR = {                      # three overlapping pills = soft gradient
    "search": ("#4aa3ff", "#2f6df6", "#8b5cf6"),
    "download": ("#ffa31a", "#ff7a3d", "#ff5a5f"),
    "none": ("#1fc76a", "#1fc76a", "#1fc76a"),
    "fail": ("#ff5a5f", "#ff5a5f", "#ff5a5f"),
}
TEXT = {"search": C_INK, "download": "#d98200", "none": "#139a50", "fail": "#d63a40"}


def check(parent):
    """Show the update dialog ON TOP of the already-open Remote window."""
    preview = bool(os.environ.get("GNHS_PREVIEW_UPDATER"))   # layout test from source
    if not getattr(sys, "frozen", False) and not preview:
        return                      # running from source: no updates
    if not (WIN or MAC or preview):
        return
    try:
        _run(parent, preview)
    except Exception:
        pass


def _run(parent, preview=False):
    import tkinter as tk

    W, H = 400, 118
    dlg = tk.Toplevel(parent)
    dlg.title("GNHS Remote")
    dlg.configure(bg=C_CARD)
    dlg.resizable(False, False)

    cv = tk.Canvas(dlg, width=W, height=H, bg=C_CARD, highlightthickness=0, bd=0)
    cv.pack()

    # colorful top stripe
    seg = W / len(STRIPE)
    for i, col in enumerate(STRIPE):
        cv.create_rectangle(i * seg, 0, (i + 1) * seg + 1, 5, fill=col, outline=col)

    msg_id = cv.create_text(W / 2, 48, text=MSG_SEARCHING, width=W - 56, justify="center",
                            font=("Helvetica", 14, "bold"), fill=TEXT["search"])

    # progress pill
    BX1, BX2, BY = 28, W - 28, 92
    cv.create_line(BX1, BY, BX2, BY, width=8, capstyle="round", fill=C_TRACK)
    bar_ids = [cv.create_line(0, BY, 0, BY, width=8, capstyle="round", fill="#fff") for _ in range(3)]

    # center it on the SCREEN (same on Mac and Windows); Remote.py centers its window too
    dlg.update_idletasks()
    x = (dlg.winfo_screenwidth() - W) // 2
    y = (dlg.winfo_screenheight() - H) // 2
    dlg.geometry("%dx%d+%d+%d" % (W, H, max(x, 0), max(y, 0)))
    dlg.transient(parent)                         # stays above the Remote window
    dlg.attributes("-topmost", True)
    dlg.protocol("WM_DELETE_WINDOW", lambda: None)   # can't be closed while working
    dlg.lift()
    try:
        dlg.wait_visibility()
        dlg.grab_set()                            # Remote window is locked = forced update
    except Exception:
        pass

    st = {"stage": "search", "result": None, "dl": None, "path": None,
          "t0": time.time(), "pos": 0.0, "dir": 1, "alive": True}

    def set_stage(stage, text):
        st["stage"] = stage
        cv.itemconfig(msg_id, text=text, fill=TEXT.get(stage, C_INK))
        for item, col in zip(bar_ids, BAR.get(stage, BAR["search"])):
            cv.itemconfig(item, fill=col)

    for item, col in zip(bar_ids, BAR["search"]):
        cv.itemconfig(item, fill=col)

    def animate():
        if not st["alive"]:
            return
        try:
            if st["stage"] in ("none", "fail"):
                # full solid bar
                for item in bar_ids:
                    cv.coords(item, BX1, BY, BX2, BY)
            else:
                # sliding block that bounces left <-> right
                span = (BX2 - BX1) - 110
                st["pos"] += 4.5 * st["dir"]
                if st["pos"] >= span:
                    st["pos"], st["dir"] = span, -1
                elif st["pos"] <= 0:
                    st["pos"], st["dir"] = 0, 1
                a = BX1 + st["pos"]
                b = a + 110
                cv.coords(bar_ids[0], a, BY, b, BY)
                cv.coords(bar_ids[1], a + 28, BY, b, BY)
                cv.coords(bar_ids[2], a + 62, BY, b, BY)
            dlg.after(16, animate)
        except Exception:
            pass

    def find():
        st["result"] = ("none",) if preview else _fetch_latest()

    def download(url):
        try:
            if WIN:
                dest = sys.executable + ".new"
            else:
                dest = os.path.join(tempfile.mkdtemp(), "update.zip")
            with _open(url, 30) as r, open(dest, "wb") as f:
                shutil.copyfileobj(r, f)
            if os.path.getsize(dest) < 1000000:
                raise RuntimeError("download too small")
            st["path"] = dest
            st["dl"] = "ok"
        except Exception:
            st["dl"] = "fail"

    def close():
        st["alive"] = False
        try:
            dlg.grab_release()
            dlg.destroy()                         # only the dialog; Remote keeps running
        except Exception:
            pass

    def fail():
        set_stage("fail", MSG_FAILED)
        dlg.after(4000, close)

    def tick():
        try:
            stage = st["stage"]
            if stage == "search":
                waited = (time.time() - st["t0"]) * 1000
                r = st["result"]
                if r is not None and waited >= SEARCH_MS:
                    if r[0] == "none":
                        set_stage("none", MSG_NO_UPDATE)
                        dlg.after(NO_UPDATE_MS, close)
                        return
                    if r[0] == "error":
                        close()
                        return
                    set_stage("download", MSG_DOWNLOADING)
                    threading.Thread(target=download, args=(r[2],), daemon=True).start()
            elif stage == "download" and st["dl"] is not None:
                if st["dl"] != "ok":
                    fail()
                    return
                try:
                    if WIN:
                        _apply_windows(st["path"])
                    else:
                        _apply_mac(st["path"])
                except Exception:
                    fail()
                    return
                os._exit(0)         # the helper script swaps the files and restarts the app
            dlg.after(200, tick)
        except Exception:
            pass                    # Remote window was closed meanwhile

    threading.Thread(target=find, daemon=True).start()
    animate()
    dlg.after(200, tick)
