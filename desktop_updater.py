"""
Auto-updater for the Windows and Mac apps (checks GitHub Releases).

Called from Remote.py at start-up:
  1. shows "searching for updates" for at least SEARCH_MS
  2. no update  -> shows "you're up to date" for NO_UPDATE_MS, then the app opens
  3. update     -> shows "downloading", replaces the app, and restarts it
  4. offline    -> the app just opens
It does nothing when Remote.py is run from source (only in the built app).
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
def check():
    if not getattr(sys, "frozen", False):
        return                      # running from source: no updates
    if not (WIN or MAC):
        return
    try:
        _run()
    except Exception:
        pass


def _run():
    import tkinter as tk

    BG, INK = "#e4f1ff", "#12306b"
    root = tk.Tk()
    root.title("GNHS Remote")
    root.configure(bg=BG)
    root.resizable(False, False)
    lbl = tk.Label(root, text=MSG_SEARCHING, font=("Helvetica", 15, "bold"), fg=INK, bg=BG,
                   wraplength=380, justify="center", padx=30, pady=40)
    lbl.pack()
    root.update_idletasks()
    w, h = 440, 150
    root.geometry("%dx%d+%d+%d" % (w, h, (root.winfo_screenwidth() - w) // 2,
                                   (root.winfo_screenheight() - h) // 3))
    root.attributes("-topmost", True)
    root.protocol("WM_DELETE_WINDOW", lambda: None)   # can't be closed while working

    st = {"stage": "search", "result": None, "dl": None, "path": None, "t0": time.time()}

    def find():
        st["result"] = _fetch_latest()

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

    def fail():
        st["stage"] = "fail"
        lbl.config(text=MSG_FAILED)
        root.after(4000, root.destroy)

    def tick():
        stage = st["stage"]
        if stage == "search":
            waited = (time.time() - st["t0"]) * 1000
            r = st["result"]
            if r is not None and waited >= SEARCH_MS:
                if r[0] == "none":
                    st["stage"] = "none"
                    lbl.config(text=MSG_NO_UPDATE)
                    root.after(NO_UPDATE_MS, root.destroy)
                    return
                if r[0] == "error":
                    root.destroy()
                    return
                st["stage"] = "download"
                lbl.config(text=MSG_DOWNLOADING)
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
            os._exit(0)             # the helper script swaps the files and restarts the app
        root.after(200, tick)

    threading.Thread(target=find, daemon=True).start()
    root.after(200, tick)
    root.mainloop()
