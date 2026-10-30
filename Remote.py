"""
PowerPoint remote + remote desktop: run on the LAPTOP, open the printed URL on the phone.

Install:  pip3 install flask pyautogui mss pyperclip pillow qrcode     (+ pywin32 on Windows for the PowerPoint link)
Run:      python3 Remote.py                        (a window with a QR code, the address and a PIN opens)

Optional: set the STARTING permissions (the phone can change them with checkboxes):
  python3 ppt_remote_desktop.py ppt          PowerPoint remote only
  python3 ppt_remote_desktop.py view         watch the desktop only (no control)
  python3 ppt_remote_desktop.py mouse        mouse only (works as a touchpad)
  python3 ppt_remote_desktop.py keys         keyboard only
  python3 ppt_remote_desktop.py ppt,view     combine with commas
  python3 ppt_remote_desktop.py all          everything
(or set RC_ALLOW=ppt,view). With no option, the phone starts with PowerPoint only.
Anything not allowed is refused by the laptop itself.

Phone: scan the QR code with the camera (it opens the browser), type the PIN shown in the window,
tick what you want to allow, tap Continue.
The main page is the PowerPoint remote; the "Permissions" button reopens the checkboxes.
Tap "Desktop control" on it to see and control the screen.

Mac permissions (System Settings > Privacy & Security), then restart the script:
  - Accessibility      -> turn on Terminal (needed for mouse and keyboard)
  - Screen Recording   -> turn on Terminal (needed for the live screen view)
"""
import io, os, sys, json, time, socket, secrets, tempfile, threading
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")
import pyautogui
from flask import Flask, request, session, redirect, jsonify, send_file, abort, Response

try:
    import win32com.client as w32
    import pythoncom
except ImportError:
    w32 = None

try:
    import mss
except ImportError:
    mss = None

try:
    import pyperclip
except ImportError:
    pyperclip = None

from PIL import Image

SCHOOL = "Guadalupe NHS"
try:
    import base64
    from logo_data import LOGO_B64          # embedded school logo (logo_data.py)
    LOGO_PNG = base64.b64decode(LOGO_B64)
except Exception:
    LOGO_PNG = None

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0
PORT = int(os.environ.get("RC_PORT", 5001))   # 5001 because macOS uses 5000 for AirPlay
PIN = os.environ.get("RC_PIN") or f"{secrets.randbelow(10**6):06d}"
MAC = sys.platform == "darwin"
MOD = "command" if MAC else "ctrl"

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)
LOCK = threading.Lock()    # PowerPoint + slideshow input
DLOCK = threading.Lock()   # desktop mouse/keyboard input
SLOCK = threading.Lock()   # screen capture
THUMBS = {}
DRAG = False


# ------------------------------------------------------------ permissions
CAP_INFO = [("ppt", "PowerPoint remote (slides, notes, laser/pen)"),
            ("view", "View the desktop (live screen)"),
            ("mouse", "Control the mouse"),
            ("keys", "Use the keyboard")]
ALIASES = {"powerpoint": "ppt", "slides": "ppt", "screen": "view", "desktop": "view",
           "pointer": "mouse", "keyboard": "keys", "key": "keys", "type": "keys"}
CAPS = set()
MOUSE_ACTIONS = {"moveto", "rel", "click", "dbl", "down", "up", "scroll"}
KEY_ACTIONS = {"key", "text"}


def parse_caps(text):
    """'ppt,view' / '1 3' / 'all' -> set of capabilities, or None if invalid."""
    names = [c for c, _ in CAP_INFO]
    out = set()
    for tok in text.replace(",", " ").replace("+", " ").lower().split():
        if tok in ("all", "everything", "full"):
            out |= set(names)
        elif tok.isdigit() and 1 <= int(tok) <= len(names):
            out.add(names[int(tok) - 1])
        elif tok in names:
            out.add(tok)
        elif tok in ALIASES:
            out.add(ALIASES[tok])
        else:
            return None
    return out or None


def choose_caps():
    """Starting permissions. The phone can change them later on /settings."""
    arg = " ".join(sys.argv[1:]) or os.environ.get("RC_ALLOW", "")
    if arg:
        caps = parse_caps(arg)
        if caps:
            return caps
        sys.exit(f"Unknown option: {arg}\nUse any of: ppt, view, mouse, keys, all  (e.g. ppt,view)")
    return {"ppt"}   # safe default: only the PowerPoint remote


def need(cap):
    if cap not in CAPS:
        abort(Response("not allowed in this mode", status=423))


def page(html):
    return (html.replace("__THEME__", THEME).replace("__BRAND__", BRAND)
                .replace("__SCHOOL__", SCHOOL)
                .replace("__CAPS__", json.dumps({c: c in CAPS for c, _ in CAP_INFO})))


# ------------------------------------------------------------ PowerPoint
def get_app():
    if not w32:
        return None
    try:
        pythoncom.CoInitialize()
        return w32.GetActiveObject("PowerPoint.Application")
    except Exception:
        return None


def get_show(a):
    try:
        if a is not None and a.SlideShowWindows.Count > 0:
            return a.SlideShowWindows(1)
    except Exception:
        pass
    return None


KEYS = {"next": ["right"], "prev": ["left"], "first": ["home"], "last": ["end"],
        "start": ["f5"], "startcur": ["shift", "f5"], "end": ["esc"],
        "black": ["b"], "white": ["w"], "pen": ["ctrl", "p"],
        "arrow": ["ctrl", "a"], "clear": ["e"]}


def press(a):
    k = KEYS.get(a)
    if k:
        pyautogui.hotkey(*k)


def act(d):
    a = d["a"]
    if a == "move":
        pyautogui.moveRel(d["dx"], d["dy"], _pause=False)
        return
    if a == "press":  # laser (ctrl+hold) or pen (plain hold)
        if d.get("ctrl"):
            pyautogui.keyDown("ctrl")
        pyautogui.mouseDown()
        return
    if a == "release":
        pyautogui.mouseUp()
        pyautogui.keyUp("ctrl")
        return

    pp = get_app()
    if pp is None:  # no COM: fall back to plain keystrokes
        if a == "goto":
            pyautogui.write(str(int(d["n"])))
            pyautogui.press("enter")
        else:
            press(a)
        return

    try:
        pres = pp.ActivePresentation
        if a in ("start", "startcur"):
            s = pres.SlideShowSettings
            if a == "startcur":
                s.StartingSlide = pp.ActiveWindow.View.Slide.SlideIndex
                s.EndingSlide = pres.Slides.Count
                s.RangeType = 2
            else:
                s.RangeType = 1
            THUMBS.clear()
            s.Run()
            return
        sw = get_show(pp)
        if sw is None:
            return  # no show running: ignore navigation
        v = sw.View
        if a == "next": v.Next()
        elif a == "prev": v.Previous()
        elif a == "first": v.First()
        elif a == "last": v.Last()
        elif a == "goto": v.GotoSlide(int(d["n"]))
        elif a == "end": v.Exit()
        elif a == "black": v.State = 1 if v.State == 3 else 3
        elif a == "white": v.State = 1 if v.State == 4 else 4
        elif a == "pen": v.PointerType = 2
        elif a == "arrow": v.PointerType = 1
        elif a == "clear": v.EraseDrawing()
    except Exception as e:
        print("PowerPoint error:", e)


def status():
    pp = get_app()
    out = {"com": pp is not None, "running": False, "cur": 0, "total": 0, "notes": "", "state": 1}
    if pp is None:
        return out
    try:
        sw = get_show(pp)
        out["total"] = pp.ActivePresentation.Slides.Count
        if sw is None:
            return out
        v = sw.View
        out.update(running=True, cur=v.CurrentShowPosition, state=v.State)
        try:
            out["notes"] = v.Slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        except Exception:
            pass
    except Exception as e:
        print("status error:", e)
    return out


def thumb(n):
    if n in THUMBS:
        return THUMBS[n]
    pp = get_app()
    if pp is None:
        return None
    pres = pp.ActivePresentation
    if not 1 <= n <= pres.Slides.Count:
        return None
    w = 640
    h = int(w * pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth)
    path = os.path.join(tempfile.gettempdir(), f"pptr_{secrets.token_hex(4)}.jpg")
    pres.Slides(n).Export(path, "JPG", w, h)
    with open(path, "rb") as f:
        data = f.read()
    os.remove(path)
    THUMBS[n] = data
    return data


# ------------------------------------------------------------ desktop control
DKEYS = {
    "enter": ["enter"], "bs": ["backspace"], "esc": ["esc"], "tab": ["tab"], "space": ["space"],
    "up": ["up"], "down": ["down"], "left": ["left"], "right": ["right"],
    "copy": [MOD, "c"], "paste": [MOD, "v"], "all": [MOD, "a"], "undo": [MOD, "z"],
    "switch": ["command", "tab"] if MAC else ["alt", "tab"],
    "spot": ["command", "space"] if MAC else ["win"],
}


def grab(width, quality):
    """Screenshot of the primary screen as JPEG bytes, scaled to `width` px wide."""
    if mss:
        with mss.mss() as sct:
            raw = sct.grab(sct.monitors[1])
            img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")
    else:
        img = pyautogui.screenshot().convert("RGB")
    if img.width > width:
        img = img.resize((width, int(img.height * width / img.width)), Image.BILINEAR)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=quality)
    return buf.getvalue()


def dact(d):
    global DRAG
    a = d["a"]
    W, H = pyautogui.size()
    x = y = None
    if "fx" in d:
        x = min(max(float(d["fx"]), 0.0), 1.0) * (W - 1)
        y = min(max(float(d["fy"]), 0.0), 1.0) * (H - 1)

    if a == "rel":  # relative move (touchpad mode, used when the screen view is off)
        px, py = pyautogui.position()
        x = min(max(px + float(d["dx"]), 0), W - 1)
        y = min(max(py + float(d["dy"]), 0), H - 1)
        a = "moveto"

    if a == "moveto":
        if DRAG:
            try:
                pyautogui.dragTo(x, y, button="left", mouseDownUp=False, _pause=False)
            except TypeError:
                pyautogui.moveTo(x, y, _pause=False)
        else:
            pyautogui.moveTo(x, y, _pause=False)
    elif a == "click":
        pyautogui.click(x, y, button=d.get("b", "left"), _pause=False)
    elif a == "dbl":
        pyautogui.doubleClick(x, y, _pause=False)
    elif a == "down":
        if x is not None:
            pyautogui.moveTo(x, y, _pause=False)
        pyautogui.mouseDown(button="left", _pause=False)
        DRAG = True
    elif a == "up":
        pyautogui.mouseUp(button="left", _pause=False)
        DRAG = False
    elif a == "scroll":
        n = int(d["n"])
        if n:
            if x is not None:
                pyautogui.moveTo(x, y, _pause=False)
            pyautogui.scroll(n, _pause=False)
    elif a == "key":
        k = DKEYS.get(d.get("k"))
        if k:
            pyautogui.hotkey(*k)
    elif a == "text":
        t = d.get("t", "")
        if t:
            if pyperclip:
                try:
                    old = pyperclip.paste()
                except Exception:
                    old = None
                pyperclip.copy(t)
                time.sleep(0.05)
                pyautogui.hotkey(MOD, "v")
                time.sleep(0.25)
                if old is not None:
                    pyperclip.copy(old)
            else:
                pyautogui.write(t, interval=0.01)  # ASCII only
        if d.get("enter"):
            pyautogui.press("enter")


# ------------------------------------------------------------ web
THEME = r"""
:root{--bg1:#e4f1ff;--bg2:#fff3d1;--ink:#12306b;--muted:#5f78a8;--line:#c6dbff;
--blue:#2f6df6;--sky:#4aa3ff;--green:#1fc76a;--red:#ff5a5f;--purple:#8b5cf6;--orange:#ffa31a}
*{box-sizing:border-box}
html,body{margin:0;color:var(--ink);font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
 background:linear-gradient(160deg,var(--bg1),var(--bg2)) fixed;overscroll-behavior:none;
 -webkit-user-select:none;user-select:none}
.brand{display:flex;align-items:center;gap:10px}
.brand img{width:46px;height:46px;border-radius:50%;box-shadow:0 2px 8px rgba(47,109,246,.35)}
.brand b{display:block;font-size:1.2rem;color:var(--blue);line-height:1.1}
.brand span{font-size:.8rem;color:var(--muted)}
.brand.small img{width:28px;height:28px}
.brand.small b{font-size:1rem}
button{font-family:inherit;font-weight:600;color:var(--ink);background:#fff;border:2px solid var(--line);
 border-radius:14px;box-shadow:0 2px 0 var(--line)}
button:active{transform:translateY(1px);box-shadow:none}
button.on{background:var(--blue);border-color:var(--blue);color:#fff;box-shadow:0 2px 0 #1d4fc4}
button.go{background:var(--green);border-color:var(--green);color:#fff;box-shadow:0 3px 0 #139a50}
button.stop{background:var(--red);border-color:var(--red);color:#fff;box-shadow:0 3px 0 #d63a40}
button.purple{background:var(--purple);border-color:var(--purple);color:#fff;box-shadow:0 3px 0 #6d3fd1}
"""

BRAND = ('<div class=brand><img src="/logo.png" alt="" onerror="this.style.display=\'none\'">'
         '<div><b>' + SCHOOL + '</b><span>PPT Remote</span></div></div>')

LOGIN = """<!doctype html><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>__SCHOOL__ - PPT Remote</title>
<link rel=icon href=/logo.png>
<style>
__THEME__
body{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px}
.card{background:#fff;border-radius:26px;padding:28px 22px;width:100%;max-width:340px;text-align:center;
 box-shadow:0 12px 32px rgba(47,109,246,.2)}
.card img{width:112px;height:112px;border-radius:50%;box-shadow:0 4px 14px rgba(47,109,246,.4)}
h2{margin:12px 0 0;color:var(--blue);font-size:1.6rem}
p.sub{margin:2px 0 18px;color:var(--muted)}
input{font-size:1.6rem;padding:.6rem;width:100%;text-align:center;letter-spacing:.3em;
 border:2px solid var(--line);border-radius:14px;color:var(--ink);background:#f4f9ff;
 -webkit-user-select:text;user-select:text}
button.go{margin-top:14px;width:100%;font-size:1.2rem;padding:.8rem}
.err{color:var(--red);min-height:1.2em;margin:10px 0 0;font-weight:600}
</style>
<div class=card><img src="/logo.png" alt="" onerror="this.style.display='none'">
<h2>__SCHOOL__</h2><p class=sub>PPT Remote</p>
<form method=post><input name=pin type=password inputmode=numeric placeholder="PIN" autofocus autocomplete=off>
<button class=go>Connect</button><p class=err>{err}</p></form></div>""".replace("__THEME__", THEME).replace("__SCHOOL__", SCHOOL)

PAGE = r"""<!doctype html>
<meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1,user-scalable=no">
<title>__SCHOOL__ - PPT Remote</title>
<link rel=icon href=/logo.png>
<style>
__THEME__
body{padding:10px;display:flex;flex-direction:column;gap:10px;min-height:100vh}
.row{display:flex;gap:8px}
.row>*{flex:1}
.top{display:flex;justify-content:space-between;align-items:center;gap:8px}
.top button{flex:0 0 auto;padding:.5rem .8rem;font-size:.85rem;background:var(--orange);
 border-color:var(--orange);color:#fff;box-shadow:0 2px 0 #d98200}
button,input{font-size:1rem;padding:.8rem .4rem}
input{border:2px solid var(--line);border-radius:12px;background:#fff;color:var(--ink);
 -webkit-user-select:text;user-select:text}
button.big{font-size:2rem;padding:1.5rem 0;background:var(--blue);border-color:var(--blue);
 color:#fff;box-shadow:0 4px 0 #1d4fc4}
button.big.prev{background:#fff;color:var(--blue);border-color:var(--blue);box-shadow:0 4px 0 var(--line)}
#info{display:flex;justify-content:space-between;align-items:center;font-size:1.4rem;font-weight:800}
#pos{color:var(--blue)} #tmr{color:var(--orange)}
.thumbs{display:flex;gap:8px}
.thumbs div{flex:1;font-size:.75rem;color:var(--muted);font-weight:600}
.thumbs img{width:100%;border-radius:10px;background:#dbe8ff;display:block;min-height:50px;
 border:2px solid #fff;box-shadow:0 2px 8px rgba(47,109,246,.2)}
#notes{background:#fff;border:2px solid var(--line);border-radius:14px;padding:10px;height:20vh;
 overflow:auto;white-space:pre-wrap;font-size:1.1rem;-webkit-user-select:text;user-select:text}
#padbox{touch-action:none;background:#fff;border:3px dashed var(--sky);border-radius:14px;height:18vh;
 display:flex;align-items:center;justify-content:center;color:var(--muted);text-align:center;font-weight:600}
#msg{color:#d98200;font-size:.85rem;font-weight:600}
</style>
<div class=top>__BRAND__<button onclick="location='/settings'">⚙ Permissions</button></div>
<div id=info><span id=pos>– / –</span><span id=tmr onclick="t0=null">00:00</span></div>
<div id=msg></div>
<div class=row id=deskrow><button class=purple onclick="location='/desktop'">🖥 Desktop control</button></div>
<div class=thumbs>
 <div>Now<img id=t1></div><div>Next<img id=t2></div>
</div>
<div id=notes>Notes appear here</div>
<div class=row>
 <button class="big prev" onclick="api({a:'prev'})">◀ PREV</button>
 <button class=big onclick="api({a:'next'})">NEXT ▶</button>
</div>
<div class=row>
 <button class=go onclick="api({a:'start'})">▶ Start</button>
 <button class=go onclick="api({a:'startcur'})">▶ From current</button>
 <button class=stop onclick="api({a:'end'})">■ End</button>
</div>
<div class=row>
 <button id=bk onclick="api({a:'black'})">Black</button>
 <button id=wh onclick="api({a:'white'})">White</button>
 <button onclick="api({a:'first'})">|◀ First</button>
 <button onclick="api({a:'last'})">Last ▶|</button>
</div>
<div class=row>
 <input id=gn type=number inputmode=numeric placeholder="Slide #">
 <button onclick="if(gn.value)api({a:'goto',n:gn.value})">Go to</button>
</div>
<div class=row id=tools>
 <button data-tool=laser class=on>Laser</button>
 <button data-tool=pen>Pen</button>
 <button data-tool=arrow>Mouse</button>
 <button onclick="api({a:'clear'})">Clear ink</button>
</div>
<div id=padbox>Touch &amp; drag here<br>(laser / pen / mouse)</div>
<script>
const $=id=>document.getElementById(id);
const api=o=>fetch('/api',{method:'POST',headers:{'Content-Type':'application/json'},
  body:JSON.stringify(o)}).catch(()=>{});
const CAPS=__CAPS__;
if(!(CAPS.view||CAPS.mouse||CAPS.keys)) $('deskrow').style.display='none';
let tool='laser', t0=null, last={cur:-1};
try{navigator.wakeLock&&navigator.wakeLock.request('screen');}catch(e){}
document.addEventListener('click',()=>{try{navigator.wakeLock&&navigator.wakeLock.request('screen');}catch(e){}});

$('tools').addEventListener('click',e=>{
  const b=e.target.closest('[data-tool]'); if(!b) return;
  tool=b.dataset.tool;
  document.querySelectorAll('[data-tool]').forEach(x=>x.classList.toggle('on',x===b));
  api({a: tool==='pen'?'pen':'arrow'});
});

async function poll(){
  try{
    const s=await (await fetch('/status')).json();
    $('msg').textContent = s.com ? (s.running?'':'Slideshow not running. Press Start.')
        : 'Keyboard mode (no PowerPoint link): the slideshow window must be in front.';
    $('pos').textContent = s.running ? `${s.cur} / ${s.total}` : (s.total?`– / ${s.total}`:'– / –');
    $('bk').classList.toggle('on',s.state===3); $('wh').classList.toggle('on',s.state===4);
    if(s.running && !t0) t0=Date.now();
    if(!s.running && s.com) t0=null;
    if(s.cur!==last.cur){
      if(s.running){
        $('t1').src='/thumb/'+s.cur; $('t2').src='/thumb/'+(s.cur+1);
        $('t2').style.visibility = s.cur<s.total ? 'visible':'hidden';
      }
      $('notes').textContent = s.notes || '(no notes)'; $('notes').scrollTop=0;
    }
    last=s;
  }catch(e){}
  setTimeout(poll,600);
}
if(CAPS.ppt) poll();
setInterval(()=>{
  const el=t0?Math.floor((Date.now()-t0)/1000):0;
  $('tmr').textContent=String(Math.floor(el/60)).padStart(2,'0')+':'+String(el%60).padStart(2,'0');
},500);

// laser / pen pad
const pad=$('padbox'); let lx,ly,ax=0,ay=0,iv=null;
pad.addEventListener('touchstart',e=>{
  e.preventDefault(); const p=e.touches[0]; lx=p.clientX; ly=p.clientY;
  if(tool!=='arrow') api({a:'press',ctrl:tool==='laser'});
  iv=setInterval(()=>{ if(ax||ay){api({a:'move',dx:ax,dy:ay}); ax=ay=0;} },30);
},{passive:false});
pad.addEventListener('touchmove',e=>{
  e.preventDefault(); const p=e.touches[0];
  ax+=(p.clientX-lx)*3; ay+=(p.clientY-ly)*3; lx=p.clientX; ly=p.clientY;
},{passive:false});
const end=e=>{ e.preventDefault(); clearInterval(iv);
  if(ax||ay){api({a:'move',dx:ax,dy:ay}); ax=ay=0;}
  if(tool!=='arrow') api({a:'release'}); };
pad.addEventListener('touchend',end,{passive:false});
pad.addEventListener('touchcancel',end,{passive:false});
</script>
"""

DESKTOP = r"""<!doctype html>
<meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1,user-scalable=no,viewport-fit=cover">
<title>__SCHOOL__ - Desktop</title>
<link rel=icon href=/logo.png>
<style>
__THEME__
html,body{height:100%}
body{display:flex;flex-direction:column;gap:6px;padding:6px;height:100dvh}
.row{display:flex;gap:6px}
.row>*{flex:1;min-width:0}
button,input,select{font-size:.95rem;padding:.6rem .3rem}
select{border:2px solid var(--line);border-radius:12px;background:#fff;color:var(--ink);font-weight:600}
input{font-size:16px;border:2px solid var(--line);border-radius:12px;background:#fff;color:var(--ink);
 -webkit-user-select:text;user-select:text}
#view{flex:1;min-height:0;display:flex;align-items:center;justify-content:center;background:#dbe8ff;
 border-radius:12px;overflow:hidden;border:2px solid #fff;box-shadow:0 2px 10px rgba(47,109,246,.25)}
#wrap{position:relative;line-height:0;touch-action:none}
#scr{display:block;width:100%;height:100%}
#hint{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;
 text-align:center;padding:12px;color:var(--muted);line-height:1.3;pointer-events:none;font-size:.9rem;font-weight:600}
#wrap.noview{background:#fff;border:3px dashed var(--sky);border-radius:14px}
#dot{position:absolute;width:14px;height:14px;margin:-7px 0 0 -7px;border:2px solid #fff;
 background:rgba(255,60,60,.8);border-radius:50%;pointer-events:none;box-shadow:0 0 0 1px #000;
 left:50%;top:50%}
#msg{color:#d98200;font-size:.75rem;min-height:1em;font-weight:600}
#keys{display:none;flex-direction:column;gap:6px}
#keys.show{display:flex}
</style>
<div class="brand small"><img src="/logo.png" alt="" onerror="this.style.display='none'"><b>__SCHOOL__</b></div>
<div class=row>
 <button id=back onclick="location='/'">◀ Remote</button>
 <select id=qual>
  <option value=fast>Fast</option>
  <option value=bal selected>Balanced</option>
  <option value=sharp>Sharp</option>
 </select>
 <button id=kb>⌨ Keys</button>
 <button onclick="location='/settings'">⚙</button>
</div>
<div id=msg>Tap = click · hold = right-click · drag = move · 2 fingers = scroll</div>
<div id=view>
 <div id=wrap><img id=scr alt=""><div id=dot></div><div id=hint></div></div>
</div>
<div class=row id=modes>
 <button data-mode=ptr class=on>Pointer</button>
 <button data-mode=drag>Drag</button>
 <button onclick="send({a:'click',b:'left'})">Click</button>
 <button onclick="send({a:'click',b:'right'})">Right</button>
 <button onclick="send({a:'dbl'})">Double</button>
</div>
<div id=keys>
 <div class=row>
  <input id=txt placeholder="Type text to send…" autocomplete=off autocapitalize=off>
  <button style="flex:0 0 4.5em" onclick="sendText(false)">Send</button>
  <button style="flex:0 0 4.5em" onclick="sendText(true)">Send ⏎</button>
 </div>
 <div class=row id=krow>
  <button data-k=enter>⏎</button><button data-k=bs>⌫</button><button data-k=esc>Esc</button>
  <button data-k=tab>Tab</button><button data-k=space>Space</button>
 </div>
 <div class=row>
  <button data-k=left>←</button><button data-k=up>↑</button>
  <button data-k=down>↓</button><button data-k=right>→</button>
 </div>
 <div class=row>
  <button data-k=copy>Copy</button><button data-k=paste>Paste</button>
  <button data-k=all>All</button><button data-k=undo>Undo</button>
  <button data-k=switch>Switch app</button><button data-k=spot>Search</button>
 </div>
</div>
<script>
const $=id=>document.getElementById(id);
const wrap=$('wrap'), img=$('scr'), dot=$('dot'), view=$('view');
let mode='ptr';

/* ---- what this session is allowed to do ---- */
const CAPS=__CAPS__;
const REL=!CAPS.view;   // no screen to point at -> behave like a touchpad
if(!CAPS.ppt) $('back').style.display='none';
if(!CAPS.view){ $('qual').style.display='none'; img.style.display='none'; dot.style.display='none'; wrap.classList.add('noview'); }
if(!CAPS.mouse) $('modes').style.display='none';
if(!CAPS.keys) $('kb').style.display='none';
if(!CAPS.view && !CAPS.mouse){ view.style.display='none'; if(CAPS.keys){ $('keys').classList.add('show'); } }
$('hint').textContent = CAPS.view ? '' : 'Touchpad: drag = move · tap = click · hold = right-click · 2 fingers = scroll';
$('msg').textContent = CAPS.mouse ? (CAPS.view ? 'Tap = click · hold = right-click · drag = move · 2 fingers = scroll' : '')
  : (CAPS.view ? 'View only' : 'Keyboard only');

/* ---- ordered send queue (move events are coalesced) ---- */
const Q=[]; let pumping=false;
function send(o){
  const last=Q[Q.length-1];
  if(o.a==='moveto' && last && last.a==='moveto') Q[Q.length-1]=o;
  else if(o.a==='rel' && last && last.a==='rel'){ last.dx+=o.dx; last.dy+=o.dy; }
  else Q.push(o);
  pump();
}
async function pump(){
  if(pumping) return; pumping=true;
  while(Q.length){
    const o=Q.shift();
    try{
      const r=await fetch('/dapi',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(o)});
      if(r.status===403){location='/login';return;}
    }catch(e){}
  }
  pumping=false;
}

/* ---- live screen ---- */
const PRESETS={fast:[700,40],bal:[1000,55],sharp:[1600,72]};
function fit(){
  const ar=CAPS.view ? (img.naturalWidth ? img.naturalWidth/img.naturalHeight : 0) : 1.6;
  if(!ar) return;
  let w=view.clientWidth, h=w/ar;
  if(h>view.clientHeight){h=view.clientHeight; w=h*ar;}
  wrap.style.width=w+'px'; wrap.style.height=h+'px';
}
fit();
img.addEventListener('load',fit);
window.addEventListener('resize',fit);
if(window.ResizeObserver) new ResizeObserver(fit).observe(view);

async function frame(){
  if(document.hidden){setTimeout(frame,500);return;}
  try{
    const [w,q]=PRESETS[$('qual').value];
    const r=await fetch(`/screen?w=${w}&q=${q}&t=${Date.now()}`);
    if(r.status===403){location='/login';return;}
    if(!r.ok){ $('msg').textContent='Screen capture failed. Check Screen Recording permission.'; setTimeout(frame,1500); return; }
    const mp=r.headers.get('X-Mouse');
    const url=URL.createObjectURL(await r.blob());
    const old=img.src; img.src=url;
    if(old.startsWith('blob:')) URL.revokeObjectURL(old);
    if(mp){ const [x,y]=mp.split(',').map(Number); dot.style.left=(x*100)+'%'; dot.style.top=(y*100)+'%'; }
  }catch(e){ await new Promise(r=>setTimeout(r,500)); }
  setTimeout(frame,30);
}
if(CAPS.view) frame();

/* ---- buttons ---- */
$('modes').addEventListener('click',e=>{
  const b=e.target.closest('[data-mode]'); if(!b) return;
  mode=b.dataset.mode;
  document.querySelectorAll('[data-mode]').forEach(x=>x.classList.toggle('on',x===b));
});
$('kb').addEventListener('click',()=>{ $('keys').classList.toggle('show'); $('kb').classList.toggle('on'); });
$('keys').addEventListener('click',e=>{
  const b=e.target.closest('[data-k]'); if(!b) return;
  send({a:'key',k:b.dataset.k});
});
function sendText(enter){
  const v=$('txt').value;
  if(!v && !enter) return;
  send({a:'text',t:v,enter:enter});
  $('txt').value='';
}
$('txt').addEventListener('keydown',e=>{ if(e.key==='Enter'){e.preventDefault(); sendText(true);} });

/* ---- touch on the screen ---- */
function fr(t){
  const r=wrap.getBoundingClientRect();
  return {fx:Math.min(1,Math.max(0,(t.clientX-r.left)/r.width)),
          fy:Math.min(1,Math.max(0,(t.clientY-r.top)/r.height))};
}
function P(t){ return REL ? {} : fr(t); }   // absolute position, or none in touchpad mode
let st=null, lp=null, lastMv=0, sc=null;

wrap.addEventListener('touchstart',e=>{
  e.preventDefault();
  if(!CAPS.mouse) return;
  clearTimeout(lp);
  if(e.touches.length>=2){
    if(st && st.drag) send({a:'up'});
    st={multi:true};
    sc={y:(e.touches[0].clientY+e.touches[1].clientY)/2, acc:0};
    return;
  }
  const t=e.touches[0], p=P(t);
  st={x:t.clientX,y:t.clientY,lx:t.clientX,ly:t.clientY,p:p,moved:false,fired:false,drag:false};
  if(mode==='drag'){
    st.drag=true; send(Object.assign({a:'down'},p));
  }else{
    lp=setTimeout(()=>{
      if(st && !st.moved && !st.multi){
        st.fired=true; send(Object.assign({a:'click',b:'right'},st.p));
        if(navigator.vibrate) navigator.vibrate(30);
      }
    },600);
  }
},{passive:false});

wrap.addEventListener('touchmove',e=>{
  e.preventDefault();
  if(!st) return;
  if(st.multi){
    if(e.touches.length<2) return;
    const y=(e.touches[0].clientY+e.touches[1].clientY)/2;
    sc.acc+=(y-sc.y)/6; sc.y=y;
    const n=Math.trunc(sc.acc);
    if(n){ sc.acc-=n; send(Object.assign({a:'scroll',n:n},P(e.touches[0]))); }
    return;
  }
  const t=e.touches[0];
  if(!st.moved && Math.hypot(t.clientX-st.x,t.clientY-st.y)>8){ st.moved=true; clearTimeout(lp); }
  if(st.moved){
    if(REL){
      send({a:'rel',dx:(t.clientX-st.lx)*2.5,dy:(t.clientY-st.ly)*2.5});
      st.lx=t.clientX; st.ly=t.clientY;
    }else{
      const now=Date.now();
      if(now-lastMv>25){ lastMv=now; send(Object.assign({a:'moveto'},fr(t))); }
    }
  }
},{passive:false});

wrap.addEventListener('touchend',e=>{
  e.preventDefault(); clearTimeout(lp);
  if(!st) return;
  if(st.multi){ if(e.touches.length===0) st=null; return; }
  const p=P(e.changedTouches[0]);
  if(st.drag){ if(!REL) send(Object.assign({a:'moveto'},p)); send({a:'up'}); }
  else if(st.moved){ if(!REL) send(Object.assign({a:'moveto'},p)); }
  else if(!st.fired){ send(Object.assign({a:'click',b:'left'},p)); }
  st=null;
},{passive:false});

wrap.addEventListener('touchcancel',e=>{
  clearTimeout(lp);
  if(st && st.drag) send({a:'up'});
  st=null;
},{passive:false});
</script>
"""


SETTINGS = r"""<!doctype html>
<meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1,user-scalable=no">
<title>__SCHOOL__ - Permissions</title>
<link rel=icon href=/logo.png>
<style>
__THEME__
body{padding:16px;display:flex;flex-direction:column;gap:12px;min-height:100vh}
h2{margin:6px 0 0;color:var(--ink);font-size:1.3rem}
.opt{display:flex;align-items:center;gap:14px;background:#fff;border-radius:16px;padding:14px 16px;
 font-size:1.1rem;font-weight:700;border:2px solid var(--line);border-left:8px solid var(--c);
 box-shadow:0 3px 10px rgba(47,109,246,.12)}
.opt input{width:28px;height:28px;flex:0 0 auto;accent-color:var(--c)}
.opt .ic{font-size:1.6rem}
.opt small{display:block;color:var(--muted);font-size:.8rem;font-weight:500;margin-top:2px}
button.go{font-size:1.2rem;padding:1rem;box-shadow:0 4px 0 #139a50}
.warn{color:#d98200;font-size:.85rem;font-weight:600}
</style>
__BRAND__
<h2>What can this phone do?</h2>
<label class=opt style="--c:var(--blue)"><input type=checkbox id=ppt><span class=ic>📊</span><span>PowerPoint remote<small>Slides, notes, laser / pen</small></span></label>
<label class=opt style="--c:var(--green)"><input type=checkbox id=view><span class=ic>🖥</span><span>View the desktop<small>Live screen, watch only</small></span></label>
<label class=opt style="--c:var(--orange)"><input type=checkbox id=mouse><span class=ic>🖱</span><span>Control the mouse<small>Tap, drag, scroll, right-click</small></span></label>
<label class=opt style="--c:var(--purple)"><input type=checkbox id=keys><span class=ic>⌨</span><span>Use the keyboard<small>Type text and press keys</small></span></label>
<div class=warn>Only turn on what you need. Anyone with the PIN can change these.</div>
<button class=go onclick="go()">Continue</button>
<script>
const CAPS=__CAPS__;
for(const k in CAPS) document.getElementById(k).checked=CAPS[k];
async function go(){
  const caps=Object.keys(CAPS).filter(k=>document.getElementById(k).checked);
  try{
    const r=await fetch('/caps',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({caps:caps})});
    if(r.status===403){location='/login';return;}
  }catch(e){}
  location='/';
}
</script>
"""


def authed():
    return session.get("ok") is True


@app.route("/logo.png")
@app.route("/favicon.ico")
def logo_route():
    if not LOGO_PNG:
        abort(404)
    return send_file(io.BytesIO(LOGO_PNG), mimetype="image/png", max_age=86400)


@app.route("/settings")
def settings_route():
    if not authed():
        return redirect("/login")
    return page(SETTINGS)


@app.route("/caps", methods=["POST"])
def caps_route():
    if not authed():
        abort(403)
    d = request.get_json(force=True)
    names = {c for c, _ in CAP_INFO}
    CAPS.clear()
    CAPS.update(set(d.get("caps", [])) & names)
    print("  Phone permissions now: " + (", ".join(sorted(CAPS)) or "none"))
    return "ok"


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if secrets.compare_digest(request.form.get("pin", ""), PIN):
            session["ok"] = True
            return redirect("/settings")
        time.sleep(1)
        return LOGIN.replace("{err}", "Wrong PIN"), 403
    return LOGIN.replace("{err}", "")


@app.route("/")
def home():
    if not authed():
        return redirect("/login")
    if "ppt" not in CAPS:
        return redirect("/desktop" if CAPS else "/settings")
    return page(PAGE)


@app.route("/desktop")
def desktop():
    if not authed():
        return redirect("/login")
    if not CAPS & {"view", "mouse", "keys"}:
        return redirect("/")
    return page(DESKTOP)


@app.route("/api", methods=["POST"])
def api():
    if not authed():
        abort(403)
    need("ppt")
    with LOCK:
        act(request.get_json(force=True))
    return "ok"


@app.route("/dapi", methods=["POST"])
def dapi():
    if not authed():
        abort(403)
    d = request.get_json(force=True)
    a = d.get("a")
    if a in MOUSE_ACTIONS:
        need("mouse")
    elif a in KEY_ACTIONS:
        need("keys")
    else:
        abort(400)
    with DLOCK:
        try:
            dact(d)
        except Exception as e:
            print("desktop input error:", e)
    return "ok"


@app.route("/screen")
def screen_route():
    if not authed():
        abort(403)
    need("view")
    try:
        w = max(300, min(int(request.args.get("w", 1000)), 2400))
        q = max(20, min(int(request.args.get("q", 55)), 90))
        with SLOCK:
            data = grab(w, q)
        x, y = pyautogui.position()
        W, H = pyautogui.size()
    except Exception as e:
        print("screen error:", e)
        abort(500)
    resp = Response(data, mimetype="image/jpeg")
    resp.headers["X-Mouse"] = f"{x / W:.4f},{y / H:.4f}"
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/status")
def status_route():
    if not authed():
        abort(403)
    need("ppt")
    with LOCK:
        return jsonify(status())


@app.route("/thumb/<int:n>")
def thumb_route(n):
    if not authed():
        abort(403)
    need("ppt")
    with LOCK:
        try:
            data = thumb(n)
        except Exception as e:
            print("thumb error:", e)
            data = None
    if data is None:
        abort(404)
    return send_file(io.BytesIO(data), mimetype="image/jpeg", max_age=3600)


def lan_ip():
    for target in (("8.8.8.8", 80), ("10.255.255.255", 1), ("192.168.255.255", 1)):
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(target)       # UDP: nothing is actually sent
            ip = s.getsockname()[0]
            if ip and not ip.startswith("127."):
                return ip
        except Exception:
            pass
        finally:
            s.close()
    try:
        return socket.gethostbyname(socket.gethostname())
    except Exception:
        return "127.0.0.1"


def run_server():
    import logging
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    app.run(host="0.0.0.0", port=PORT, threaded=True, use_reloader=False)


def make_qr(text):
    import qrcode
    qr = qrcode.QRCode(border=2, box_size=8, error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#0a2a66", back_color="white")
    return getattr(img, "get_image", lambda: img)().convert("RGB")


def show_window(url):
    import tkinter as tk
    from PIL import ImageTk
    BG, BLUE, INK, MUTED, RED = "#e4f1ff", "#2f6df6", "#12306b", "#5f78a8", "#ff5a5f"
    root = tk.Tk()
    root.title(f"{SCHOOL} - PPT Remote")
    root.configure(bg=BG)
    root.resizable(False, False)

    logo = None
    if LOGO_PNG:
        try:
            pil = Image.open(io.BytesIO(LOGO_PNG)).convert("RGBA")
            root._icon = ImageTk.PhotoImage(pil)
            root.iconphoto(True, root._icon)               # window icon
            logo = ImageTk.PhotoImage(pil.resize((84, 84), Image.LANCZOS))
        except Exception:
            logo = None

    head = tk.Frame(root, bg=BG)
    head.pack(padx=28, pady=(20, 2))
    if logo:
        lg = tk.Label(head, image=logo, bg=BG)
        lg.image = logo
        lg.pack(side="left", padx=(0, 14))
    box = tk.Frame(head, bg=BG)
    box.pack(side="left")
    tk.Label(box, text=SCHOOL, font=("Helvetica", 24, "bold"), fg=BLUE, bg=BG).pack(anchor="w")
    tk.Label(box, text="PPT Remote", font=("Helvetica", 13), fg=MUTED, bg=BG).pack(anchor="w")

    tk.Label(root, text="Scan with your phone camera", font=("Helvetica", 16, "bold"),
             fg=INK, bg=BG).pack(pady=(12, 0))
    tk.Label(root, text="Phone and computer must be on the same Wi-Fi",
             font=("Helvetica", 11), fg=MUTED, bg=BG).pack()

    photo = ImageTk.PhotoImage(make_qr(url).resize((280, 280), Image.NEAREST))
    card = tk.Frame(root, bg="white", highlightbackground=BLUE, highlightcolor=BLUE, highlightthickness=4)
    card.pack(pady=12)
    ql = tk.Label(card, image=photo, bg="white")
    ql.image = photo
    ql.pack(padx=8, pady=8)

    tk.Label(root, text="Or type this address in the phone's browser:",
             font=("Helvetica", 11), fg=MUTED, bg=BG).pack()
    tk.Label(root, text=url, font=("Courier", 17, "bold"), fg=INK, bg=BG).pack()
    tk.Label(root, text="Then enter this PIN:", font=("Helvetica", 11), fg=MUTED, bg=BG).pack(pady=(10, 0))
    tk.Label(root, text=PIN, font=("Courier", 32, "bold"), fg=BLUE, bg=BG).pack()

    stop = tk.Label(root, text="Stop", font=("Helvetica", 14, "bold"), fg="white", bg=RED,
                    padx=44, pady=8, cursor="hand2")
    stop.bind("<Button-1>", lambda e: root.destroy())
    stop.pack(pady=(16, 22))
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()


def port_free():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("0.0.0.0", PORT))
        return True
    except OSError:
        return False
    finally:
        s.close()


if __name__ == "__main__":
    # Auto-update check (only runs in the built Windows/Mac app, see desktop_updater.py)
    try:
        import desktop_updater
        desktop_updater.check()
    except Exception:
        pass

    CAPS.update(choose_caps())
    if not port_free():
        msg = f"Port {PORT} is already in use. Is PPT Remote already running?"
        print(msg)
        try:
            import tkinter as tk
            from tkinter import messagebox
            tk.Tk().withdraw()
            messagebox.showerror("PPT Remote", msg)
        except Exception:
            pass
        sys.exit(1)

    threading.Thread(target=run_server, daemon=True).start()
    url = f"http://{lan_ip()}:{PORT}"
    mode = "PowerPoint link (COM)" if w32 else "keyboard fallback (PowerPoint link needs pywin32 on Windows)"
    print(f"\n  Mode:           {mode}")
    print(f"  Screen capture: {'mss (fast)' if mss else 'pyautogui (slower; pip3 install mss for speed)'}")
    print("  Starts with:    " + ", ".join(lbl.split(" (")[0] for c, lbl in CAP_INFO if c in CAPS))
    print(f"  Open on phone:  {url}")
    print(f"  PIN:            {PIN}\n")
    try:
        show_window(url)
    except Exception as e:   # no tkinter / qrcode: fall back to the terminal
        print(f"  (QR window unavailable: {e})")
        try:
            import qrcode
            q = qrcode.QRCode(border=1)
            q.add_data(url)
            q.print_ascii(invert=True)
        except Exception:
            pass
        print("  Press Ctrl+C to stop.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
