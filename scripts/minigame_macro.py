"""
AutoFish by Herlove v2.0
Cyberpunk Neon Fishing Macro
"""

import time, sys, os, json, threading, ctypes, random
import tkinter as tk
from tkinter import messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ═══ Auto Admin ═══
def is_admin():
    if sys.platform != 'win32': return True
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

if not is_admin() and sys.platform == 'win32':
    try:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable,
            f'"{os.path.abspath(__file__)}"', None, 1)
    except: pass
    sys.exit()

try:
    import keyboard as kb
    import mss
    import cv2
    import numpy as np
    from PIL import Image, ImageTk, ImageGrab
    import pytesseract
except ImportError as e:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("AutoFish", f"Missing: {e}\npip install keyboard mss opencv-python numpy Pillow pytesseract")
    sys.exit(1)

for p in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
          r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]:
    if os.path.exists(p): pytesseract.pytesseract.tesseract_cmd = p; break

# ═══ SendInput Scancode ═══
SCAN = {'q':0x10,'w':0x11,'e':0x12,'r':0x13,'a':0x1E,'s':0x1F,'d':0x20,'f':0x21,'space':0x39}
if sys.platform == 'win32':
    PUL = ctypes.POINTER(ctypes.c_ulong)
    class KEYBDINPUT(ctypes.Structure):
        _fields_=[("wVk",ctypes.c_ushort),("wScan",ctypes.c_ushort),("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
    class HARDWAREINPUT(ctypes.Structure):
        _fields_=[("uMsg",ctypes.c_ulong),("wParamL",ctypes.c_short),("wParamH",ctypes.c_ushort)]
    class MOUSEINPUT(ctypes.Structure):
        _fields_=[("dx",ctypes.c_long),("dy",ctypes.c_long),("mouseData",ctypes.c_ulong),("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
    class INPUT_UNION(ctypes.Union):
        _fields_=[("ki",KEYBDINPUT),("mi",MOUSEINPUT),("hi",HARDWAREINPUT)]
    class INPUT(ctypes.Structure):
        _fields_=[("type",ctypes.c_ulong),("ii",INPUT_UNION)]

def press_key(key):
    sc = SCAN.get(key.lower())
    if not sc or sys.platform != 'win32': return False
    try:
        extra = ctypes.c_ulong(0)
        ii = INPUT_UNION(); ii.ki = KEYBDINPUT(0, sc, 0x0008, 0, ctypes.pointer(extra))
        x = INPUT(1, ii)
        ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))
        time.sleep(0.05)
        ii2 = INPUT_UNION(); ii2.ki = KEYBDINPUT(0, sc, 0x0008|0x0002, 0, ctypes.pointer(extra))
        x2 = INPUT(1, ii2)
        ctypes.windll.user32.SendInput(1, ctypes.pointer(x2), ctypes.sizeof(x2))
        return True
    except: return False

# ═══ Fast Capture ═══
_sct = None
def get_sct():
    global _sct
    if _sct is None:
        try: _sct = mss.mss()
        except: pass
    return _sct

def grab(r):
    left,top,w,h = int(r["left"]),int(r["top"]),int(r["width"]),int(r["height"])
    s = get_sct()
    if s:
        try:
            arr = np.array(s.grab({"left":left,"top":top,"width":w,"height":h}))
            if arr.size > 0: return cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)
        except: pass
    try:
        img = ImageGrab.grab(bbox=(left,top,left+w,top+h))
        if img: return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    except: pass
    return None

# ═══ Fast OCR ═══
VALID = set("qweasd")
def read_all(gray, num_keys):
    if gray is None or gray.size < 100: return None
    kernel = np.array([[0,-1,0],[-1,5,-1],[0,-1,0]])
    gray = cv2.filter2D(gray, -1, kernel)
    gray = cv2.GaussianBlur(gray, (3,3), 0)
    big = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
    for method in range(3):
        try:
            if method == 0:
                _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
            elif method == 1:
                _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
                t = cv2.bitwise_not(t)
            else:
                _, t = cv2.threshold(big, 180, 255, cv2.THRESH_BINARY)
            if np.count_nonzero(t) > t.size // 2: t = cv2.bitwise_not(t)
            text = pytesseract.image_to_string(t, config='--psm 7 -c tessedit_char_whitelist=QWEASDqweasd').strip().lower()
            chars = [c for c in text if c in VALID]
            if len(chars) == num_keys and all(c in VALID for c in chars): return chars
        except: continue
    return None

def make_debug(frame, keys, num):
    d = frame.copy()
    h,w = d.shape[:2]; sw = w // num
    # HUD border
    cv2.rectangle(d, (0,0), (w-1,h-1), (56,189,248), 2)
    for i in range(num):
        x1, x2 = i*sw, min((i+1)*sw, w)
        if keys and i < len(keys) and keys[i]:
            cv2.rectangle(d, (x1+2,2), (x2-2,h-2), (0,255,0), 2)
            cv2.putText(d, keys[i].upper(), (x1+8,24), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,255,0), 2)
        else:
            cv2.rectangle(d, (x1+2,2), (x2-2,h-2), (0,0,255), 1)
    cv2.putText(d, "AUTOFISH ACTIVE", (8, h-8), cv2.FONT_HERSHEY_SIMPLEX, 0.35, (0,255,255), 1)
    return d

# ═══ Config ═══
CFG = os.path.join(SCRIPT_DIR, "config.json")
region = None
def load_cfg():
    global region
    try:
        with open(CFG,"r") as f: c=json.load(f); region=c.get("region"); return c
    except: return {}
def save_cfg(**kw):
    c=load_cfg(); c.update(kw); c["region"]=region
    try:
        with open(CFG,"w") as f: json.dump(c,f,indent=2)
    except: pass
load_cfg()

# ═══ Theme ═══
BG = "#020617"
CARD = "#0f172a"
ACCENT = "#22d3ee"
GREEN = "#4ade80"
RED = "#f87171"
DIM = "#334155"
WHITE = "#e2e8f0"
GLOW = "#0ea5e9"

# ═══ App ═══
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AutoFish by Herlove")
        self.root.geometry("480x700")
        self.root.configure(bg=BG)
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        self.running = False
        self.session_keys = 0
        self.session_start = 0
        self.debug_img = None
        self.fish_count = 0
        self.last_time = time.time()
        self.num_keys = tk.IntVar(value=load_cfg().get("num_keys", 5))
        self.key_delay = tk.IntVar(value=60)

        # Icon
        try:
            ico = os.path.join(SCRIPT_DIR, "logo.png")
            if os.path.exists(ico):
                img = ImageTk.PhotoImage(Image.open(ico).resize((32,32), Image.LANCZOS))
                self.root.iconphoto(True, img); self._ico = img
        except: pass

        self._build()
        self._preview_loop()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        # ═══ Header ═══
        hdr = tk.Frame(self.root, bg=BG)
        hdr.pack(fill=tk.X, padx=16, pady=(12,0))

        # Logo
        try:
            lp = os.path.join(SCRIPT_DIR, "mascot.png")
            if not os.path.exists(lp): lp = os.path.join(SCRIPT_DIR, "logo.png")
            if os.path.exists(lp):
                li = Image.open(lp).resize((56, 56), Image.LANCZOS)
                self._logo = ImageTk.PhotoImage(li)
                tk.Label(hdr, image=self._logo, bg=BG).pack(side=tk.LEFT, padx=(0,12))
        except: pass

        tf = tk.Frame(hdr, bg=BG); tf.pack(side=tk.LEFT)
        tk.Label(tf, text="AutoFish", bg=BG, fg=ACCENT, font=("Segoe UI",22,"bold")).pack(anchor="w")
        tk.Label(tf, text="AUTO FISHING SYSTEM v2.0", bg=BG, fg=GLOW,
                 font=("Consolas",9,"bold")).pack(anchor="w")
        tk.Label(tf, text="by Herlove", bg=BG, fg=DIM, font=("Segoe UI",8)).pack(anchor="w")

        self.lbl_st = tk.Label(hdr, text="○ OFF", bg=BG, fg=DIM, font=("Consolas",12,"bold"))
        self.lbl_st.pack(side=tk.RIGHT)

        # Gradient line
        line = tk.Canvas(self.root, bg=BG, height=3, highlightthickness=0)
        line.pack(fill=tk.X, padx=16, pady=(8,0))
        for i in range(448):
            t = i/448
            r = int(14+(34-14)*t); g = int(165+(211-165)*t); b = int(233+(238-233)*t)
            line.create_line(i+16, 0, i+16, 3, fill=f"#{r:02x}{g:02x}{b:02x}")

        # ═══ Preview ═══
        pf = tk.Frame(self.root, bg=CARD, highlightbackground=GLOW, highlightthickness=1)
        pf.pack(fill=tk.X, padx=16, pady=(10,6))
        self.preview = tk.Label(pf, bg="#030a1a", text="F6 Select Region",
                                 fg=DIM, font=("Consolas",9))
        self.preview.pack(fill=tk.X, ipady=32)

        # ═══ START ═══
        self.btn = tk.Button(self.root, text="START  (F5)", bg=GREEN, fg="#020617",
            font=("Segoe UI",15,"bold"), relief="flat", pady=8, cursor="hand2",
            activebackground="#22c55e", highlightbackground=ACCENT, highlightthickness=1,
            command=self.toggle)
        self.btn.pack(fill=tk.X, padx=16, pady=(2,6))
        self.btn.bind("<Enter>", lambda e: self.btn.config(bg="#22c55e" if not self.running else "#ef4444"))
        self.btn.bind("<Leave>", lambda e: self.btn.config(bg=GREEN if not self.running else RED))

        # ═══ Stats Dashboard ═══
        sf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        sf.pack(fill=tk.X, padx=16, pady=(0,6))
        si = tk.Frame(sf, bg=CARD); si.pack(fill=tk.X, padx=14, pady=10)

        # Keys count
        kf = tk.Frame(si, bg=CARD); kf.pack(side=tk.LEFT)
        self.lbl_cnt = tk.Label(kf, text="0", bg=CARD, fg=ACCENT, font=("Consolas",30,"bold"))
        self.lbl_cnt.pack()
        tk.Label(kf, text="KEYS", bg=CARD, fg=DIM, font=("Consolas",7,"bold")).pack()

        sep = tk.Frame(si, bg=DIM, width=1); sep.pack(side=tk.LEFT, fill=tk.Y, padx=12)

        # Fish count
        ff = tk.Frame(si, bg=CARD); ff.pack(side=tk.LEFT)
        self.lbl_fish = tk.Label(ff, text="0", bg=CARD, fg=GLOW, font=("Consolas",30,"bold"))
        self.lbl_fish.pack()
        tk.Label(ff, text="FISH", bg=CARD, fg=DIM, font=("Consolas",7,"bold")).pack()

        sep2 = tk.Frame(si, bg=DIM, width=1); sep2.pack(side=tk.LEFT, fill=tk.Y, padx=12)

        # Speed
        spf = tk.Frame(si, bg=CARD); spf.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.lbl_spd = tk.Label(spf, text="0.0", bg=CARD, fg="#94a3b8", font=("Consolas",18,"bold"))
        self.lbl_spd.pack()
        tk.Label(spf, text="FPS", bg=CARD, fg=DIM, font=("Consolas",7,"bold")).pack()

        # ═══ Controls ═══
        cf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        cf.pack(fill=tk.X, padx=16, pady=(0,6))
        ci = tk.Frame(cf, bg=CARD); ci.pack(fill=tk.X, padx=8, pady=6)

        for txt, clr, cmd in [
            ("Select (F6)", ACCENT, self.pick_region),
            ("Test Read", GREEN, self.test_read),
            ("Test Key", "#a78bfa", self.test_press)
        ]:
            b = tk.Button(ci, text=txt, bg="#0c1629", fg=clr, font=("Segoe UI",9,"bold"),
                relief="flat", padx=8, pady=3, cursor="hand2", activebackground="#1e293b",
                command=cmd)
            b.pack(side=tk.LEFT, padx=2)

        self.lbl_reg = tk.Label(ci, text="not set", bg=CARD, fg="#f97316", font=("Consolas",9))
        self.lbl_reg.pack(side=tk.RIGHT, padx=6)

        # ═══ Settings ═══
        stf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        stf.pack(fill=tk.X, padx=16, pady=(0,6))
        self._sl(stf, "Keys / Lane", self.num_keys, 2, 10, ACCENT)
        self._sl(stf, "Key Delay (ms)", self.key_delay, 30, 200, GREEN)

        # ═══ Log ═══
        lhdr = tk.Frame(self.root, bg=BG)
        lhdr.pack(fill=tk.X, padx=18, pady=(4,1))
        tk.Label(lhdr, text="> SYSTEM LOG", bg=BG, fg=DIM, font=("Consolas",8,"bold")).pack(side=tk.LEFT)

        self.log_box = tk.Text(self.root, bg="#030a1a", fg=GREEN, font=("Consolas",9),
            height=6, relief="flat", wrap=tk.WORD, state=tk.DISABLED, padx=10, pady=6,
            insertbackground=ACCENT, selectbackground="#1e3a5f")
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0,8))

        # Footer
        tk.Label(self.root, text="F5 Start/Stop  |  F6 Region  |  Admin Mode",
                 bg=BG, fg="#0f172a", font=("Segoe UI",7)).pack(pady=(0,6))

        self.root.bind("<F5>", lambda e: self.toggle())
        self.root.bind("<F6>", lambda e: self.pick_region())

        # Init log
        self.log("> AutoFish by Herlove v2.0")
        self.log(f"> Admin: {'YES' if is_admin() else 'NO!'}")
        try:
            v = pytesseract.get_tesseract_version()
            self.log(f"> Tesseract {v}")
        except:
            self.log("> ERROR: Tesseract not found!")
        self.log("> Ready. Select region (F6) then Start (F5)")

    def _sl(self, p, label, var, lo, hi, color):
        f = tk.Frame(p, bg=CARD); f.pack(fill=tk.X, padx=10, pady=2)
        tk.Label(f, text=label, bg=CARD, fg=WHITE, font=("Segoe UI",9)).pack(side=tk.LEFT)
        vl = tk.Label(f, text=str(var.get()), bg=CARD, fg=color, font=("Consolas",10,"bold"), width=4)
        vl.pack(side=tk.RIGHT)
        tk.Scale(f, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var, bg=CARD, fg=CARD,
            troughcolor=BG, highlightthickness=0, showvalue=False, length=130, sliderlength=14,
            activebackground=color,
            command=lambda v,l=vl:l.config(text=str(int(float(v))))).pack(side=tk.RIGHT, padx=4)

    def pick_region(self):
        global region
        if self.running: self.log("> Stop first!"); return
        self.root.iconify(); time.sleep(0.3)
        sel = tk.Toplevel(self.root)
        sel.overrideredirect(True); sel.attributes("-topmost",True); sel.attributes("-alpha",0.25)
        sw,sh = sel.winfo_screenwidth(),sel.winfo_screenheight()
        sel.geometry(f"{sw}x{sh}+0+0"); sel.configure(bg="black")
        c = tk.Canvas(sel, bg="black", highlightthickness=0, cursor="cross")
        c.pack(fill=tk.BOTH, expand=True)
        c.create_rectangle(sw//2-220,15,sw//2+220,80, fill="black", outline=ACCENT, width=2)
        c.create_text(sw//2,33, text="Drag over letter boxes (exclude counter)",
            fill=ACCENT, font=("Segoe UI",13,"bold"))
        c.create_text(sw//2,58, text="ESC = Cancel", fill="#64748b", font=("Segoe UI",10))
        pos = c.create_text(sw//2,100, text="", fill=ACCENT, font=("Consolas",12,"bold"))
        st = {"sx":0,"sy":0,"r":None}
        c.bind("<Motion>", lambda e: c.itemconfig(pos, text=f"X:{e.x}  Y:{e.y}"))
        def _p(e):
            st["sx"],st["sy"]=e.x,e.y
            if st["r"]: c.delete(st["r"])
            st["r"]=c.create_rectangle(e.x,e.y,e.x,e.y, outline=ACCENT, width=3)
        def _d(e):
            c.coords(st["r"],st["sx"],st["sy"],e.x,e.y); c.delete("sz")
            c.create_text((st["sx"]+e.x)//2,min(st["sy"],e.y)-14,
                text=f"{abs(e.x-st['sx'])}x{abs(e.y-st['sy'])}", fill=ACCENT,
                font=("Consolas",13,"bold"), tags="sz")
        def _r(e):
            global region
            x1,y1=min(st["sx"],e.x),min(st["sy"],e.y); x2,y2=max(st["sx"],e.x),max(st["sy"],e.y)
            if (x2-x1)>10 and (y2-y1)>10:
                region={"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}; save_cfg()
                self.lbl_reg.config(text=f"{x2-x1}x{y2-y1}", fg=GREEN)
                self.log(f"> Region: {x2-x1}x{y2-y1} @ ({x1},{y1})")
            sel.destroy(); self.root.deiconify()
        c.bind("<ButtonPress-1>",_p); c.bind("<B1-Motion>",_d); c.bind("<ButtonRelease-1>",_r)
        sel.bind("<Escape>", lambda e:(sel.destroy(),self.root.deiconify()))
        sel.after(50, sel.focus_force)

    def test_read(self):
        if not region: self.log("> Select region first!"); return
        t0 = time.time()
        frame = grab(region)
        if frame is None: self.log("> Capture failed!"); return
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        keys = read_all(gray, self.num_keys.get())
        ms = int((time.time()-t0)*1000)
        self.debug_img = make_debug(frame, keys, self.num_keys.get())
        self._show_debug()
        if keys:
            self.log(f"> Read ({ms}ms): {' '.join(k.upper() for k in keys)}")
        else:
            self.log(f"> Failed ({ms}ms) - adjust region")

    def test_press(self):
        self.log("> Pressing D+S in 3s... (click game!)")
        def _do():
            time.sleep(3)
            press_key('d'); time.sleep(0.2); press_key('s')
            self.root.after(0, self.log, "> Pressed D S!")
        threading.Thread(target=_do, daemon=True).start()

    def toggle(self):
        if self.running:
            self.running = False
        else:
            if not region: messagebox.showwarning("AutoFish","Select Region (F6)!"); return
            self.running = True; save_cfg(num_keys=self.num_keys.get())
            self.btn.config(text="STOP  (F5)", bg=RED, activebackground="#ef4444")
            self.lbl_st.config(text="● ON", fg=GREEN)
            self.log("> Switch to game in 3s!")
            threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        num = self.num_keys.get()
        kd = self.key_delay.get() / 1000
        self.session_keys = 0; self.session_start = time.time()
        self.fish_count = 0; self.last_time = time.time()

        test = grab(region)
        if test is None:
            self.root.after(0, self.log, "> Capture failed!")
            self.running = False; self.root.after(0, self._reset); return

        time.sleep(3)
        self.root.after(0, self.log, "> Scanning...")
        last_seq = ""

        while self.running:
            try:
                frame = grab(region)
                if frame is None: time.sleep(0.06); continue
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                # FPS
                now = time.time()
                fps = 1.0 / max(now - self.last_time, 0.001)
                self.last_time = now

                keys = read_all(gray, num)
                self.debug_img = make_debug(frame, keys, num)

                if keys and len(keys) == num:
                    seq = "".join(keys)
                    if seq != last_seq:
                        self.fish_count += 1
                        display = " ".join(k.upper() for k in keys)
                        self.root.after(0, self.log, f"> #{self.fish_count} {display}")

                        # Anti-ban: random pause 5% chance
                        if random.random() < 0.05:
                            time.sleep(random.uniform(0.5, 1.5))

                        for key in keys:
                            if not self.running: break
                            press_key(key)
                            self.session_keys += 1
                            time.sleep(kd + random.uniform(0.02, 0.08))

                        # Sound
                        if sys.platform == 'win32':
                            try:
                                import winsound
                                winsound.Beep(800, 80)
                            except: pass

                        self.root.after(0, self._update_stats, fps)
                        last_seq = seq
                        time.sleep(0.8 + random.uniform(0.1, 0.3))
                    else:
                        time.sleep(0.04)
                else:
                    time.sleep(0.06)

            except Exception as e:
                self.root.after(0, self.log, f"> Error: {e}")
                time.sleep(0.5)

        self.root.after(0, self._reset)

    def _update_stats(self, fps=0):
        self.lbl_cnt.config(text=str(self.session_keys))
        self.lbl_fish.config(text=str(self.fish_count))
        self.lbl_spd.config(text=f"{fps:.1f}")

    def _reset(self):
        self.btn.config(text="START  (F5)", bg=GREEN, activebackground="#22c55e")
        self.lbl_st.config(text="○ OFF", fg=DIM)
        self.log("> Stopped")

    def _show_debug(self):
        if self.debug_img is None: return
        try:
            rgb = cv2.cvtColor(self.debug_img, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(rgb)
            w,h = img.size; nw=448; nh=max(int(h*nw/w),20)
            img = img.resize((nw,nh), Image.NEAREST)
            photo = ImageTk.PhotoImage(img)
            self.preview.config(image=photo, text="")
            self.preview.image = photo
        except: pass

    def _preview_loop(self):
        if self.running and self.debug_img is not None:
            self._show_debug()
        elif region:
            try:
                frame = grab(region)
                if frame is not None:
                    d = make_debug(frame, None, self.num_keys.get())
                    rgb = cv2.cvtColor(d, cv2.COLOR_BGR2RGB)
                    img = Image.fromarray(rgb)
                    w,h = img.size; nw=448; nh=max(int(h*nw/w),20)
                    img = img.resize((nw,nh), Image.NEAREST)
                    photo = ImageTk.PhotoImage(img)
                    self.preview.config(image=photo, text="")
                    self.preview.image = photo
            except: pass
        self.root.after(150, self._preview_loop)

    def log(self, msg):
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_box.see(tk.END)
        lines = int(self.log_box.index("end-1c").split(".")[0])
        if lines > 80: self.log_box.delete("1.0", f"{lines-80}.0")
        self.log_box.config(state=tk.DISABLED)

    def _quit(self): self.running=False; time.sleep(0.1); self.root.destroy()
    def run(self): self.root.mainloop()

if __name__ == "__main__":
    App().run()
