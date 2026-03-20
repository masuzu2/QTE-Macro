"""
QTE Minigame Macro v3 - Simple & Accurate
Template Matching Only - No Tesseract needed
"""

import time, sys, os, json, threading, math
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# ============================================================
# Lib Check (ไม่ต้อง Tesseract แล้ว)
# ============================================================
MISSING = []
for lib, pkg in [("keyboard","keyboard"),("mss","mss"),
                  ("cv2","opencv-python"),("numpy","numpy"),("PIL","Pillow")]:
    try: __import__(lib)
    except ImportError: MISSING.append(pkg)

if MISSING:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing",
        f"Install:\npip install {' '.join(MISSING)}")
    sys.exit(1)

import keyboard as kb
import mss, mss.tools
import cv2, numpy as np
from PIL import Image, ImageTk, ImageDraw, ImageGrab

# ============================================================
# Screen Capture - ลองหลายวิธี
# ============================================================
class ScreenCapture:
    def __init__(self):
        self.method = None
        self.sct = None
        self._detect()

    def _detect(self):
        try:
            s = mss.mss()
            t = s.grab(s.monitors[0])
            if t.size[0] > 0:
                self.method = "mss"; self.sct = s; return
        except: pass
        try:
            import pyautogui
            t = pyautogui.screenshot(region=(0,0,50,50))
            if t: self.method = "pyautogui"; return
        except: pass
        try:
            t = ImageGrab.grab(bbox=(0,0,50,50))
            if t: self.method = "pil"; return
        except: pass

    def grab(self, region):
        x,y,w,h = region["left"],region["top"],region["width"],region["height"]
        try:
            if self.method == "mss":
                shot = self.sct.grab(region)
                return cv2.cvtColor(np.array(shot), cv2.COLOR_BGRA2BGR)
            elif self.method == "pyautogui":
                import pyautogui
                return cv2.cvtColor(np.array(pyautogui.screenshot(region=(x,y,w,h))), cv2.COLOR_RGB2BGR)
            elif self.method == "pil":
                return cv2.cvtColor(np.array(ImageGrab.grab(bbox=(x,y,x+w,y+h))), cv2.COLOR_RGB2BGR)
        except:
            self._detect()
        return None

    def grab_pil(self, region):
        x,y,w,h = region["left"],region["top"],region["width"],region["height"]
        try:
            if self.method == "mss":
                s = self.sct.grab(region)
                return Image.frombytes("RGB", s.size, s.bgra, "raw", "BGRX")
            elif self.method == "pyautogui":
                import pyautogui
                return pyautogui.screenshot(region=(x,y,w,h))
            elif self.method == "pil":
                return ImageGrab.grab(bbox=(x,y,x+w,y+h))
        except: pass
        return None

    def close(self):
        if self.sct:
            try: self.sct.close()
            except: pass


# ============================================================
# Template Engine - หัวใจของระบบ
# ============================================================
class TemplateEngine:
    def __init__(self, template_dir):
        self.dir = template_dir
        self.templates = {}   # key -> list of gray images
        self.scaled = {}      # key -> list of (scale, gray) pre-computed
        os.makedirs(template_dir, exist_ok=True)
        self.load()

    def load(self):
        self.templates.clear()
        self.scaled.clear()
        if not os.path.isdir(self.dir): return
        for f in os.listdir(self.dir):
            if not f.lower().endswith(".png"): continue
            name = f.rsplit(".",1)[0].lower()
            key = name.split("_")[0]
            if len(key) != 1: continue
            img = cv2.imread(os.path.join(self.dir, f), cv2.IMREAD_GRAYSCALE)
            if img is not None:
                if key not in self.templates:
                    self.templates[key] = []
                self.templates[key].append(img)
        self._precompute_scales()

    def _precompute_scales(self):
        """Pre-compute scaled versions ทุก template ทุก scale ล่วงหน้า → loop เร็วขึ้น 5x"""
        self.scaled.clear()
        scales = [1.0, 0.95, 1.05, 0.9, 1.1]
        for key, tmpls in self.templates.items():
            self.scaled[key] = []
            for tmpl in tmpls:
                for sc in scales:
                    tw, th = int(tmpl.shape[1]*sc), int(tmpl.shape[0]*sc)
                    if tw < 3 or th < 3: continue
                    self.scaled[key].append(cv2.resize(tmpl, (tw, th)))

    def detect(self, gray, threshold=0.75):
        """ตรวจจับ key จาก grayscale frame (ส่ง gray เข้ามาเลย ไม่ต้อง convert ซ้ำ)
           return (key, confidence) หรือ (None, 0)"""
        best_key, best_val = None, 0
        gh, gw = gray.shape[:2]

        for key, scaled_list in self.scaled.items():
            for tmpl in scaled_list:
                th, tw = tmpl.shape[:2]
                if th > gh or tw > gw: continue
                res = cv2.matchTemplate(gray, tmpl, cv2.TM_CCOEFF_NORMED)
                _, mx, _, _ = cv2.minMaxLoc(res)
                if mx > threshold and mx > best_val:
                    best_key, best_val = key, mx
                    # Early exit ถ้า confidence สูงมาก (เร็วขึ้น)
                    if mx > 0.95:
                        return best_key, best_val

        return best_key, best_val

    def save(self, key, frame_bgr):
        existing = [f for f in os.listdir(self.dir) if f.lower().startswith(f"{key}") and f.endswith(".png")]
        fname = f"{key}.png" if not existing else f"{key}_{len(existing)+1}.png"
        path = os.path.join(self.dir, fname)
        cv2.imwrite(path, frame_bgr)
        gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY)
        if key not in self.templates:
            self.templates[key] = []
        self.templates[key].append(gray)
        self._precompute_scales()  # rebuild scaled cache
        return fname

    @property
    def count(self):
        return sum(len(v) for v in self.templates.values())

    @property
    def keys_list(self):
        return sorted(self.templates.keys())


# ============================================================
# Frame Change Detector
# ============================================================
class FrameDetector:
    """ตรวจจับว่าหน้าจอเปลี่ยน → key ใหม่มา

    Logic:
    - หน้าจอเปลี่ยน → reset pressed flag → พร้อมกดใหม่
    - กดปุ่มแล้ว → set pressed = True → รอจนหน้าจอเปลี่ยนอีก
    - Retry: ถ้า frame เปลี่ยนแล้วแต่ match ไม่เจอ จะลองได้อีกหลาย frame
      จนกว่าจะกดสำเร็จ หรือหน้าจอเปลี่ยนอีกครั้ง
    """
    def __init__(self, threshold=12, min_cd_ms=80):
        self.threshold = threshold
        self.min_cd = min_cd_ms / 1000.0
        self.last_gray = None
        self.last_time = 0
        self.pressed = False      # กด key จาก frame ชุดนี้แล้วหรือยัง
        self.change_pct = 0       # % ที่เปลี่ยน (สำหรับ debug)

    def check_raw(self, small_gray):
        """เช็ค frame change จาก pre-resized 64x64 gray"""
        if self.last_gray is None:
            self.last_gray = small_gray.copy()
            self.pressed = False
            self.change_pct = 100
            return

        diff = cv2.absdiff(self.last_gray, small_gray)
        self.change_pct = (np.count_nonzero(diff > 25) / diff.size) * 100

        if self.change_pct >= self.threshold:
            # หน้าจอเปลี่ยนจริง → key ใหม่มา → reset
            self.last_gray = small_gray.copy()
            self.pressed = False

    def can_press(self):
        """ยังไม่ได้กด key จาก frame ชุดนี้ + ผ่าน min cooldown"""
        if self.pressed:
            return False
        if (time.time() - self.last_time) < self.min_cd:
            return False
        return True

    def record(self):
        self.last_time = time.time()
        self.pressed = True

    def reset(self):
        self.last_gray = None
        self.last_time = 0
        self.pressed = False
        self.change_pct = 0


# ============================================================
# Theme
# ============================================================
C = {
    "bg":"#0a0e17","bg2":"#0f1923","card":"#111a2b","card2":"#162032",
    "border":"#1e3050","glow":"#1a3a5c",
    "blue":"#00d4ff","purple":"#a855f7","green":"#22ff88",
    "red":"#ff4466","orange":"#ff9f43","cyan":"#00e5ff",
    "white":"#eaf0f6","dim":"#5a6a7e","dim2":"#3d4f63","input":"#0c1220",
}

# ============================================================
# Glow Button
# ============================================================
class Btn(tk.Canvas):
    def __init__(self, parent, text="", color="#00d4ff", w=200, h=44, fs=12, cmd=None, **kw):
        super().__init__(parent, width=w, height=h, bg=C["bg"], highlightthickness=0, **kw)
        self._w, self._h, self.color, self.text, self.cmd, self.fs = w, h, color, text, cmd, fs
        self._hover = False
        self.bind("<Enter>", lambda e: self._set(True))
        self.bind("<Leave>", lambda e: self._set(False))
        self.bind("<Button-1>", lambda e: self.cmd() if self.cmd else None)
        self._draw()

    def _blend(self, c1, c2, t):
        r1,g1,b1 = [int(c1[i:i+2],16) for i in (1,3,5)]
        r2,g2,b2 = [int(c2[i:i+2],16) for i in (1,3,5)]
        return f"#{int(r1+(r2-r1)*t):02x}{int(g1+(g2-g1)*t):02x}{int(b1+(b2-b1)*t):02x}"

    def _draw(self):
        self.delete("all")
        w,h,r = self._w,self._h,10
        a = 0.4 if self._hover else 0.0
        fill = self._blend(C["card"], self.color, 0.15+a*0.15)
        brd = self._blend(self.color, "#ffffff", a*0.3)
        pts = [2+r,2, w-2-r,2, w-2,2, w-2,2+r, w-2,h-2-r, w-2,h-2,
               w-2-r,h-2, 2+r,h-2, 2,h-2, 2,h-2-r, 2,2+r, 2,2]
        self.create_polygon(pts, smooth=True, fill=fill, outline=brd, width=2)
        tc = self._blend(self.color, "#ffffff", a*0.5)
        self.create_text(w//2, h//2, text=self.text, fill=tc, font=("Segoe UI", self.fs, "bold"))

    def _set(self, h): self._hover = h; self._draw()
    def set_text(self, t): self.text = t; self._draw()
    def set_color(self, c): self.color = c; self._draw()


# ============================================================
# Region Selector
# ============================================================
class RegionSelector(tk.Toplevel):
    def __init__(self, master, callback):
        super().__init__(master)
        self.callback = callback
        self.sx = self.sy = 0
        self.rect = None
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        try: self.attributes("-alpha", 0.25)
        except: pass
        sw,sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0")
        self.configure(bg="black")
        self.c = tk.Canvas(self, bg="black", highlightthickness=0, cursor="cross")
        self.c.pack(fill=tk.BOTH, expand=True)
        self.c.create_text(sw//2, 40, text="DRAG TO SELECT", fill="#00ffaa", font=("Segoe UI",22,"bold"))
        self.c.create_text(sw//2, 72, text="ลากเลือกบริเวณที่ตัวอักษรโผล่  |  ESC = ยกเลิก", fill="#ccc", font=("Segoe UI",12))
        self.c.bind("<ButtonPress-1>", self._p)
        self.c.bind("<B1-Motion>", self._d)
        self.c.bind("<ButtonRelease-1>", self._r)
        self.bind("<Escape>", lambda e: self.destroy())
        self.after(100, self.focus_force)

    def _p(self, e):
        self.sx,self.sy = e.x,e.y
        if self.rect: self.c.delete(self.rect)
        self.rect = self.c.create_rectangle(e.x,e.y,e.x,e.y, outline="#00ffaa", width=2, dash=(6,3))
    def _d(self, e):
        self.c.coords(self.rect, self.sx, self.sy, e.x, e.y)
        self.c.delete("sz")
        self.c.create_text((self.sx+e.x)//2, min(self.sy,e.y)-18,
            text=f"{abs(e.x-self.sx)} x {abs(e.y-self.sy)}", fill="#00ffaa", font=("Consolas",13,"bold"), tags="sz")
    def _r(self, e):
        x1,y1 = min(self.sx,e.x),min(self.sy,e.y)
        x2,y2 = max(self.sx,e.x),max(self.sy,e.y)
        if (x2-x1)>5 and (y2-y1)>5:
            self.callback({"left":x1,"top":y1,"width":x2-x1,"height":y2-y1})
        self.destroy()


# ============================================================
# MAIN APP
# ============================================================
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QTE Macro v3")
        self.root.geometry("540x780")
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)

        # Config
        self.cfg = self._load_cfg()
        self.running = False
        self.region = self.cfg.get("region")
        self.session_keys = 0
        self.session_start = None

        # Engine
        self.capture = ScreenCapture()
        self.engine = TemplateEngine(os.path.join(get_base_path(), "templates"))
        self.frame_det = FrameDetector(
            self.cfg.get("change_thresh", 12),
            self.cfg.get("min_cd_ms", 80)
        )

        # Vars
        self.v_scan = tk.IntVar(value=self.cfg.get("scan_ms", 25))
        self.v_thresh = tk.IntVar(value=int(self.cfg.get("match_thresh", 0.75) * 100))
        self.v_change = tk.IntVar(value=self.cfg.get("change_thresh", 12))
        self.v_mincd = tk.IntVar(value=self.cfg.get("min_cd_ms", 80))
        self.v_ontop = tk.BooleanVar(value=self.cfg.get("ontop", True))
        self.v_sound = tk.BooleanVar(value=self.cfg.get("sound", True))

        self._build()
        self.root.attributes("-topmost", self.v_ontop.get())
        self._update_region()
        self._update_templates()
        self._preview_loop()
        self._timer_loop()

        self.root.bind_all("<F5>", lambda e: self.toggle())
        self.root.bind_all("<F6>", lambda e: self.pick_region())
        self.root.bind_all("<Escape>", lambda e: self._stop())
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    # ── Config ──
    def _cfg_path(self): return os.path.join(get_base_path(), "macro_config.json")
    def _load_cfg(self):
        try:
            with open(self._cfg_path(),"r",encoding="utf-8") as f: return json.load(f)
        except: return {}
    def _save_cfg(self):
        self.cfg.update({
            "scan_ms":self.v_scan.get(), "match_thresh":self.v_thresh.get()/100,
            "change_thresh":self.v_change.get(), "min_cd_ms":self.v_mincd.get(),
            "region":self.region, "ontop":self.v_ontop.get(), "sound":self.v_sound.get(),
        })
        try:
            with open(self._cfg_path(),"w",encoding="utf-8") as f:
                json.dump(self.cfg, f, indent=2, ensure_ascii=False)
        except: pass

    # ── BUILD UI ──
    def _build(self):
        # Header gradient
        hdr = tk.Canvas(self.root, bg=C["bg"], height=70, highlightthickness=0)
        hdr.pack(fill=tk.X)
        for i in range(540):
            t=i/540; r,g,b = int(168*t), int(212+(85-212)*t), int(255+(247-255)*t)
            hdr.create_line(i,0,i,3, fill=f"#{r:02x}{g:02x}{b:02x}")
        hdr.create_text(24,28, text="QTE MACRO v3", anchor="w", fill=C["blue"], font=("Segoe UI",22,"bold"))
        cap_ok = self.capture.method is not None
        hdr.create_text(24,50, text=f"Capture: {self.capture.method or 'FAIL'} | Templates: {self.engine.count}",
                        anchor="w", fill=C["green"] if cap_ok else C["red"], font=("Consolas",9))
        self.lbl_status = hdr.create_text(500,28, text="OFF", anchor="w", fill=C["dim"], font=("Segoe UI",10,"bold"))
        self.lbl_timer = hdr.create_text(520,48, text="00:00", anchor="e", fill=C["dim2"], font=("Consolas",9))
        self.hdr = hdr

        main = tk.Frame(self.root, bg=C["bg"])
        main.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0,8))

        # ── START ──
        self.btn_start = Btn(main, text="START (F5)", color=C["green"], w=508, h=52, fs=17, cmd=self.toggle)
        self.btn_start.pack(pady=(4,8))

        # ── Stats ──
        sf = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        sf.pack(fill=tk.X, pady=(0,8))
        si = tk.Frame(sf, bg=C["card"]); si.pack(fill=tk.X, padx=14, pady=10)

        left = tk.Frame(si, bg=C["card"]); left.pack(side=tk.LEFT)
        self.lbl_count = tk.Label(left, text="0", bg=C["card"], fg=C["green"], font=("Consolas",32,"bold"))
        self.lbl_count.pack(anchor="w")
        tk.Label(left, text="KEYS PRESSED", bg=C["card"], fg=C["dim"], font=("Segoe UI",8,"bold")).pack(anchor="w")

        mid = tk.Frame(si, bg=C["card"]); mid.pack(side=tk.LEFT, padx=16)
        self.lbl_last = tk.Label(mid, text="Last: -", bg=C["card"], fg=C["purple"], font=("Segoe UI",13,"bold"))
        self.lbl_last.pack(anchor="w")
        self.lbl_speed = tk.Label(mid, text="0.0 /sec", bg=C["card"], fg=C["blue"], font=("Consolas",11))
        self.lbl_speed.pack(anchor="w")
        self.lbl_debug = tk.Label(mid, text="Match: - (0%)", bg=C["card"], fg=C["dim"], font=("Consolas",9))
        self.lbl_debug.pack(anchor="w")
        self.lbl_frame = tk.Label(mid, text="Frame: 0%", bg=C["card"], fg=C["dim"], font=("Consolas",9))
        self.lbl_frame.pack(anchor="w")

        # History (last 10 keys)
        self.lbl_history = tk.Label(si, text="", bg=C["card"], fg=C["cyan"],
                                     font=("Consolas",11,"bold"), anchor="e", width=14)
        self.lbl_history.pack(side=tk.RIGHT, padx=(0,8))
        self.key_history = []

        # Preview
        pf = tk.Frame(si, bg=C["bg2"], width=110, height=70, highlightbackground=C["border"], highlightthickness=1)
        pf.pack(side=tk.RIGHT); pf.pack_propagate(False)
        self.preview = tk.Label(pf, bg=C["bg2"], text="Preview", fg=C["dim2"], font=("Segoe UI",8))
        self.preview.pack(expand=True)

        # ── STEP 1: Region ──
        self._section(main, "STEP 1 — SELECT REGION")
        rc = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        rc.pack(fill=tk.X, pady=(0,6))
        ri = tk.Frame(rc, bg=C["card"]); ri.pack(fill=tk.X, padx=12, pady=8)
        Btn(ri, text="Select Region (F6)", color=C["blue"], w=170, h=32, fs=10, cmd=self.pick_region).pack(side=tk.LEFT)
        Btn(ri, text="Test", color=C["orange"], w=60, h=32, fs=9, cmd=self._test_capture).pack(side=tk.LEFT, padx=6)
        self.lbl_region = tk.Label(ri, text="Not set", bg=C["card"], fg=C["orange"], font=("Consolas",9))
        self.lbl_region.pack(side=tk.RIGHT)

        # ── STEP 2: Templates ──
        self._section(main, "STEP 2 — SETUP KEYS  (จับภาพตัวอักษรจากเกม)")
        tc = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        tc.pack(fill=tk.X, pady=(0,6))
        ti = tk.Frame(tc, bg=C["card"]); ti.pack(fill=tk.X, padx=12, pady=8)

        Btn(ti, text="Quick Setup", color=C["green"], w=130, h=34, fs=10, cmd=self._quick_setup).pack(side=tk.LEFT)
        Btn(ti, text="Capture 1 Key", color=C["purple"], w=120, h=34, fs=9, cmd=self._capture_one).pack(side=tk.LEFT, padx=6)
        Btn(ti, text="Clear All", color=C["red"], w=90, h=34, fs=9, cmd=self._clear_templates).pack(side=tk.LEFT, padx=6)

        self.lbl_tmpls = tk.Label(ti, text="0 templates", bg=C["card"], fg=C["dim"], font=("Segoe UI",9))
        self.lbl_tmpls.pack(side=tk.RIGHT)

        # Template keys display
        self.lbl_keys = tk.Label(tc, text="Keys: (none)", bg=C["card"], fg=C["cyan"],
                                  font=("Consolas",11,"bold"), anchor="w")
        self.lbl_keys.pack(padx=12, pady=(0,8))

        # ── STEP 3: Settings ──
        self._section(main, "STEP 3 — SETTINGS")
        stf = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        stf.pack(fill=tk.X, pady=(0,6))
        sti = tk.Frame(stf, bg=C["card"]); sti.pack(fill=tk.X, padx=12, pady=8)

        self._slider(sti, "Scan Speed", "ms", self.v_scan, 10, 100, C["cyan"])
        self._slider(sti, "Match Threshold", "%", self.v_thresh, 50, 95, C["orange"])
        self._slider(sti, "Frame Change", "%", self.v_change, 3, 40, C["purple"])
        self._slider(sti, "Min Cooldown", "ms", self.v_mincd, 30, 300, C["red"])

        chk = tk.Frame(sti, bg=C["card"]); chk.pack(fill=tk.X, pady=(4,0))
        tk.Checkbutton(chk, text="Always on Top", variable=self.v_ontop, bg=C["card"], fg=C["white"],
                       selectcolor=C["bg"], activebackground=C["card"], font=("Segoe UI",9),
                       command=lambda: self.root.attributes("-topmost", self.v_ontop.get())).pack(side=tk.LEFT, padx=(0,12))
        tk.Checkbutton(chk, text="Sound", variable=self.v_sound, bg=C["card"], fg=C["white"],
                       selectcolor=C["bg"], activebackground=C["card"], font=("Segoe UI",9)).pack(side=tk.LEFT)

        # ── LOG ──
        self._section(main, "LOG")
        lf = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        lf.pack(fill=tk.BOTH, expand=True, pady=(0,4))
        self.log_box = tk.Text(lf, bg=C["bg"], fg=C["green"], font=("Consolas",9),
                                height=5, relief="flat", wrap=tk.WORD, state=tk.DISABLED, padx=8, pady=6)
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self.log_box.tag_configure("key", foreground=C["cyan"], font=("Consolas",9,"bold"))
        self.log_box.tag_configure("err", foreground=C["red"])
        self.log_box.tag_configure("warn", foreground=C["orange"])
        self.log_box.tag_configure("ok", foreground=C["green"])
        self.log_box.tag_configure("dim", foreground=C["dim"])

        # Footer
        ft = tk.Frame(self.root, bg=C["bg2"], height=26); ft.pack(fill=tk.X, side=tk.BOTTOM)
        ft.pack_propagate(False)
        tk.Label(ft, text="F5 Start/Stop   F6 Region   ESC Stop", bg=C["bg2"], fg=C["dim2"],
                 font=("Segoe UI",8)).pack(pady=4)

        self.log(f"Capture: {self.capture.method or 'FAILED'}", "ok" if self.capture.method else "err")
        if self.engine.count > 0:
            self.log(f"Loaded {self.engine.count} templates: {' '.join(k.upper() for k in self.engine.keys_list)}", "ok")

    def _section(self, p, t):
        f = tk.Frame(p, bg=C["bg"]); f.pack(fill=tk.X, pady=(6,2))
        tk.Label(f, text=t, bg=C["bg"], fg=C["dim2"], font=("Segoe UI",8,"bold")).pack(side=tk.LEFT, padx=2)
        s = tk.Canvas(f, bg=C["bg"], height=1, highlightthickness=0)
        s.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6,0))
        s.create_line(0,0,400,0, fill=C["border"])

    def _slider(self, p, label, unit, var, lo, hi, color):
        row = tk.Frame(p, bg=C["card"]); row.pack(fill=tk.X, pady=2)
        tk.Label(row, text=f"{label}:", bg=C["card"], fg=C["white"], font=("Segoe UI",9)).pack(side=tk.LEFT)
        vl = tk.Label(row, text=f"{var.get()}{unit}", bg=C["card"], fg=color,
                       font=("Consolas",10,"bold"), width=6, anchor="e"); vl.pack(side=tk.RIGHT)
        tk.Scale(row, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var,
                 bg=C["card"], fg=C["card"], troughcolor=C["bg"], highlightthickness=0,
                 showvalue=False, length=160, sliderlength=14, activebackground=color,
                 command=lambda v, l=vl, u=unit: l.config(text=f"{int(float(v))}{u}")
                 ).pack(side=tk.RIGHT, padx=(4,2))

    # ── Region ──
    def pick_region(self):
        if self.running: self.log("Stop first!", "err"); return
        self.root.iconify()
        self.root.after(400, self._sel)
    def _sel(self):
        def done(r):
            self.region = r; self._save_cfg(); self.root.deiconify(); self._update_region()
            self.log(f"Region: {r['width']}x{r['height']} @ ({r['left']},{r['top']})", "ok")
        s = RegionSelector(self.root, done)
        s.protocol("WM_DELETE_WINDOW", lambda: (s.destroy(), self.root.deiconify()))

    def _update_region(self):
        if self.region:
            r = self.region
            self.lbl_region.config(text=f"{r['width']}x{r['height']}  ({r['left']},{r['top']})", fg=C["green"])
        else:
            self.lbl_region.config(text="Not set", fg=C["orange"])

    def _update_templates(self):
        self.lbl_tmpls.config(text=f"{self.engine.count} templates")
        keys = self.engine.keys_list
        self.lbl_keys.config(text=f"Keys: {' '.join(k.upper() for k in keys)}" if keys else "Keys: (none)")

    def _test_capture(self):
        if not self.region: messagebox.showwarning("Test","Select region first!"); return
        frame = self.capture.grab(self.region)
        if frame is None:
            self.log("Capture FAILED!", "err"); return
        self.log(f"Capture OK  {frame.shape[1]}x{frame.shape[0]}  method={self.capture.method}", "ok")
        if self.engine.count > 0:
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            key, conf = self.engine.detect(gray, self.v_thresh.get()/100)
            if key:
                self.log(f"  Detected: '{key.upper()}'  confidence: {conf:.0%}", "key")
            else:
                self.log("  No match found - ลองปรับ region หรือจับ template ใหม่", "warn")

    # ── Templates ──
    def _quick_setup(self):
        """Quick Setup - จับทุก key ทีละตัว"""
        if not self.region:
            messagebox.showwarning("Setup","Select region first!"); return

        keys_to_capture = list("qweasd")
        dlg = tk.Toplevel(self.root); dlg.title("Quick Setup"); dlg.geometry("400x320")
        dlg.configure(bg=C["bg"]); dlg.attributes("-topmost", True); dlg.resizable(False,False)

        tk.Label(dlg, text="Quick Setup", bg=C["bg"], fg=C["blue"],
                 font=("Segoe UI",16,"bold")).pack(pady=(16,4))
        tk.Label(dlg, text="ให้ตัวอักษรโชว์ในเกม แล้วกดปุ่ม Capture\nทำทีละตัว จนครบ",
                 bg=C["bg"], fg=C["dim"], font=("Segoe UI",10), justify="center").pack(pady=(0,12))

        # Current key label
        self._setup_idx = 0
        self._setup_keys = keys_to_capture
        self._setup_dlg = dlg

        # Key entry (custom keys)
        ef = tk.Frame(dlg, bg=C["bg"]); ef.pack(pady=4)
        tk.Label(ef, text="Keys to capture:", bg=C["bg"], fg=C["dim"], font=("Segoe UI",9)).pack(side=tk.LEFT)
        self._setup_entry = tk.Entry(ef, bg=C["input"], fg=C["cyan"], font=("Consolas",12),
                                      relief="flat", width=14, justify="center")
        self._setup_entry.insert(0, "qweasd")
        self._setup_entry.pack(side=tk.LEFT, padx=6)

        def apply_keys():
            txt = self._setup_entry.get().strip().lower()
            if txt:
                self._setup_keys = list(dict.fromkeys(txt))  # unique, keep order
                self._setup_idx = 0
                update_display()

        Btn(ef, text="Set", color=C["blue"], w=50, h=28, fs=9, cmd=apply_keys).pack(side=tk.LEFT)

        self.setup_current = tk.Label(dlg, text="", bg=C["bg"], fg=C["green"],
                                       font=("Orbitron" if os.name=="nt" else "Consolas", 48, "bold"))
        self.setup_current.pack(pady=8)

        self.setup_info = tk.Label(dlg, text="", bg=C["bg"], fg=C["dim"], font=("Segoe UI",10))
        self.setup_info.pack()

        # Preview of what's captured
        self.setup_preview = tk.Label(dlg, bg=C["bg2"], width=10, height=3)
        self.setup_preview.pack(pady=6)

        btn_frame = tk.Frame(dlg, bg=C["bg"]); btn_frame.pack(pady=6)

        def do_capture():
            if self._setup_idx >= len(self._setup_keys):
                return
            key = self._setup_keys[self._setup_idx]
            frame = self.capture.grab(self.region)
            if frame is None:
                self.log("Capture failed!", "err"); return

            fname = self.engine.save(key, frame)
            self.log(f"Saved template: '{key.upper()}' -> {fname}", "ok")
            self._update_templates()

            # Show preview
            try:
                img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
                img = img.resize((80, 50), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self.setup_preview.config(image=photo); self.setup_preview._photo = photo
            except: pass

            self._setup_idx += 1
            update_display()

        def update_display():
            if self._setup_idx >= len(self._setup_keys):
                self.setup_current.config(text="DONE!", fg=C["green"])
                self.setup_info.config(text=f"จับครบ {len(self._setup_keys)} ตัวแล้ว! ปิดหน้านี้ได้เลย")
                return
            key = self._setup_keys[self._setup_idx]
            self.setup_current.config(text=key.upper())
            self.setup_info.config(
                text=f"({self._setup_idx+1}/{len(self._setup_keys)})  ให้ตัว {key.upper()} โชว์ในเกม แล้วกด Capture")

        Btn(btn_frame, text="Capture", color=C["green"], w=140, h=40, fs=13, cmd=do_capture).pack(side=tk.LEFT, padx=4)
        Btn(btn_frame, text="Skip", color=C["dim"], w=80, h=40, fs=10,
            cmd=lambda: (setattr(self, '_setup_idx', self._setup_idx+1), update_display())).pack(side=tk.LEFT, padx=4)

        update_display()

    def _capture_one(self):
        if not self.region: messagebox.showwarning("QTE","Select region first!"); return
        dlg = tk.Toplevel(self.root); dlg.title("Capture Key"); dlg.geometry("280x150")
        dlg.configure(bg=C["bg"]); dlg.attributes("-topmost", True)
        tk.Label(dlg, text="Key on screen now?", bg=C["bg"], fg=C["white"], font=("Segoe UI",12)).pack(pady=(16,8))
        e = tk.Entry(dlg, bg=C["input"], fg=C["cyan"], font=("Consolas",22,"bold"),
                     justify="center", width=4, relief="flat")
        e.pack(pady=4); e.focus_set()
        def do(event=None):
            k = e.get().strip().lower()
            if not k or len(k)!=1: return
            frame = self.capture.grab(self.region)
            if frame is None: self.log("Capture failed!","err"); return
            fname = self.engine.save(k, frame)
            self.log(f"Saved: '{k.upper()}' -> {fname}", "ok")
            self._update_templates(); dlg.destroy()
        e.bind("<Return>", do)
        Btn(dlg, text="Capture!", color=C["purple"], w=120, h=34, fs=11, cmd=do).pack(pady=8)

    def _clear_templates(self):
        if not messagebox.askyesno("Clear","Delete all templates?"): return
        import shutil
        d = self.engine.dir
        if os.path.isdir(d): shutil.rmtree(d)
        os.makedirs(d, exist_ok=True)
        self.engine.load(); self._update_templates()
        self.log("All templates cleared", "warn")

    # ── Macro ──
    def toggle(self):
        if self.running: self._stop()
        else: self._start()

    def _start(self):
        if not self.region: messagebox.showwarning("QTE","Select region first! (F6)"); return
        if not self.capture.method: messagebox.showerror("QTE","Screen capture not working!"); return
        if self.engine.count == 0: messagebox.showwarning("QTE","No templates! Do Quick Setup first."); return

        self.frame_det = FrameDetector(self.v_change.get(), self.v_mincd.get())
        self.running = True; self.session_keys = 0; self.session_start = time.time()

        self.btn_start.set_text("STOP (F5)"); self.btn_start.set_color(C["red"])
        self.hdr.itemconfig(self.lbl_status, text="ON", fill=C["green"])
        self.log("Macro STARTED", "ok")
        self._save_cfg()
        threading.Thread(target=self._loop, daemon=True).start()

    def _stop(self):
        if not self.running: return
        self.running = False
        self.btn_start.set_text("START (F5)"); self.btn_start.set_color(C["green"])
        self.hdr.itemconfig(self.lbl_status, text="OFF", fill=C["dim"])
        e = time.time() - (self.session_start or time.time())
        self.log(f"Stopped  ({self.session_keys} keys / {e:.1f}s)")

    def _loop(self):
        thresh = self.v_thresh.get() / 100
        while self.running:
            try:
                frame = self.capture.grab(self.region)
                if frame is None:
                    time.sleep(0.05); continue

                # Convert gray ครั้งเดียว ใช้ทั้ง frame detect + template match
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                # 1. เช็คว่าหน้าจอเปลี่ยนหรือยัง
                small = cv2.resize(gray, (64, 64))
                self.frame_det.check_raw(small)

                # 2. ถ้ายังไม่ได้กด key จาก frame นี้ → ลอง detect ทุก frame
                #    (ไม่ใช่แค่ตอน frame เปลี่ยน เพราะ transition อาจ blur)
                if self.frame_det.can_press():
                    key, conf = self.engine.detect(gray, thresh)

                    if key:
                        kb.press_and_release(key)
                        self.frame_det.record()
                        self.session_keys += 1
                        self.root.after(0, self._on_key, key, conf)

                        if self.v_sound.get():
                            try:
                                import winsound; winsound.Beep(900, 20)
                            except: pass

            except Exception as e:
                self.root.after(0, self.log, str(e), "err")

            time.sleep(self.v_scan.get() / 1000.0)

    def _on_key(self, key, conf):
        self.lbl_count.config(text=str(self.session_keys))
        self.lbl_last.config(text=f"Last:  {key.upper()}")
        self.lbl_debug.config(text=f"Match: {key.upper()} ({conf:.0%})",
                               fg=C["green"] if conf > 0.85 else C["orange"])
        e = time.time() - (self.session_start or time.time())
        if e > 0: self.lbl_speed.config(text=f"{self.session_keys/e:.1f} /sec")

        # History
        self.key_history.append(key.upper())
        if len(self.key_history) > 12: self.key_history.pop(0)
        self.lbl_history.config(text=" ".join(self.key_history[-12:]))

        self.log(f"  [{self.session_keys}]  {key.upper()}  ({conf:.0%})", "key")

    # ── Loops ──
    def _preview_loop(self):
        if self.region and self.capture.method:
            try:
                img = self.capture.grab_pil(self.region)
                if img:
                    img = img.resize((105, 66), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    self.preview.config(image=photo, text=""); self.preview._photo = photo
            except: pass
        self.root.after(500, self._preview_loop)

    def _timer_loop(self):
        if self.running and self.session_start:
            e=int(time.time()-self.session_start); m,s=divmod(e,60)
            self.hdr.itemconfig(self.lbl_timer, text=f"{m:02d}:{s:02d}", fill=C["green"])
            # Live frame change %
            pct = self.frame_det.change_pct
            fc = C["green"] if pct > self.v_change.get() else C["dim"]
            self.lbl_frame.config(text=f"Frame: {pct:.1f}%  {'NEW' if not self.frame_det.pressed else 'done'}", fg=fc)
        else:
            self.hdr.itemconfig(self.lbl_timer, text="00:00", fill=C["dim2"])
            self.lbl_frame.config(text="Frame: -", fg=C["dim"])
        self.root.after(200, self._timer_loop)

    # ── Log ──
    def log(self, msg, tag=None):
        def _do():
            self.log_box.config(state=tk.NORMAL)
            ts = datetime.now().strftime("%H:%M:%S")
            if tag:
                self.log_box.insert(tk.END, f"[{ts}] ", "dim")
                self.log_box.insert(tk.END, msg+"\n", tag)
            else: self.log_box.insert(tk.END, f"[{ts}] {msg}\n")
            self.log_box.see(tk.END)
            lines = int(self.log_box.index("end-1c").split(".")[0])
            if lines > 200: self.log_box.delete("1.0", f"{lines-200}.0")
            self.log_box.config(state=tk.DISABLED)
        self.root.after(0, _do)

    def _quit(self):
        self.running = False; self._save_cfg(); self.capture.close()
        time.sleep(0.1); self.root.destroy()

    def run(self): self.root.mainloop()


if __name__ == "__main__":
    App().run()
