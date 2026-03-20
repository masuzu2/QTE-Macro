"""
QTE Minigame Macro - Ultimate Glass UI Edition
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
# Lib Check
# ============================================================
MISSING = []
for lib, pkg in [("pyautogui","pyautogui"),("keyboard","keyboard"),
                  ("mss","mss"),("cv2","opencv-python"),
                  ("numpy","numpy"),("pytesseract","pytesseract"),("PIL","Pillow")]:
    try: __import__(lib)
    except ImportError: MISSING.append(pkg)

if MISSING:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing",
        f"Install:\npip install {' '.join(MISSING)}")
    sys.exit(1)

import pyautogui, keyboard as kb, mss
import cv2, numpy as np, pytesseract
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageFilter

# ============================================================
# Color Palette - Cyberpunk Neon
# ============================================================
C = {
    "bg":       "#0a0e17",
    "bg2":      "#0f1923",
    "card":     "#111a2b",
    "card2":    "#162032",
    "border":   "#1e3050",
    "glow":     "#1a3a5c",
    "neon_blue":    "#00d4ff",
    "neon_purple":  "#a855f7",
    "neon_green":   "#22ff88",
    "neon_red":     "#ff4466",
    "neon_orange":  "#ff9f43",
    "neon_yellow":  "#ffe066",
    "neon_pink":    "#ff6bcb",
    "neon_cyan":    "#00e5ff",
    "white":    "#eaf0f6",
    "dim":      "#5a6a7e",
    "dim2":     "#3d4f63",
    "input":    "#0c1220",
}

# ============================================================
# Config
# ============================================================
CFG_PATH = os.path.join(get_base_path(), "macro_config.json")
DEF_CFG = {
    "scan_ms": 30, "cooldown_ms": 120, "mode": "ocr",
    "tesseract_path": r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    "region": None, "always_on_top": True, "sound": True,
    "keys": "qweasdzxcrtyfghvbn1234567890",
    "template_threshold": 0.78, "profiles": {},
}

def cfg_load():
    if os.path.exists(CFG_PATH):
        try:
            with open(CFG_PATH,"r",encoding="utf-8") as f:
                return {**DEF_CFG, **json.load(f)}
        except: pass
    return dict(DEF_CFG)

def cfg_save(c):
    try:
        with open(CFG_PATH,"w",encoding="utf-8") as f:
            json.dump(c, f, indent=2, ensure_ascii=False)
    except: pass

# ============================================================
# Custom Widgets
# ============================================================

class GlowButton(tk.Canvas):
    def __init__(self, parent, text="", color="#00d4ff", width=200, height=48,
                 font_size=13, command=None, **kw):
        super().__init__(parent, width=width, height=height,
                         bg=C["bg"], highlightthickness=0, **kw)
        self.w, self.h = width, height
        self.color = color
        self.text = text
        self.command = command
        self.font_size = font_size
        self.hover = False
        self._anim_alpha = 0.0
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self._draw()

    def _hex_to_rgb(self, h):
        h = h.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0,2,4))

    def _blend(self, c1, c2, t):
        r1,g1,b1 = self._hex_to_rgb(c1)
        r2,g2,b2 = self._hex_to_rgb(c2)
        return f"#{int(r1+(r2-r1)*t):02x}{int(g1+(g2-g1)*t):02x}{int(b1+(b2-b1)*t):02x}"

    def _draw(self):
        self.delete("all")
        w, h, r = self.w, self.h, 12
        if self.hover or self._anim_alpha > 0:
            gc = self._blend(C["bg"], self.color, min(self._anim_alpha, 0.15))
            self._rrect(0, 0, w, h, r+4, fill=gc, outline="")
        fill = self._blend(C["card"], self.color, 0.15 + self._anim_alpha * 0.15)
        border = self._blend(self.color, "#ffffff", self._anim_alpha * 0.3)
        self._rrect(2, 2, w-2, h-2, r, fill=fill, outline=border, width=2)
        tc = self._blend(self.color, "#ffffff", self._anim_alpha * 0.5)
        self.create_text(w//2, h//2, text=self.text, fill=tc,
                         font=("Segoe UI", self.font_size, "bold"))

    def _rrect(self, x1, y1, x2, y2, r, **kw):
        pts = [x1+r,y1, x2-r,y1, x2,y1, x2,y1+r, x2,y2-r, x2,y2,
               x2-r,y2, x1+r,y2, x1,y2, x1,y2-r, x1,y1+r, x1,y1]
        return self.create_polygon(pts, smooth=True, **kw)

    def _on_enter(self, e):
        self.hover = True; self._animate_in()
    def _on_leave(self, e):
        self.hover = False; self._animate_out()
    def _animate_in(self):
        if self._anim_alpha < 1.0 and self.hover:
            self._anim_alpha = min(self._anim_alpha + 0.15, 1.0)
            self._draw(); self.after(16, self._animate_in)
    def _animate_out(self):
        if self._anim_alpha > 0 and not self.hover:
            self._anim_alpha = max(self._anim_alpha - 0.1, 0)
            self._draw(); self.after(16, self._animate_out)
    def _on_click(self, e):
        if self.command: self.command()
    def set_text(self, t): self.text = t; self._draw()
    def set_color(self, c): self.color = c; self._draw()


class NeonCard(tk.Frame):
    def __init__(self, parent, glow_color=None, **kw):
        super().__init__(parent, bg=C["card"],
                         highlightbackground=glow_color or C["border"],
                         highlightthickness=1, **kw)


class PulsingDot(tk.Canvas):
    def __init__(self, parent, size=16, color="#555", **kw):
        super().__init__(parent, width=size+10, height=size+10,
                         bg=C["bg"], highlightthickness=0, **kw)
        self.size, self.color, self.active, self._phase = size, color, False, 0.0
        self._draw_dot()

    def _hex_to_rgb(self, h):
        h = h.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0,2,4))

    def _draw_dot(self):
        self.delete("all")
        cx, cy = (self.size+10)//2, (self.size+10)//2
        r = self.size // 2
        if self.active:
            pulse = 0.3 + 0.7 * abs(math.sin(self._phase))
            pr = r + 3 + int(pulse * 3)
            rgb = self._hex_to_rgb(self.color)
            ah = f"#{int(rgb[0]*pulse):02x}{int(rgb[1]*pulse):02x}{int(rgb[2]*pulse):02x}"
            self.create_oval(cx-pr, cy-pr, cx+pr, cy+pr, fill="", outline=ah, width=1)
        self.create_oval(cx-r, cy-r, cx+r, cy+r, fill=self.color, outline="")

    def set_active(self, active, color=None):
        self.active = active
        if color: self.color = color
        if active: self._animate()
        else: self._draw_dot()

    def _animate(self):
        if self.active:
            self._phase += 0.12; self._draw_dot(); self.after(33, self._animate)


class KeyGrid(tk.Canvas):
    def __init__(self, parent, width=480, height=130, **kw):
        super().__init__(parent, width=width, height=height,
                         bg=C["card"], highlightthickness=0, **kw)
        self.w, self.h = width, height
        self.pressed, self.key_stats = {}, {}
        self._draw()

    def _draw(self):
        self.delete("all")
        rows = [list("qwertyuiop"), list("asdfghjkl"), list("zxcvbnm")]
        kw, kh, gap = 38, 38, 4
        now = time.time()
        total_h = len(rows) * (kh + gap)
        sy = (self.h - total_h) // 2
        for ri, row in enumerate(rows):
            tw = len(row) * (kw + gap) - gap
            sx = (self.w - tw) // 2
            for ci, key in enumerate(row):
                x, y = sx + ci*(kw+gap), sy + ri*(kh+gap)
                age = now - self.pressed.get(key, 0)
                if age < 1.0:
                    intensity = 1.0 - age
                    gr = int(intensity * 8)
                    self._rrect(x-gr, y-gr, x+kw+gr, y+kh+gr, 8,
                        fill=self._fade(C["neon_cyan"], intensity*0.3), outline="")
                    fill = self._fade(C["neon_cyan"], 0.3+intensity*0.4)
                    border, tc = C["neon_cyan"], "#ffffff"
                else:
                    fill, border, tc = C["card2"], C["border"], C["dim"]
                self._rrect(x, y, x+kw, y+kh, 6, fill=fill, outline=border, width=1)
                self.create_text(x+kw//2, y+kh//2, text=key.upper(), fill=tc,
                                font=("Segoe UI", 11, "bold"))
                cnt = self.key_stats.get(key, 0)
                if cnt > 0:
                    self.create_text(x+kw-4, y+4, text=str(cnt), fill=C["neon_orange"],
                                    font=("Consolas", 7), anchor="ne")

    def _fade(self, hc, a):
        r,g,b = int(hc[1:3],16), int(hc[3:5],16), int(hc[5:7],16)
        br,bg_,bb = int(C["card"][1:3],16), int(C["card"][3:5],16), int(C["card"][5:7],16)
        return f"#{max(0,min(255,int(br+(r-br)*a))):02x}{max(0,min(255,int(bg_+(g-bg_)*a))):02x}{max(0,min(255,int(bb+(b-bb)*a))):02x}"

    def _rrect(self, x1, y1, x2, y2, r, **kw):
        pts = [x1+r,y1, x2-r,y1, x2,y1, x2,y1+r, x2,y2-r, x2,y2, x2-r,y2, x1+r,y2, x1,y2, x1,y2-r, x1,y1+r, x1,y1]
        return self.create_polygon(pts, smooth=True, **kw)

    def flash_key(self, key):
        key = key.lower()
        self.pressed[key] = time.time()
        self.key_stats[key] = self.key_stats.get(key, 0) + 1
        self._draw(); self.after(1000, self._draw)

    def reset(self):
        self.pressed.clear(); self.key_stats.clear(); self._draw()


class MeterBar(tk.Canvas):
    def __init__(self, parent, width=200, height=6, color=C["neon_blue"], **kw):
        super().__init__(parent, width=width, height=height,
                         bg=C["card"], highlightthickness=0, **kw)
        self.w, self.h, self.color, self.value = width, height, color, 0.0
        self._draw()
    def _draw(self):
        self.delete("all")
        self.create_rectangle(0,0,self.w,self.h, fill=C["bg2"], outline="")
        fw = int(self.w * min(self.value, 1.0))
        if fw > 0: self.create_rectangle(0,0,fw,self.h, fill=self.color, outline="")
    def set_value(self, v):
        self.value = max(0, min(v, 1.0)); self._draw()


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
        self.attributes("-alpha", 0.30)
        self.state("zoomed")
        self.configure(bg="black")
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0")
        self.canvas = tk.Canvas(self, bg="black", highlightthickness=0, cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.create_text(sw//2, 36, text="DRAG TO SELECT AREA",
            fill="#00ffaa", font=("Segoe UI", 20, "bold"))
        self.canvas.create_text(sw//2, 66, text="Select where the QTE key appears  |  ESC = cancel",
            fill="#aaaaaa", font=("Segoe UI", 12))
        self.hline = self.canvas.create_line(0,0,0,0, fill="#00ffaa33", dash=(4,4))
        self.vline = self.canvas.create_line(0,0,0,0, fill="#00ffaa33", dash=(4,4))
        self.canvas.bind("<Motion>", self._motion)
        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._drag)
        self.canvas.bind("<ButtonRelease-1>", self._release)
        self.bind("<Escape>", lambda e: self.destroy())
        self.focus_force()

    def _motion(self, e):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.canvas.coords(self.hline, 0, e.y, sw, e.y)
        self.canvas.coords(self.vline, e.x, 0, e.x, sh)
    def _press(self, e):
        self.sx, self.sy = e.x, e.y
        if self.rect: self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(e.x,e.y,e.x,e.y, outline="#00ffaa", width=2, dash=(6,3))
    def _drag(self, e):
        self.canvas.coords(self.rect, self.sx, self.sy, e.x, e.y)
        self.canvas.delete("sz")
        w, h = abs(e.x-self.sx), abs(e.y-self.sy)
        self.canvas.create_text((self.sx+e.x)//2, min(self.sy,e.y)-18,
            text=f"{w} x {h} px", fill="#00ffaa", font=("Consolas", 13, "bold"), tags="sz")
    def _release(self, e):
        x1,y1 = min(self.sx,e.x), min(self.sy,e.y)
        x2,y2 = max(self.sx,e.x), max(self.sy,e.y)
        if (x2-x1)>10 and (y2-y1)>10:
            self.callback({"left":x1,"top":y1,"width":x2-x1,"height":y2-y1})
        self.destroy()


# ============================================================
# Template Matcher
# ============================================================
class TemplateMatcher:
    def __init__(self, d):
        self.templates, self.template_dir = {}, d
        os.makedirs(d, exist_ok=True); self.load()
    def load(self):
        self.templates.clear()
        if not os.path.isdir(self.template_dir): return
        for f in os.listdir(self.template_dir):
            if f.lower().endswith(".png"):
                k = f.rsplit(".",1)[0].lower()
                img = cv2.imread(os.path.join(self.template_dir,f), cv2.IMREAD_GRAYSCALE)
                if img is not None: self.templates[k] = img
    def detect(self, frame, thresh=0.78):
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        best, bv = None, 0
        for k, t in self.templates.items():
            if t.shape[0]>gray.shape[0] or t.shape[1]>gray.shape[1]: continue
            res = cv2.matchTemplate(gray, t, cv2.TM_CCOEFF_NORMED)
            _,mx,_,_ = cv2.minMaxLoc(res)
            if mx>thresh and mx>bv: best,bv = k,mx
        return best, bv
    def save_template(self, key, region):
        with mss.mss() as sct:
            shot = sct.grab(region)
            img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
            p = os.path.join(self.template_dir, f"{key}.png"); img.save(p)
            self.templates[key] = cv2.imread(p, cv2.IMREAD_GRAYSCALE)
    @property
    def count(self): return len(self.templates)


# ============================================================
# Main App
# ============================================================
class MacroApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QTE Macro")
        self.root.geometry("540x880")
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)
        self.cfg = cfg_load()
        self.running = False
        self.region = self.cfg.get("region")
        self.last_press = 0
        self.session_keys = 0
        self.session_start = None
        self.total_keys = 0
        self.last_detected = "-"
        self._images = []

        self.v_mode = tk.StringVar(value=self.cfg["mode"])
        self.v_scan = tk.IntVar(value=self.cfg["scan_ms"])
        self.v_cool = tk.IntVar(value=self.cfg["cooldown_ms"])
        self.v_ontop = tk.BooleanVar(value=self.cfg.get("always_on_top", True))
        self.v_sound = tk.BooleanVar(value=self.cfg.get("sound", True))
        self.v_tess = tk.StringVar(value=self.cfg["tesseract_path"])
        self.v_keys = tk.StringVar(value=self.cfg["keys"])
        self.v_thresh = tk.DoubleVar(value=self.cfg.get("template_threshold", 0.78))
        self.valid_keys = set(self.cfg["keys"])
        self.matcher = TemplateMatcher(os.path.join(get_base_path(), "templates"))

        self._build()
        self._apply_ontop()
        self._update_region_label()
        self._preview_loop()
        self._clock_loop()

        self.root.bind_all("<F5>", lambda e: self.toggle())
        self.root.bind_all("<F6>", lambda e: self.pick_region())
        self.root.bind_all("<Escape>", lambda e: self._stop())
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        hdr = tk.Canvas(self.root, bg=C["bg"], height=80, highlightthickness=0)
        hdr.pack(fill=tk.X)
        for i in range(540):
            t = i/540
            r,g,b = int(0+(168)*t), int(212+(85-212)*t), int(255+(247-255)*t)
            hdr.create_line(i,0,i,3, fill=f"#{r:02x}{g:02x}{b:02x}")
        hdr.create_text(24, 32, text="QTE MACRO", anchor="w",
                        fill=C["neon_blue"], font=("Segoe UI", 24, "bold"))
        hdr.create_text(24, 56, text="Ultimate Minigame Macro", anchor="w",
                        fill=C["dim"], font=("Segoe UI", 10))
        self.dot = PulsingDot(hdr, size=12, color="#444")
        hdr.create_window(470, 30, window=self.dot)
        self.lbl_status = hdr.create_text(490, 30, text="OFFLINE", anchor="w",
                                           fill=C["dim"], font=("Segoe UI", 10, "bold"))
        self.hdr = hdr
        self.lbl_timer = hdr.create_text(516, 50, text="00:00", anchor="e",
                                          fill=C["dim2"], font=("Consolas", 9))

        main = tk.Frame(self.root, bg=C["bg"])
        main.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        self.btn_start = GlowButton(main, text="START", color=C["neon_green"],
                                      width=508, height=56, font_size=18, command=self.toggle)
        self.btn_start.pack(pady=(4, 10))

        stats_card = NeonCard(main, glow_color=C["glow"])
        stats_card.pack(fill=tk.X, pady=(0, 8))
        si = tk.Frame(stats_card, bg=C["card"]); si.pack(fill=tk.X, padx=16, pady=12)

        left = tk.Frame(si, bg=C["card"]); left.pack(side=tk.LEFT)
        self.lbl_count = tk.Label(left, text="0", bg=C["card"], fg=C["neon_green"],
                                   font=("Consolas", 36, "bold"))
        self.lbl_count.pack(anchor="w")
        tk.Label(left, text="KEYS PRESSED", bg=C["card"], fg=C["dim"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")

        mid = tk.Frame(si, bg=C["card"]); mid.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=20)
        self.lbl_speed = tk.Label(mid, text="0.0 /sec", bg=C["card"], fg=C["neon_blue"],
                                   font=("Consolas", 13, "bold"), anchor="w")
        self.lbl_speed.pack(anchor="w")
        self.speed_meter = MeterBar(mid, width=160, height=4, color=C["neon_blue"])
        self.speed_meter.pack(anchor="w", pady=(2, 6))
        self.lbl_last = tk.Label(mid, text="Last: -", bg=C["card"], fg=C["neon_purple"],
                                  font=("Segoe UI", 11, "bold"), anchor="w")
        self.lbl_last.pack(anchor="w")

        pf = tk.Frame(si, bg=C["bg2"], width=100, height=66,
                       highlightbackground=C["border"], highlightthickness=1)
        pf.pack(side=tk.RIGHT); pf.pack_propagate(False)
        self.preview_lbl = tk.Label(pf, bg=C["bg2"], text="Preview",
                                     fg=C["dim2"], font=("Segoe UI", 8))
        self.preview_lbl.pack(expand=True)

        self._sec(main, "KEY ACTIVITY")
        self.key_grid = KeyGrid(main, width=508, height=130); self.key_grid.pack(pady=(0,8))

        self._sec(main, "CAPTURE REGION")
        rc = NeonCard(main); rc.pack(fill=tk.X, pady=(0,8))
        ri = tk.Frame(rc, bg=C["card"]); ri.pack(fill=tk.X, padx=12, pady=10)
        GlowButton(ri, text="Select Region (F6)", color=C["neon_blue"],
                    width=180, height=36, font_size=10, command=self.pick_region).pack(side=tk.LEFT)
        self.lbl_region = tk.Label(ri, text="Not set", bg=C["card"], fg=C["neon_orange"],
                                    font=("Consolas", 10)); self.lbl_region.pack(side=tk.RIGHT)

        self._sec(main, "DETECTION MODE")
        mc = NeonCard(main); mc.pack(fill=tk.X, pady=(0,8))
        mi = tk.Frame(mc, bg=C["card"]); mi.pack(fill=tk.X, padx=12, pady=10)
        for val, icon, name, desc in [
            ("ocr","A","OCR","Tesseract reads letters"),
            ("template","T","Template","Image matching (accurate)")]:
            f = tk.Frame(mi, bg=C["card"]); f.pack(fill=tk.X, pady=1)
            tk.Radiobutton(f, text=f"  {icon}   {name}", variable=self.v_mode, value=val,
                           bg=C["card"], fg=C["white"], selectcolor=C["bg"],
                           activebackground=C["card"], activeforeground=C["neon_blue"],
                           font=("Segoe UI", 11), anchor="w",
                           command=self._on_mode_change).pack(side=tk.LEFT)
            tk.Label(f, text=desc, bg=C["card"], fg=C["dim"],
                     font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=8)

        self.f_ocr = tk.Frame(mc, bg=C["card"])
        oc = tk.Frame(self.f_ocr, bg=C["card"]); oc.pack(fill=tk.X, padx=12, pady=(0,8))
        tk.Label(oc, text="Tesseract:", bg=C["card"], fg=C["dim"], font=("Segoe UI",9)).pack(side=tk.LEFT)
        tk.Entry(oc, textvariable=self.v_tess, bg=C["input"], fg=C["white"],
                 insertbackground=C["white"], font=("Consolas",9), relief="flat", width=38).pack(side=tk.LEFT, padx=6)

        self.f_tmpl = tk.Frame(mc, bg=C["card"])
        tc = tk.Frame(self.f_tmpl, bg=C["card"]); tc.pack(fill=tk.X, padx=12, pady=(0,8))
        GlowButton(tc, text="Capture Key", color=C["neon_purple"], width=120, height=30,
                    font_size=9, command=self._capture_template).pack(side=tk.LEFT, padx=(0,6))
        GlowButton(tc, text="Load Folder", color=C["dim"], width=110, height=30,
                    font_size=9, command=self._load_templates).pack(side=tk.LEFT)
        self.lbl_tmpl_n = tk.Label(tc, text=f"{self.matcher.count} templates",
                                    bg=C["card"], fg=C["dim"], font=("Segoe UI",9))
        self.lbl_tmpl_n.pack(side=tk.RIGHT)
        self._on_mode_change()

        self._sec(main, "SETTINGS")
        sf = NeonCard(main); sf.pack(fill=tk.X, pady=(0,8))
        sc = tk.Frame(sf, bg=C["card"]); sc.pack(fill=tk.X, padx=12, pady=10)
        self._slider(sc, "Scan Speed", "ms", self.v_scan, 10, 200, C["neon_cyan"])
        self._slider(sc, "Key Cooldown", "ms", self.v_cool, 30, 500, C["neon_orange"])
        kr = tk.Frame(sc, bg=C["card"]); kr.pack(fill=tk.X, pady=4)
        tk.Label(kr, text="Active Keys:", bg=C["card"], fg=C["dim"], font=("Segoe UI",9)).pack(side=tk.LEFT)
        tk.Entry(kr, textvariable=self.v_keys, bg=C["input"], fg=C["neon_blue"],
                 insertbackground=C["white"], font=("Consolas",10), relief="flat", width=30).pack(side=tk.RIGHT)
        chk = tk.Frame(sc, bg=C["card"]); chk.pack(fill=tk.X, pady=(4,0))
        tk.Checkbutton(chk, text="Always on Top", variable=self.v_ontop, bg=C["card"], fg=C["white"],
                       selectcolor=C["bg"], activebackground=C["card"], font=("Segoe UI",9),
                       command=self._apply_ontop).pack(side=tk.LEFT, padx=(0,12))
        tk.Checkbutton(chk, text="Sound Feedback", variable=self.v_sound, bg=C["card"], fg=C["white"],
                       selectcolor=C["bg"], activebackground=C["card"], font=("Segoe UI",9)).pack(side=tk.LEFT)

        self._sec(main, "PROFILES")
        pfc = NeonCard(main); pfc.pack(fill=tk.X, pady=(0,8))
        pfi = tk.Frame(pfc, bg=C["card"]); pfi.pack(fill=tk.X, padx=12, pady=10)
        GlowButton(pfi, text="Save", color=C["neon_blue"], width=80, height=30,
                    font_size=9, command=self._save_profile).pack(side=tk.LEFT, padx=(0,6))
        GlowButton(pfi, text="Load", color=C["dim"], width=80, height=30,
                    font_size=9, command=self._load_profile).pack(side=tk.LEFT)
        self.lbl_profile = tk.Label(pfi, text="No profile", bg=C["card"], fg=C["dim"],
                                     font=("Segoe UI",9)); self.lbl_profile.pack(side=tk.RIGHT)

        self._sec(main, "LOG")
        lc = NeonCard(main); lc.pack(fill=tk.BOTH, expand=True, pady=(0,4))
        self.log_text = tk.Text(lc, bg=C["bg"], fg=C["neon_green"], font=("Consolas",9),
                                 height=4, relief="flat", wrap=tk.WORD, state=tk.DISABLED, padx=10, pady=6)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self.log_text.tag_configure("key", foreground=C["neon_cyan"], font=("Consolas",9,"bold"))
        self.log_text.tag_configure("err", foreground=C["neon_red"])
        self.log_text.tag_configure("info", foreground=C["dim"])

        ft = tk.Frame(self.root, bg=C["bg2"], height=28); ft.pack(fill=tk.X, side=tk.BOTTOM)
        ft.pack_propagate(False)
        fl = tk.Canvas(self.root, bg=C["bg"], height=2, highlightthickness=0)
        fl.pack(fill=tk.X, side=tk.BOTTOM)
        for i in range(540):
            t = i/540; r,g,b = int(168-168*t), int(85+(212-85)*t), int(247+(255-247)*t)
            fl.create_line(i,0,i,2, fill=f"#{r:02x}{g:02x}{b:02x}")
        tk.Label(ft, text="F5 Start/Stop   F6 Region   ESC Stop",
                 bg=C["bg2"], fg=C["dim2"], font=("Segoe UI",8)).pack(pady=5)

    def _sec(self, p, t):
        f = tk.Frame(p, bg=C["bg"]); f.pack(fill=tk.X, pady=(6,3))
        tk.Label(f, text=t, bg=C["bg"], fg=C["dim2"], font=("Segoe UI",8,"bold")).pack(side=tk.LEFT, padx=2)
        s = tk.Canvas(f, bg=C["bg"], height=1, highlightthickness=0)
        s.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8,0), pady=1)
        s.create_line(0,0,400,0, fill=C["border"])

    def _slider(self, p, label, unit, var, lo, hi, color):
        row = tk.Frame(p, bg=C["card"]); row.pack(fill=tk.X, pady=3)
        tk.Label(row, text=f"{label}:", bg=C["card"], fg=C["white"], font=("Segoe UI",10)).pack(side=tk.LEFT)
        vl = tk.Label(row, text=f"{var.get()}{unit}", bg=C["card"], fg=color,
                       font=("Consolas",10,"bold"), width=7, anchor="e"); vl.pack(side=tk.RIGHT)
        tk.Scale(row, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var,
                 bg=C["card"], fg=C["card"], troughcolor=C["bg"], highlightthickness=0,
                 showvalue=False, length=180, sliderlength=14, activebackground=color,
                 command=lambda v, l=vl, u=unit: l.config(text=f"{int(float(v))}{u}")
                 ).pack(side=tk.RIGHT, padx=(8,4))

    # === Region ===
    def pick_region(self):
        if self.running: self.log("Stop macro first!","err"); return
        self.root.iconify(); self.root.after(300, self._open_sel)
    def _open_sel(self):
        def done(r):
            self.region = r; self.cfg["region"] = r; cfg_save(self.cfg)
            self.root.deiconify(); self._update_region_label()
            self.log(f"Region: {r['width']}x{r['height']} @ ({r['left']},{r['top']})")
        sel = RegionSelector(self.root, done)
        sel.protocol("WM_DELETE_WINDOW", lambda: (sel.destroy(), self.root.deiconify()))
    def _update_region_label(self):
        if self.region:
            r = self.region
            self.lbl_region.config(text=f"{r['width']}x{r['height']}  ({r['left']},{r['top']})", fg=C["neon_green"])
        else: self.lbl_region.config(text="Not set", fg=C["neon_orange"])

    # === Mode ===
    def _on_mode_change(self):
        if self.v_mode.get()=="ocr": self.f_tmpl.pack_forget(); self.f_ocr.pack(fill=tk.X)
        else: self.f_ocr.pack_forget(); self.f_tmpl.pack(fill=tk.X)

    def _capture_template(self):
        if not self.region: messagebox.showwarning("QTE","Select region first!"); return
        dlg = tk.Toplevel(self.root); dlg.title("Capture Template"); dlg.geometry("300x180")
        dlg.configure(bg=C["bg"]); dlg.attributes("-topmost", True); dlg.resizable(False, False)
        tk.Label(dlg, text="Which key is on screen now?", bg=C["bg"], fg=C["white"],
                 font=("Segoe UI",12)).pack(pady=(20,8))
        entry = tk.Entry(dlg, bg=C["input"], fg=C["neon_cyan"], insertbackground=C["white"],
                         font=("Consolas",22,"bold"), justify="center", width=4, relief="flat")
        entry.pack(pady=6); entry.focus_set()
        def do(event=None):
            k = entry.get().strip().lower()
            if not k or len(k)!=1: return
            self.matcher.save_template(k, self.region)
            self.lbl_tmpl_n.config(text=f"{self.matcher.count} templates")
            self.log(f"Template saved: '{k}'"); dlg.destroy()
        entry.bind("<Return>", do)
        GlowButton(dlg, text="Capture!", color=C["neon_purple"], width=140, height=36,
                    font_size=11, command=do).pack(pady=8)

    def _load_templates(self):
        d = filedialog.askdirectory(title="Templates folder")
        if d:
            self.matcher.template_dir = d; self.matcher.load()
            self.lbl_tmpl_n.config(text=f"{self.matcher.count} templates")
            self.log(f"Loaded {self.matcher.count} templates")

    # === Profiles ===
    def _save_profile(self):
        dlg = tk.Toplevel(self.root); dlg.title("Save Profile"); dlg.geometry("280x130")
        dlg.configure(bg=C["bg"]); dlg.attributes("-topmost",True)
        tk.Label(dlg, text="Profile name:", bg=C["bg"], fg=C["white"], font=("Segoe UI",11)).pack(pady=(14,6))
        e = tk.Entry(dlg, bg=C["input"], fg=C["neon_blue"], insertbackground=C["white"],
                     font=("Segoe UI",12), justify="center", width=18, relief="flat"); e.pack(); e.focus_set()
        def save(event=None):
            n = e.get().strip()
            if not n: return
            self.cfg["profiles"][n] = {"scan_ms":self.v_scan.get(),"cooldown_ms":self.v_cool.get(),
                "mode":self.v_mode.get(),"region":self.region,"keys":self.v_keys.get(),
                "tesseract_path":self.v_tess.get()}
            cfg_save(self.cfg); self.lbl_profile.config(text=n, fg=C["neon_blue"])
            self.log(f"Profile saved: {n}"); dlg.destroy()
        e.bind("<Return>", save)
        GlowButton(dlg, text="Save", color=C["neon_green"], width=100, height=30, font_size=10, command=save).pack(pady=8)

    def _load_profile(self):
        profiles = self.cfg.get("profiles", {})
        if not profiles: messagebox.showinfo("Profiles","No profiles saved!"); return
        dlg = tk.Toplevel(self.root); dlg.title("Load Profile"); dlg.geometry("280x280")
        dlg.configure(bg=C["bg"]); dlg.attributes("-topmost",True)
        lb = tk.Listbox(dlg, bg=C["card"], fg=C["white"], selectbackground=C["neon_blue"],
                         font=("Segoe UI",11), relief="flat", height=8)
        lb.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)
        for n in profiles: lb.insert(tk.END, n)
        def load(event=None):
            sel = lb.curselection()
            if not sel: return
            n = lb.get(sel[0]); p = profiles[n]
            self.v_scan.set(p.get("scan_ms",30)); self.v_cool.set(p.get("cooldown_ms",120))
            self.v_mode.set(p.get("mode","ocr")); self.v_keys.set(p.get("keys",self.cfg["keys"]))
            self.v_tess.set(p.get("tesseract_path",self.cfg["tesseract_path"]))
            if p.get("region"): self.region = p["region"]; self._update_region_label()
            self._on_mode_change(); self.lbl_profile.config(text=n, fg=C["neon_blue"])
            self.log(f"Profile: {n}"); dlg.destroy()
        lb.bind("<Double-1>", load)
        GlowButton(dlg, text="Load", color=C["neon_blue"], width=100, height=30, font_size=10, command=load).pack(pady=(0,10))

    # === Macro ===
    def toggle(self):
        if self.running: self._stop()
        else: self._start()

    def _start(self):
        if not self.region: messagebox.showwarning("QTE","Select region first! (F6)"); return
        if self.v_mode.get()=="template" and self.matcher.count==0:
            messagebox.showwarning("QTE","No templates!"); return
        if self.v_mode.get()=="ocr":
            pytesseract.pytesseract.tesseract_cmd = self.v_tess.get()
        self.valid_keys = set(self.v_keys.get())
        self.running = True; self.session_keys = 0; self.session_start = time.time()
        self.key_grid.reset()
        self.btn_start.set_text("STOP"); self.btn_start.set_color(C["neon_red"])
        self.dot.set_active(True, C["neon_green"])
        self.hdr.itemconfig(self.lbl_status, text="ACTIVE", fill=C["neon_green"])
        self.log("Macro STARTED"); self._save_cfg()
        threading.Thread(target=self._loop, daemon=True).start()

    def _stop(self):
        if not self.running: return
        self.running = False
        self.btn_start.set_text("START"); self.btn_start.set_color(C["neon_green"])
        self.dot.set_active(False, "#555")
        self.hdr.itemconfig(self.lbl_status, text="OFFLINE", fill=C["dim"])
        elapsed = time.time() - (self.session_start or time.time())
        self.log(f"Stopped ({self.session_keys} keys / {elapsed:.1f}s)")

    def _loop(self):
        sct = mss.mss()
        while self.running:
            try:
                shot = sct.grab(self.region); frame = np.array(shot)
                key, conf = (None, 0)
                if self.v_mode.get()=="ocr": key, conf = self._detect_ocr(frame)
                else:
                    key, conf = self.matcher.detect(frame, self.v_thresh.get())
                    if key is None: conf = 0
                now = time.time(); cd = self.v_cool.get()/1000.0
                if key and key in self.valid_keys and (now-self.last_press)>cd:
                    kb.press_and_release(key); self.last_press = now
                    self.session_keys += 1; self.total_keys += 1
                    self.root.after(0, self._on_key, key)
                    if self.v_sound.get():
                        try: import winsound; winsound.Beep(900, 25)
                        except: pass
            except Exception as e:
                self.root.after(0, self.log, str(e), "err")
            time.sleep(self.v_scan.get()/1000.0)
        sct.close()

    def _detect_ocr(self, frame):
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        gray = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
        gray = cv2.convertScaleAbs(gray, alpha=2.0, beta=0)
        thresh = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,11,2)
        inv = cv2.bitwise_not(thresh)
        wl = self.v_keys.get()+self.v_keys.get().upper()
        cfg = f'--psm 10 -c tessedit_char_whitelist={wl}'
        for img in [thresh, inv]:
            txt = pytesseract.image_to_string(Image.fromarray(img), config=cfg).strip().lower()
            for ch in txt:
                if ch in self.valid_keys: return ch, 100
        return None, 0

    def _on_key(self, key):
        self.lbl_count.config(text=str(self.session_keys))
        self.lbl_last.config(text=f"Last:  {key.upper()}")
        elapsed = time.time() - (self.session_start or time.time())
        if elapsed > 0:
            kps = self.session_keys / elapsed
            self.lbl_speed.config(text=f"{kps:.1f} /sec")
            self.speed_meter.set_value(min(kps/5.0, 1.0))
        self.key_grid.flash_key(key)
        self.log(f"  [{self.session_keys}]  {key.upper()}", "key")

    # === Loops ===
    def _preview_loop(self):
        if self.region:
            try:
                with mss.mss() as sct:
                    shot = sct.grab(self.region)
                    img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
                    img = img.resize((96, 62), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    self.preview_lbl.config(image=photo, text="")
                    self.preview_lbl._photo = photo
            except: pass
        self.root.after(400, self._preview_loop)

    def _clock_loop(self):
        if self.running and self.session_start:
            e = int(time.time()-self.session_start); m,s = divmod(e,60)
            self.hdr.itemconfig(self.lbl_timer, text=f"{m:02d}:{s:02d}", fill=C["neon_green"])
        else:
            self.hdr.itemconfig(self.lbl_timer, text="00:00", fill=C["dim2"])
        self.root.after(1000, self._clock_loop)

    # === Misc ===
    def log(self, msg, tag=None):
        def _do():
            self.log_text.config(state=tk.NORMAL)
            ts = datetime.now().strftime("%H:%M:%S")
            if tag:
                self.log_text.insert(tk.END, f"[{ts}] ", "info")
                self.log_text.insert(tk.END, msg+"\n", tag)
            else: self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
            self.log_text.see(tk.END)
            lines = int(self.log_text.index("end-1c").split(".")[0])
            if lines > 300: self.log_text.delete("1.0", f"{lines-300}.0")
            self.log_text.config(state=tk.DISABLED)
        self.root.after(0, _do)

    def _apply_ontop(self): self.root.attributes("-topmost", self.v_ontop.get())

    def _save_cfg(self):
        self.cfg.update({"scan_ms":self.v_scan.get(),"cooldown_ms":self.v_cool.get(),
            "mode":self.v_mode.get(),"region":self.region,"keys":self.v_keys.get(),
            "tesseract_path":self.v_tess.get(),"always_on_top":self.v_ontop.get(),"sound":self.v_sound.get()})
        cfg_save(self.cfg)

    def _quit(self):
        self.running = False; self._save_cfg(); time.sleep(0.1); self.root.destroy()

    def run(self): self.root.mainloop()

if __name__ == "__main__":
    MacroApp().run()
