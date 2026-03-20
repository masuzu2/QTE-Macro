"""
QTE Macro v5 - Zero Setup
เลือก region → START → จบ ไม่ต้อง setup อะไรเลย
ระบบสร้าง template ตัวอักษร A-Z 0-9 อัตโนมัติจาก font
"""

import time, sys, os, json, threading
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

MISSING = []
for lib, pkg in [("keyboard","keyboard"),("mss","mss"),
                  ("cv2","opencv-python"),("numpy","numpy"),("PIL","Pillow")]:
    try: __import__(lib)
    except ImportError: MISSING.append(pkg)
if MISSING:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing", f"pip install {' '.join(MISSING)}")
    sys.exit(1)

import keyboard as kb
import mss
import cv2, numpy as np
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageGrab


# ═══════════════════════════════════════════════════
# Screen Capture
# ═══════════════════════════════════════════════════
class Cap:
    def __init__(self):
        self.sct = None
        try:
            self.sct = mss.mss()
            self.sct.grab(self.sct.monitors[0])
            self.ok = True
        except:
            self.ok = False

    def grab(self, region):
        try:
            if self.sct:
                s = self.sct.grab(region)
                return cv2.cvtColor(np.array(s), cv2.COLOR_BGRA2BGR)
            else:
                img = ImageGrab.grab(bbox=(region["left"], region["top"],
                    region["left"]+region["width"], region["top"]+region["height"]))
                return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        except:
            return None

    def grab_pil(self, region):
        try:
            if self.sct:
                s = self.sct.grab(region)
                return Image.frombytes("RGB", s.size, s.bgra, "raw", "BGRX")
            else:
                return ImageGrab.grab(bbox=(region["left"], region["top"],
                    region["left"]+region["width"], region["top"]+region["height"]))
        except:
            return None


# ═══════════════════════════════════════════════════
# Auto Character Templates - สร้างเองจาก font ไม่ต้อง setup
# ═══════════════════════════════════════════════════
class AutoTemplates:
    """สร้าง template ตัวอักษร A-Z 0-9 อัตโนมัติ หลายขนาด หลาย font"""

    def __init__(self):
        self.templates = {}  # char -> list of (gray_image)
        self._generate()

    def _generate(self):
        chars = "qweasdzxcrtyfghvbn1234567890"
        sizes = [28, 36, 44, 52, 64, 80]

        # หา font ที่มีในเครื่อง
        fonts = self._find_fonts()

        for ch in chars:
            self.templates[ch] = []
            for font in fonts:
                for sz in sizes:
                    img = self._render_char(ch.upper(), font, sz)
                    if img is not None:
                        self.templates[ch].append(img)
                    # ตัวเล็กด้วย
                    img2 = self._render_char(ch.lower(), font, sz)
                    if img2 is not None:
                        self.templates[ch].append(img2)

    def _find_fonts(self):
        """หา font ที่ใช้ได้"""
        fonts = []

        # ลอง font ยอดนิยมในเกม
        font_names = [
            "arial.ttf", "arialbd.ttf", "calibri.ttf", "calibrib.ttf",
            "segoeui.ttf", "segoeuib.ttf", "consola.ttf", "consolab.ttf",
            "tahoma.ttf", "tahomabd.ttf", "verdana.ttf", "verdanab.ttf",
            "impact.ttf", "trebuc.ttf", "trebucbd.ttf",
            "comic.ttf", "comicbd.ttf", "cour.ttf", "courbd.ttf",
        ]

        # Windows font paths
        win_paths = [
            os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
        ]

        for fname in font_names:
            for fdir in win_paths:
                fpath = os.path.join(fdir, fname)
                if os.path.exists(fpath):
                    fonts.append(fpath)
                    break

        # Fallback: PIL default font
        if not fonts:
            fonts.append(None)

        return fonts[:6]  # max 6 fonts

    def _render_char(self, char, font_path, size):
        """render 1 ตัวอักษรเป็น grayscale image"""
        try:
            pad = 4
            img = Image.new("L", (size + pad*2, size + pad*2), 0)
            draw = ImageDraw.Draw(img)

            if font_path:
                font = ImageFont.truetype(font_path, size)
            else:
                font = ImageFont.load_default()

            # วาดตัวอักษรขาวบนดำ
            bbox = draw.textbbox((0, 0), char, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            x = (img.width - tw) // 2 - bbox[0]
            y = (img.height - th) // 2 - bbox[1]
            draw.text((x, y), char, fill=255, font=font)

            # Crop ให้พอดีตัวอักษร
            arr = np.array(img)
            coords = np.argwhere(arr > 50)
            if len(coords) < 10:
                return None
            y0, x0 = coords.min(axis=0)
            y1, x1 = coords.max(axis=0)
            cropped = arr[max(0,y0-2):y1+3, max(0,x0-2):x1+3]

            if cropped.shape[0] < 5 or cropped.shape[1] < 5:
                return None

            return cropped

        except:
            return None

    def match_slot(self, gray_slot, threshold=0.55):
        """match 1 ช่องกับทุก template, return (char, confidence)"""
        # Preprocess slot
        slot = self._preprocess(gray_slot)
        if slot is None:
            return None, 0

        best_char, best_val = None, 0
        sh, sw = slot.shape[:2]

        for char, tmpls in self.templates.items():
            for tmpl in tmpls:
                th, tw = tmpl.shape[:2]
                # Skip ถ้าขนาดต่างกันมาก
                if tw > sw or th > sh:
                    continue
                if tw < sw * 0.2 or th < sh * 0.2:
                    continue

                # Resize template ให้พอดี slot
                scale_w = sw / tw
                scale_h = sh / th
                scale = min(scale_w, scale_h) * 0.85
                ntw, nth = int(tw * scale), int(th * scale)
                if ntw < 3 or nth < 3 or ntw > sw or nth > sh:
                    continue

                resized = cv2.resize(tmpl, (ntw, nth))
                res = cv2.matchTemplate(slot, resized, cv2.TM_CCOEFF_NORMED)
                _, mx, _, _ = cv2.minMaxLoc(res)

                if mx > threshold and mx > best_val:
                    best_char, best_val = char, mx
                    if mx > 0.85:
                        return best_char, best_val

        return best_char, best_val

    def _preprocess(self, gray):
        """ทำให้ตัวอักษรเป็น ขาวบนดำ ชัดที่สุด"""
        if gray is None or gray.size == 0:
            return None

        # ขยาย
        h, w = gray.shape[:2]
        if w < 20 or h < 20:
            scale = max(40/w, 40/h, 1)
            gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

        # ลองทั้ง 2 แบบ: ปกติ + invert เลือกแบบที่ตัวอักษรเป็นสีขาว
        _, thresh1 = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        thresh2 = cv2.bitwise_not(thresh1)

        # เลือกแบบที่มี content น้อยกว่า (ตัวอักษรควรเป็นส่วนน้อย)
        w1 = np.count_nonzero(thresh1)
        w2 = np.count_nonzero(thresh2)

        return thresh1 if w1 < w2 else thresh2

    def match_lane(self, gray_full, num_keys, threshold=0.55):
        """match ทั้งแถว return list of (char, confidence)"""
        h, w = gray_full.shape[:2]
        slot_w = w // num_keys
        results = []

        for i in range(num_keys):
            x1 = i * slot_w
            x2 = min(x1 + slot_w, w)
            pad = slot_w // 8
            rx1 = max(0, x1 - pad)
            rx2 = min(w, x2 + pad)
            roi = gray_full[:, rx1:rx2]

            if roi.shape[1] < 5:
                results.append((None, 0))
                continue

            char, conf = self.match_slot(roi, threshold)
            results.append((char, conf))

        return results


# ═══════════════════════════════════════════════════
# Lane Detector
# ═══════════════════════════════════════════════════
class LaneDetect:
    def __init__(self):
        self.last_hash = None

    def is_new(self, results):
        h = "".join(k or "?" for k, _ in results)
        if h != self.last_hash and "?" not in h:
            return True
        return False

    def record(self, results):
        self.last_hash = "".join(k or "?" for k, _ in results)

    def reset(self):
        self.last_hash = None


# ═══════════════════════════════════════════════════
# Theme
# ═══════════════════════════════════════════════════
C = {
    "bg":"#0a0e17", "card":"#111a2b", "border":"#1e3050",
    "blue":"#00d4ff", "green":"#22ff88", "red":"#ff4466",
    "orange":"#ff9f43", "purple":"#a855f7", "cyan":"#00e5ff",
    "white":"#eaf0f6", "dim":"#5a6a7e", "dim2":"#3d4f63", "input":"#0c1220",
}


# ═══════════════════════════════════════════════════
# Region Picker (screenshot + drag)
# ═══════════════════════════════════════════════════
class RegionPicker(tk.Toplevel):
    def __init__(self, master, callback):
        super().__init__(master)
        self.callback = callback
        self.sx = self.sy = 0
        self.rect = None

        # Screenshot ก่อนเปิด overlay
        try:
            self.screenshot = ImageGrab.grab()
            self.photo = ImageTk.PhotoImage(self.screenshot)
        except:
            self.screenshot = None
            self.photo = None

        self.overrideredirect(True)
        self.attributes("-topmost", True)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0")

        self.c = tk.Canvas(self, highlightthickness=0, cursor="cross")
        self.c.pack(fill=tk.BOTH, expand=True)

        if self.photo:
            self.c.create_image(0, 0, anchor="nw", image=self.photo)
        self.c.create_rectangle(0, 0, sw, sh, fill="black", stipple="gray25")

        # Guide
        self.c.create_rectangle(sw//2-200, 16, sw//2+200, 72, fill="#000", outline="#00ffaa")
        self.c.create_text(sw//2, 32, text="ลากเลือกบริเวณที่ตัวอักษรขึ้น",
            fill="#00ffaa", font=("Segoe UI", 13, "bold"))
        self.c.create_text(sw//2, 56, text="ลากให้ครอบ ทั้งแถว ตัวอักษร  |  ESC = ยกเลิก",
            fill="#ccc", font=("Segoe UI", 10))

        self.pos_txt = self.c.create_text(sw//2, 88, text="", fill="#00ffaa", font=("Consolas", 10))

        self.c.bind("<Motion>", lambda e: self.c.itemconfig(self.pos_txt, text=f"X:{e.x}  Y:{e.y}"))
        self.c.bind("<ButtonPress-1>", self._press)
        self.c.bind("<B1-Motion>", self._drag)
        self.c.bind("<ButtonRelease-1>", self._release)
        self.bind("<Escape>", lambda e: self._cancel())
        self.after(50, self.focus_force)

    def _press(self, e):
        self.sx, self.sy = e.x, e.y
        if self.rect: self.c.delete(self.rect)
        self.c.delete("hint"); self.c.delete("sz")
        self.rect = self.c.create_rectangle(e.x, e.y, e.x, e.y, outline="#00ffaa", width=2)

    def _drag(self, e):
        self.c.coords(self.rect, self.sx, self.sy, e.x, e.y)
        self.c.delete("sz")
        w, h = abs(e.x-self.sx), abs(e.y-self.sy)
        self.c.create_text((self.sx+e.x)//2, min(self.sy, e.y)-16,
            text=f"{w} x {h}", fill="#00ffaa", font=("Consolas", 12, "bold"), tags="sz")

    def _release(self, e):
        x1, y1 = min(self.sx, e.x), min(self.sy, e.y)
        x2, y2 = max(self.sx, e.x), max(self.sy, e.y)
        if (x2-x1) > 10 and (y2-y1) > 10:
            self.callback({"left": x1, "top": y1, "width": x2-x1, "height": y2-y1})
            self.destroy()

    def _cancel(self):
        self.master.deiconify()
        self.destroy()


# ═══════════════════════════════════════════════════
# MAIN APP
# ═══════════════════════════════════════════════════
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QTE Macro v5")
        self.root.geometry("480x600")
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)

        self.cfg = self._load_cfg()
        self.running = False
        self.region = self.cfg.get("region")
        self.session_keys = 0
        self.session_start = None
        self.key_history = []

        self.cap = Cap()
        self.engine = AutoTemplates()
        self.lane = LaneDetect()

        self.v_num = tk.IntVar(value=self.cfg.get("num_keys", 6))
        self.v_scan = tk.IntVar(value=self.cfg.get("scan_ms", 40))
        self.v_keydelay = tk.IntVar(value=self.cfg.get("key_delay", 50))
        self.v_lanedelay = tk.IntVar(value=self.cfg.get("lane_delay", 300))
        self.v_ontop = tk.BooleanVar(value=self.cfg.get("ontop", True))

        self._build()
        self.root.attributes("-topmost", self.v_ontop.get())
        self._update_region()
        self._preview_loop()
        self._timer_loop()

        self.root.bind_all("<F5>", lambda e: self.toggle())
        self.root.bind_all("<F6>", lambda e: self.pick_region())
        self.root.bind_all("<Escape>", lambda e: self._stop())
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

        n = sum(len(v) for v in self.engine.templates.values())
        self.log(f"Auto-generated {n} character templates")

    # ── Config ──
    def _cfg_path(self): return os.path.join(get_base_path(), "macro_config.json")
    def _load_cfg(self):
        try:
            with open(self._cfg_path(), "r") as f: return json.load(f)
        except: return {}
    def _save_cfg(self):
        self.cfg.update({"num_keys":self.v_num.get(), "scan_ms":self.v_scan.get(),
            "key_delay":self.v_keydelay.get(), "lane_delay":self.v_lanedelay.get(),
            "region":self.region, "ontop":self.v_ontop.get()})
        try:
            with open(self._cfg_path(), "w") as f: json.dump(self.cfg, f, indent=2)
        except: pass

    # ── UI ──
    def _build(self):
        # Header
        hdr = tk.Canvas(self.root, bg=C["bg"], height=56, highlightthickness=0)
        hdr.pack(fill=tk.X)
        for i in range(480):
            t = i/480
            hdr.create_line(i, 0, i, 3, fill=f"#{int(168*t):02x}{int(212-127*t):02x}{int(255-8*t):02x}")
        hdr.create_text(20, 22, text="QTE MACRO", anchor="w", fill=C["blue"], font=("Segoe UI", 20, "bold"))
        hdr.create_text(20, 44, text="Zero Setup - เลือก region แล้วกด START เลย", anchor="w",
                        fill=C["dim"], font=("Segoe UI", 9))
        self.lbl_st = hdr.create_text(440, 22, text="OFF", anchor="w", fill=C["dim"], font=("Segoe UI", 10, "bold"))
        self.lbl_tm = hdr.create_text(460, 44, text="00:00", anchor="e", fill=C["dim2"], font=("Consolas", 9))
        self.hdr = hdr

        main = tk.Frame(self.root, bg=C["bg"])
        main.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 8))

        # START button
        self.btn_start = tk.Button(main, text="  START (F5)  ", bg=C["green"], fg="#000",
            font=("Segoe UI", 16, "bold"), relief="flat", cursor="hand2", pady=8,
            activebackground="#1ecc6a", command=self.toggle)
        self.btn_start.pack(fill=tk.X, pady=(4, 8))

        # Stats
        sf = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        sf.pack(fill=tk.X, pady=(0, 8))
        si = tk.Frame(sf, bg=C["card"]); si.pack(fill=tk.X, padx=12, pady=8)

        self.lbl_count = tk.Label(si, text="0", bg=C["card"], fg=C["green"], font=("Consolas", 30, "bold"))
        self.lbl_count.pack(side=tk.LEFT)

        mid = tk.Frame(si, bg=C["card"]); mid.pack(side=tk.LEFT, padx=14)
        self.lbl_lane = tk.Label(mid, text="", bg=C["card"], fg=C["cyan"], font=("Consolas", 14, "bold"))
        self.lbl_lane.pack(anchor="w")
        self.lbl_speed = tk.Label(mid, text="", bg=C["card"], fg=C["dim"], font=("Consolas", 9))
        self.lbl_speed.pack(anchor="w")

        # Preview
        self.preview = tk.Label(sf, bg="#0c1220", text="Select region first (F6)", fg=C["dim2"],
                                 font=("Segoe UI", 8), height=3)
        self.preview.pack(fill=tk.X, padx=12, pady=(0, 8))

        # History
        self.lbl_hist = tk.Label(sf, text="", bg=C["card"], fg=C["dim"], font=("Consolas", 9), anchor="w")
        self.lbl_hist.pack(fill=tk.X, padx=12, pady=(0, 6))

        # Region
        self._sec(main, "STEP 1 — SELECT REGION")
        rf = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        rf.pack(fill=tk.X, pady=(0, 6))
        ri = tk.Frame(rf, bg=C["card"]); ri.pack(fill=tk.X, padx=12, pady=8)

        tk.Button(ri, text="  Select Region (F6)  ", bg=C["blue"], fg="#000",
            font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
            activebackground="#0099cc", command=self.pick_region).pack(side=tk.LEFT)

        tk.Button(ri, text=" Test ", bg=C["orange"], fg="#000",
            font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
            command=self._test).pack(side=tk.LEFT, padx=6)

        self.lbl_reg = tk.Label(ri, text="Not set", bg=C["card"], fg=C["orange"], font=("Consolas", 9))
        self.lbl_reg.pack(side=tk.RIGHT)

        # Settings
        self._sec(main, "STEP 2 — SETTINGS")
        stf = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        stf.pack(fill=tk.X, pady=(0, 6))
        sti = tk.Frame(stf, bg=C["card"]); sti.pack(fill=tk.X, padx=12, pady=8)

        self._slider(sti, "Keys per Lane", "", self.v_num, 1, 12, C["blue"])
        self._slider(sti, "Scan Speed", "ms", self.v_scan, 20, 150, C["cyan"])
        self._slider(sti, "Key Delay", "ms", self.v_keydelay, 20, 200, C["purple"])
        self._slider(sti, "Lane Delay", "ms", self.v_lanedelay, 100, 1000, C["green"])

        chk = tk.Frame(sti, bg=C["card"]); chk.pack(fill=tk.X, pady=(4, 0))
        tk.Checkbutton(chk, text="Always on Top", variable=self.v_ontop, bg=C["card"],
            fg=C["white"], selectcolor=C["bg"], activebackground=C["card"], font=("Segoe UI", 9),
            command=lambda: self.root.attributes("-topmost", self.v_ontop.get())).pack(side=tk.LEFT)

        # Log
        self._sec(main, "LOG")
        lf = tk.Frame(main, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_box = tk.Text(lf, bg=C["bg"], fg=C["green"], font=("Consolas", 9),
            height=4, relief="flat", wrap=tk.WORD, state=tk.DISABLED, padx=8, pady=4)
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self.log_box.tag_configure("key", foreground=C["cyan"], font=("Consolas", 9, "bold"))
        self.log_box.tag_configure("err", foreground=C["red"])
        self.log_box.tag_configure("ok", foreground=C["green"])
        self.log_box.tag_configure("warn", foreground=C["orange"])
        self.log_box.tag_configure("dim", foreground=C["dim"])
        self.log_box.tag_configure("lane", foreground=C["purple"], font=("Consolas", 10, "bold"))

        # Footer
        ft = tk.Label(self.root, text="F5 Start/Stop   F6 Region   ESC Stop",
            bg="#0f1923", fg=C["dim2"], font=("Segoe UI", 8))
        ft.pack(fill=tk.X, pady=4)

    def _sec(self, p, t):
        f = tk.Frame(p, bg=C["bg"]); f.pack(fill=tk.X, pady=(6, 2))
        tk.Label(f, text=t, bg=C["bg"], fg=C["dim2"], font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT)

    def _slider(self, p, label, unit, var, lo, hi, color):
        row = tk.Frame(p, bg=C["card"]); row.pack(fill=tk.X, pady=2)
        tk.Label(row, text=f"{label}:", bg=C["card"], fg=C["white"], font=("Segoe UI", 9)).pack(side=tk.LEFT)
        vl = tk.Label(row, text=f"{var.get()}{unit}", bg=C["card"], fg=color,
            font=("Consolas", 10, "bold"), width=5, anchor="e")
        vl.pack(side=tk.RIGHT)
        tk.Scale(row, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var,
            bg=C["card"], fg=C["card"], troughcolor=C["bg"], highlightthickness=0,
            showvalue=False, length=150, sliderlength=14, activebackground=color,
            command=lambda v, l=vl, u=unit: l.config(text=f"{int(float(v))}{u}")
        ).pack(side=tk.RIGHT, padx=4)

    # ── Region ──
    def pick_region(self):
        if self.running: self.log("Stop first!", "err"); return
        self.root.iconify()
        time.sleep(0.3)
        def done(r):
            self.region = r; self._save_cfg(); self.root.deiconify(); self._update_region()
            self.log(f"Region: {r['width']}x{r['height']} @ ({r['left']},{r['top']})", "ok")
        p = RegionPicker(self.root, done)
        p.protocol("WM_DELETE_WINDOW", lambda: (self.root.deiconify(), p.destroy()))

    def _update_region(self):
        if self.region:
            r = self.region
            self.lbl_reg.config(text=f"{r['width']}x{r['height']}  ({r['left']},{r['top']})", fg=C["green"])
        else:
            self.lbl_reg.config(text="Not set", fg=C["orange"])

    def _test(self):
        if not self.region: messagebox.showwarning("Test", "Select region first!"); return
        frame = self.cap.grab(self.region)
        if frame is None: self.log("Capture FAILED!", "err"); return
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        results = self.engine.match_lane(gray, self.v_num.get())
        seq = " ".join(f"{k.upper()}({c:.0%})" if k else "??" for k, c in results)
        self.log(f"Test: {seq}", "key")

    # ── Macro ──
    def toggle(self):
        if self.running: self._stop()
        else: self._start()

    def _start(self):
        if not self.region: messagebox.showwarning("QTE", "Select region first! (F6)"); return
        self.lane.reset()
        self.running = True; self.session_keys = 0; self.session_start = time.time()
        self.key_history = []
        self.btn_start.config(text="  STOP (F5)  ", bg=C["red"])
        self.hdr.itemconfig(self.lbl_st, text="ON", fill=C["green"])
        self.log("Macro STARTED", "ok")
        self._save_cfg()
        threading.Thread(target=self._loop, daemon=True).start()

    def _stop(self):
        if not self.running: return
        self.running = False
        self.btn_start.config(text="  START (F5)  ", bg=C["green"])
        self.hdr.itemconfig(self.lbl_st, text="OFF", fill=C["dim"])
        e = time.time() - (self.session_start or time.time())
        self.log(f"Stopped ({self.session_keys} keys / {e:.1f}s)")

    def _loop(self):
        num = self.v_num.get()
        kd = self.v_keydelay.get() / 1000.0
        ld = self.v_lanedelay.get() / 1000.0

        while self.running:
            try:
                frame = self.cap.grab(self.region)
                if frame is None: time.sleep(0.05); continue

                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                results = self.engine.match_lane(gray, num)

                # ทุกช่องต้อง match ได้
                valid = all(k is not None for k, _ in results)
                if not valid:
                    time.sleep(self.v_scan.get() / 1000.0)
                    continue

                # เป็น lane ใหม่?
                if not self.lane.is_new(results):
                    time.sleep(self.v_scan.get() / 1000.0)
                    continue

                # Lane ใหม่! กดทีละตัว
                self.lane.record(results)
                seq = " ".join(k.upper() for k, _ in results)
                self.root.after(0, self.lbl_lane.config, {"text": seq})
                self.root.after(0, self.log, f"  Lane: {seq}", "lane")

                for key, conf in results:
                    if not self.running: break
                    if key:
                        kb.press_and_release(key)
                        self.session_keys += 1
                        self.root.after(0, self._on_key, key)
                        time.sleep(kd)

                time.sleep(ld)

            except Exception as e:
                self.root.after(0, self.log, str(e), "err")
                time.sleep(0.1)

    def _on_key(self, key):
        self.lbl_count.config(text=str(self.session_keys))
        e = time.time() - (self.session_start or time.time())
        if e > 0: self.lbl_speed.config(text=f"{self.session_keys/e:.1f} keys/sec")
        self.key_history.append(key.upper())
        if len(self.key_history) > 30: self.key_history.pop(0)
        self.lbl_hist.config(text=" ".join(self.key_history[-20:]))

    # ── Loops ──
    def _preview_loop(self):
        if self.region and self.cap.ok:
            try:
                img = self.cap.grab_pil(self.region)
                if img:
                    w, h = img.size
                    nw = min(450, w); nh = int(h * (nw / w))
                    img = img.resize((nw, max(nh, 20)), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    self.preview.config(image=photo, text=""); self.preview._photo = photo
            except: pass
        self.root.after(500, self._preview_loop)

    def _timer_loop(self):
        if self.running and self.session_start:
            e = int(time.time() - self.session_start); m, s = divmod(e, 60)
            self.hdr.itemconfig(self.lbl_tm, text=f"{m:02d}:{s:02d}", fill=C["green"])
        else:
            self.hdr.itemconfig(self.lbl_tm, text="00:00", fill=C["dim2"])
        self.root.after(500, self._timer_loop)

    def log(self, msg, tag=None):
        def _do():
            self.log_box.config(state=tk.NORMAL)
            ts = datetime.now().strftime("%H:%M:%S")
            if tag:
                self.log_box.insert(tk.END, f"[{ts}] ", "dim")
                self.log_box.insert(tk.END, msg + "\n", tag)
            else:
                self.log_box.insert(tk.END, f"[{ts}] {msg}\n")
            self.log_box.see(tk.END)
            lines = int(self.log_box.index("end-1c").split(".")[0])
            if lines > 150: self.log_box.delete("1.0", f"{lines-150}.0")
            self.log_box.config(state=tk.DISABLED)
        self.root.after(0, _do)

    def _quit(self):
        self.running = False; self._save_cfg()
        time.sleep(0.1); self.root.destroy()

    def run(self): self.root.mainloop()


if __name__ == "__main__":
    App().run()
