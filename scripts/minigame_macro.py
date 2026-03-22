"""
QTE Macro - Final Working Version
tkinter UI + Windows SendInput API + Auto Learn
"""

import time, sys, os, json, threading, ctypes, random
import tkinter as tk
from tkinter import messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ═══ Libs ═══
try:
    import keyboard as kb
    import mss
    import cv2
    import numpy as np
    from PIL import Image, ImageTk, ImageGrab
except ImportError as e:
    print(f"Missing: {e}\npip install keyboard mss opencv-python numpy Pillow")
    sys.exit(1)

# ═══════════════════════════════════
# Windows SendInput - กดปุ่มเข้าเกมได้ 100%
# ═══════════════════════════════════
SCAN = {'q':0x10,'w':0x11,'e':0x12,'r':0x13,'a':0x1E,'s':0x1F,'d':0x20,'f':0x21,
        'z':0x2C,'x':0x2D,'c':0x2E,'v':0x2F,'1':0x02,'2':0x03,'3':0x04,'4':0x05}

def press_key(key):
    """กดปุ่มด้วย Windows SendInput API (เข้าเกมได้ทุกเกม)"""
    key = key.lower()
    sc = SCAN.get(key)
    if sc is None:
        return False

    if sys.platform == 'win32':
        try:
            PUL = ctypes.POINTER(ctypes.c_ulong)
            class KEYBDINPUT(ctypes.Structure):
                _fields_ = [("wVk",ctypes.c_ushort),("wScan",ctypes.c_ushort),
                            ("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),
                            ("dwExtraInfo",PUL)]
            class HARDWAREINPUT(ctypes.Structure):
                _fields_ = [("uMsg",ctypes.c_ulong),("wParamL",ctypes.c_short),("wParamH",ctypes.c_ushort)]
            class MOUSEINPUT(ctypes.Structure):
                _fields_ = [("dx",ctypes.c_long),("dy",ctypes.c_long),("mouseData",ctypes.c_ulong),
                            ("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
            class INPUT_UNION(ctypes.Union):
                _fields_ = [("ki",KEYBDINPUT),("mi",MOUSEINPUT),("hi",HARDWAREINPUT)]
            class INPUT(ctypes.Structure):
                _fields_ = [("type",ctypes.c_ulong),("ii",INPUT_UNION)]

            extra = ctypes.c_ulong(0)
            # Key down (scancode)
            ii = INPUT_UNION(); ii.ki = KEYBDINPUT(0, sc, 0x0008, 0, ctypes.pointer(extra))
            inp = INPUT(ctypes.c_ulong(1), ii)
            ctypes.windll.user32.SendInput(1, ctypes.pointer(inp), ctypes.sizeof(inp))
            time.sleep(0.03)
            # Key up
            ii2 = INPUT_UNION(); ii2.ki = KEYBDINPUT(0, sc, 0x0008|0x0002, 0, ctypes.pointer(extra))
            inp2 = INPUT(ctypes.c_ulong(1), ii2)
            ctypes.windll.user32.SendInput(1, ctypes.pointer(inp2), ctypes.sizeof(inp2))
            return True
        except:
            pass

    # Fallback: pyautogui / keyboard
    try:
        import pyautogui; pyautogui.PAUSE = 0; pyautogui.press(key); return True
    except: pass
    try:
        kb.press_and_release(key); return True
    except: pass
    return False


# ═══════════════════════════════════
# Screen Capture - ลองทุกวิธี + log error
# ═══════════════════════════════════
grab_error = ""

def grab(r):
    """จับหน้าจอ ลอง 3 วิธี"""
    global grab_error
    left, top = int(r["left"]), int(r["top"])
    w, h = int(r["width"]), int(r["height"])

    # 1. mss (สร้างใหม่ทุกครั้ง ป้องกัน stale handle)
    try:
        with mss.mss() as s:
            shot = s.grab({"left":left, "top":top, "width":w, "height":h})
            arr = np.array(shot)
            if arr.size > 0 and arr.shape[0] > 0 and arr.shape[1] > 0:
                return cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)
    except Exception as e:
        grab_error = f"mss:{e}"

    # 2. PIL ImageGrab
    try:
        img = ImageGrab.grab(bbox=(left, top, left+w, top+h))
        if img:
            arr = np.array(img)
            if arr.size > 0:
                return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
    except Exception as e:
        grab_error = f"PIL:{e}"

    # 3. pyautogui
    try:
        import pyautogui
        img = pyautogui.screenshot(region=(left, top, w, h))
        if img:
            return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    except Exception as e:
        grab_error = f"pyautogui:{e}"

    return None


# ═══════════════════════════════════
# Templates
# ═══════════════════════════════════
TMPL_DIR = os.path.join(SCRIPT_DIR, "game_templates")
os.makedirs(TMPL_DIR, exist_ok=True)
templates = {}
TH = 48

def load_templates():
    global templates; templates.clear()
    for f in os.listdir(TMPL_DIR):
        if not f.endswith(".png"): continue
        key = f.split("_")[0].split(".")[0].lower()
        if len(key) != 1: continue
        img = cv2.imread(os.path.join(TMPL_DIR, f), cv2.IMREAD_GRAYSCALE)
        if img is not None:
            h,w = img.shape[:2]
            templates.setdefault(key, []).append(cv2.resize(img, (max(int(w*TH/h),3), TH)))

def save_template(key, gray_img):
    key = key.lower()
    n = len([f for f in os.listdir(TMPL_DIR) if f.startswith(key)])
    fname = f"{key}.png" if n==0 else f"{key}_{n+1}.png"
    cv2.imwrite(os.path.join(TMPL_DIR, fname), gray_img)
    h,w = gray_img.shape[:2]
    templates.setdefault(key, []).append(cv2.resize(gray_img, (max(int(w*TH/h),3), TH)))

def match_tmpl(gray_roi, threshold=0.55):
    if not templates or gray_roi is None or gray_roi.size < 20: return None, 0
    h,w = gray_roi.shape[:2]
    if h<3 or w<3: return None, 0
    roi = cv2.resize(gray_roi, (max(int(w*TH/h),3), TH))
    best_c, best_v = None, 0
    for ch, tmpls in templates.items():
        for t in tmpls:
            tw = t.shape[1]
            if tw > roi.shape[1]*1.5 or tw < roi.shape[1]*0.3: continue
            target = roi
            if tw > roi.shape[1] or t.shape[0] > roi.shape[0]:
                px = max(0, (tw - roi.shape[1])//2 + 3)
                py = max(0, (t.shape[0] - roi.shape[0])//2 + 3)
                target = cv2.copyMakeBorder(roi, py, py, px, px, cv2.BORDER_CONSTANT, value=0)
            try:
                res = cv2.matchTemplate(target, t, cv2.TM_CCOEFF_NORMED)
                _, mx, _, _ = cv2.minMaxLoc(res)
                if mx > threshold and mx > best_v:
                    best_c, best_v = ch, mx
                    if mx > 0.85: return best_c, best_v
            except: continue
    return best_c, best_v

load_templates()


# ═══════════════════════════════════
# Image Utils
# ═══════════════════════════════════
def get_slot(gray, i, n):
    h,w = gray.shape[:2]; sw = w//n
    pad = max(sw//10, 2)
    return gray[:, max(0,i*sw-pad):min(w,(i+1)*sw+pad)]

def prep(gray):
    if gray is None or gray.size < 20: return None
    h,w = gray.shape[:2]
    if h < 25 or w < 15:
        s = max(30/max(h,1), 30/max(w,1), 1)
        gray = cv2.resize(gray, None, fx=s, fy=s, interpolation=cv2.INTER_CUBIC)
    best, best_s = None, 0
    for tval in [0, 127, 180]:
        try:
            if tval == 0:
                _, t = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            else:
                _, t = cv2.threshold(gray, tval, 255, cv2.THRESH_BINARY)
            ti = cv2.bitwise_not(t)
            r = t if np.count_nonzero(t) < np.count_nonzero(ti) else ti
            cs = np.argwhere(r > 127)
            if len(cs) < 5: continue
            y0,x0 = cs.min(0); y1,x1 = cs.max(0)
            c = r[max(0,y0-1):y1+2, max(0,x0-1):x1+2]
            if c.shape[0]<4 or c.shape[1]<4: continue
            ratio = np.count_nonzero(c>127) / max(c.size,1)
            if 0.05 < ratio < 0.65 and c.size > best_s:
                best = c; best_s = c.size
        except: continue
    return best

def slot_changed(orig, curr, thresh=20):
    if orig is None: return False
    diff = cv2.absdiff(orig, curr)
    if len(diff.shape)==3: diff = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
    return (np.count_nonzero(diff > 30) / max(diff.size,1)) * 100 > thresh


# ═══════════════════════════════════
# Config
# ═══════════════════════════════════
CFG = os.path.join(SCRIPT_DIR, "config.json")
region = None

def load_cfg():
    global region
    try:
        with open(CFG,"r") as f: c = json.load(f); region = c.get("region"); return c
    except: return {}

def save_cfg(**kw):
    c = load_cfg(); c.update(kw); c["region"] = region
    try:
        with open(CFG,"w") as f: json.dump(c, f, indent=2)
    except: pass

load_cfg()


# ═══════════════════════════════════
# Main App (tkinter)
# ═══════════════════════════════════
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QTE Macro")
        self.root.geometry("420x580")
        self.root.configure(bg="#0a0e17")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        self.running = False
        self.session_keys = 0
        self.session_start = 0
        self.user_pressed = None
        self.num_keys = tk.IntVar(value=load_cfg().get("num_keys", 5))
        self.key_delay = tk.IntVar(value=50)
        self.lane_delay = tk.IntVar(value=500)

        # Keyboard listener
        kb.on_press(self._on_key)

        self._build()
        self._update_region_label()
        self._update_template_label()
        self._preview_loop()

        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _on_key(self, event):
        k = event.name.lower()
        if k in "qweasd" and self.running:
            self.user_pressed = k

    def _build(self):
        BG = "#0a0e17"
        CARD = "#111a2b"
        DIM = "#5a6a7e"

        # Header
        tk.Label(self.root, text="QTE Macro", bg=BG, fg="#22d3ee",
                 font=("Segoe UI", 16, "bold")).pack(pady=(10,0))
        tk.Label(self.root, text="เล่นเองรอบแรก → Auto ตลอด", bg=BG, fg=DIM,
                 font=("Segoe UI", 9)).pack()

        # Preview
        self.preview_label = tk.Label(self.root, bg="#0c1220", text="กด Select Region ก่อน",
                                       fg=DIM, font=("Segoe UI", 8), height=3)
        self.preview_label.pack(fill=tk.X, padx=14, pady=(8,4))

        # Status
        self.status_label = tk.Label(self.root, text="OFF", bg=BG, fg=DIM,
                                      font=("Consolas", 10, "bold"))
        self.status_label.pack()

        # START
        self.btn_start = tk.Button(self.root, text="  START  ", bg="#10b981", fg="white",
            font=("Segoe UI", 14, "bold"), relief="flat", cursor="hand2", pady=6,
            activebackground="#059669", command=self.toggle)
        self.btn_start.pack(fill=tk.X, padx=14, pady=6)

        # Count + Speed
        f1 = tk.Frame(self.root, bg=CARD); f1.pack(fill=tk.X, padx=14, pady=2)
        self.lbl_count = tk.Label(f1, text="0", bg=CARD, fg="#10b981",
                                   font=("Consolas", 24, "bold"))
        self.lbl_count.pack(side=tk.LEFT, padx=10, pady=6)
        self.lbl_speed = tk.Label(f1, text="", bg=CARD, fg=DIM, font=("Consolas", 10))
        self.lbl_speed.pack(side=tk.LEFT)
        self.lbl_learned = tk.Label(f1, text="", bg=CARD, fg="#a78bfa", font=("Consolas", 10, "bold"))
        self.lbl_learned.pack(side=tk.RIGHT, padx=10)

        # Region
        tk.Label(self.root, text="REGION", bg=BG, fg=DIM, font=("Segoe UI", 8, "bold"),
                 anchor="w").pack(fill=tk.X, padx=16, pady=(8,2))
        f2 = tk.Frame(self.root, bg=CARD); f2.pack(fill=tk.X, padx=14, pady=2)
        tk.Button(f2, text="Select Region (F6)", bg="#1e3a5f", fg="#22d3ee",
            font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4,
            command=self.pick_region).pack(side=tk.LEFT, padx=6, pady=6)
        tk.Button(f2, text="Test Press", bg="#3b1764", fg="#a78bfa",
            font=("Segoe UI", 9, "bold"), relief="flat", padx=8, pady=4,
            command=self.test_press).pack(side=tk.LEFT, padx=2, pady=6)
        self.lbl_region = tk.Label(f2, text="not set", bg=CARD, fg="#f97316",
                                    font=("Consolas", 9))
        self.lbl_region.pack(side=tk.RIGHT, padx=8)

        # Settings
        tk.Label(self.root, text="SETTINGS", bg=BG, fg=DIM, font=("Segoe UI", 8, "bold"),
                 anchor="w").pack(fill=tk.X, padx=16, pady=(8,2))
        sf = tk.Frame(self.root, bg=CARD); sf.pack(fill=tk.X, padx=14, pady=2)
        self._slider(sf, "Keys/Lane", self.num_keys, 2, 10)
        self._slider(sf, "Key Delay ms", self.key_delay, 20, 200)
        self._slider(sf, "Round Delay ms", self.lane_delay, 200, 1500)

        # Clear
        tk.Button(sf, text="Clear Templates", bg="#3b1520", fg="#ef4444",
            font=("Segoe UI", 8), relief="flat", padx=6, pady=2,
            command=self.clear_templates).pack(anchor="w", padx=8, pady=(2,6))

        # Log
        tk.Label(self.root, text="LOG", bg=BG, fg=DIM, font=("Segoe UI", 8, "bold"),
                 anchor="w").pack(fill=tk.X, padx=16, pady=(8,2))
        self.log_text = tk.Text(self.root, bg="#070c18", fg="#22ff88", font=("Consolas", 9),
            height=8, relief="flat", wrap=tk.WORD, state=tk.DISABLED, padx=8, pady=4)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=14, pady=(2,10))

        # Hotkeys
        self.root.bind("<F5>", lambda e: self.toggle())
        self.root.bind("<F6>", lambda e: self.pick_region())

        self.log("Ready")
        self.log(f"Key method: {'SendInput' if sys.platform=='win32' else 'fallback'}")
        lk = sorted(templates.keys())
        if lk:
            self.log(f"Loaded templates: {' '.join(k.upper() for k in lk)}")

    def _slider(self, parent, label, var, lo, hi):
        f = tk.Frame(parent, bg="#111a2b"); f.pack(fill=tk.X, padx=8, pady=2)
        tk.Label(f, text=label+":", bg="#111a2b", fg="#eee", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        vl = tk.Label(f, text=str(var.get()), bg="#111a2b", fg="#22d3ee",
                       font=("Consolas", 9, "bold"), width=5)
        vl.pack(side=tk.RIGHT)
        tk.Scale(f, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var,
                 bg="#111a2b", fg="#111a2b", troughcolor="#0a0e17", highlightthickness=0,
                 showvalue=False, length=140, sliderlength=14,
                 command=lambda v, l=vl: l.config(text=str(int(float(v))))).pack(side=tk.RIGHT, padx=4)

    # ═══ Region ═══
    def pick_region(self):
        global region
        if self.running: self.log("Stop first!"); return
        self.root.iconify()
        time.sleep(0.3)

        sel = tk.Toplevel(self.root)
        sel.overrideredirect(True)
        sel.attributes("-topmost", True)
        sel.attributes("-alpha", 0.25)  # โปร่งใส 75% → เห็นเกมทะลุ
        sw, sh = sel.winfo_screenwidth(), sel.winfo_screenheight()
        sel.geometry(f"{sw}x{sh}+0+0")
        sel.configure(bg="black")

        c = tk.Canvas(sel, bg="black", highlightthickness=0, cursor="cross")
        c.pack(fill=tk.BOTH, expand=True)

        # กรอบคำแนะนำ
        c.create_rectangle(sw//2-200, 15, sw//2+200, 90, fill="black", outline="#00ff88", width=2)
        c.create_text(sw//2, 35, text="ลากครอบแถวตัวอักษร",
            fill="#00ff88", font=("Segoe UI", 16, "bold"))
        c.create_text(sw//2, 58, text="ไม่รวม counter ด้านขวา",
            fill="#00ff88", font=("Segoe UI", 11))
        c.create_text(sw//2, 78, text="ESC = ยกเลิก",
            fill="#aaaaaa", font=("Segoe UI", 10))

        # พิกัดเมาส์
        pos_txt = c.create_text(sw//2, 108, text="", fill="#00ff88", font=("Consolas", 12, "bold"))

        st = {"sx":0, "sy":0, "r":None}

        def _motion(e):
            c.itemconfig(pos_txt, text=f"X: {e.x}   Y: {e.y}")

        def _p(e):
            st["sx"],st["sy"] = e.x, e.y
            if st["r"]: c.delete(st["r"])
            st["r"] = c.create_rectangle(e.x, e.y, e.x, e.y,
                outline="#00ff88", width=3)

        def _d(e):
            c.coords(st["r"], st["sx"], st["sy"], e.x, e.y)
            c.delete("sz")
            w, h = abs(e.x-st["sx"]), abs(e.y-st["sy"])
            c.create_text((st["sx"]+e.x)//2, min(st["sy"],e.y)-16,
                text=f"{w} x {h}", fill="#00ff88", font=("Consolas", 14, "bold"), tags="sz")

        def _r(e):
            global region
            x1,y1 = min(st["sx"],e.x), min(st["sy"],e.y)
            x2,y2 = max(st["sx"],e.x), max(st["sy"],e.y)
            if (x2-x1)>10 and (y2-y1)>10:
                region = {"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}
                save_cfg(); self._update_region_label()
                self.log(f"Region: {x2-x1}x{y2-y1} @ ({x1},{y1})")
            sel.destroy(); self.root.deiconify()

        def _esc(e): sel.destroy(); self.root.deiconify()

        c.bind("<Motion>", _motion)
        c.bind("<ButtonPress-1>", _p)
        c.bind("<B1-Motion>", _d)
        c.bind("<ButtonRelease-1>", _r)
        sel.bind("<Escape>", _esc)
        sel.after(50, sel.focus_force)

    def _update_region_label(self):
        if region:
            r = region; self.lbl_region.config(text=f"{r['width']}x{r['height']}", fg="#10b981")
        else:
            self.lbl_region.config(text="not set", fg="#f97316")

    def _update_template_label(self):
        lk = sorted(templates.keys())
        self.lbl_learned.config(text=" ".join(k.upper() for k in lk) if lk else "")

    def test_press(self):
        """ทดสอบว่ากดปุ่มเข้าเกมได้ไหม"""
        self.log("Test: กดปุ่ม Q ใน 3 วินาที... (คลิกที่เกมก่อน!)")
        def _do():
            time.sleep(3)
            ok = press_key('q')
            self.root.after(0, self.log, f"Press Q: {'OK' if ok else 'FAILED'}")
        threading.Thread(target=_do, daemon=True).start()

    def clear_templates(self):
        import shutil
        if os.path.isdir(TMPL_DIR): shutil.rmtree(TMPL_DIR)
        os.makedirs(TMPL_DIR, exist_ok=True)
        templates.clear()
        self._update_template_label()
        self.log("Templates cleared!")

    # ═══ Toggle ═══
    def toggle(self):
        if self.running:
            self.running = False
        else:
            if not region: messagebox.showwarning("QTE","กด F6 เลือก Region ก่อน!"); return
            self.running = True
            save_cfg(num_keys=self.num_keys.get())
            self.btn_start.config(text="  STOP  ", bg="#ef4444")
            self.status_label.config(text="RUNNING", fg="#10b981")
            lk = sorted(templates.keys())
            if lk:
                self.log(f"AUTO mode! ({' '.join(k.upper() for k in lk)})")
                self.log("ถ้ากดผิด → กด Clear Templates แล้ว Start ใหม่")
            else:
                self.log("LEARN mode: คลิกที่เกมก่อน แล้วกดปุ่มตามเกมปกติ")
                self.log("macro จะจำตัวอักษรจากเกมเอง")
            threading.Thread(target=self._run, daemon=True).start()

    # ═══ Main Loop (AFK Fishing) ═══
    def _run(self):
        """
        AFK 100% Loop:
        1. รอมินิเกมขึ้น (สแกน 5 ช่องพร้อมกัน)
        2. เจอครบ → กดรวดเดียว 5 ตัว
        3. รอตัวละครดึงปลา (~5 วิ)
        4. กดโยนเบ็ดใหม่
        5. กลับข้อ 1
        """
        num = self.num_keys.get()
        kd = self.key_delay.get() / 1000
        self.session_keys = 0
        self.session_start = time.time()
        self.user_pressed = None

        # ทดสอบจับหน้าจอ
        test = grab(region)
        if test is None:
            self.root.after(0, self.log, f"จับหน้าจอไม่ได้! [{grab_error}]")
            self.root.after(0, self.log, "แก้: รัน Admin + เกม Borderless Windowed")
            self.running = False
            self.root.after(0, self._reset_ui); return

        self.root.after(0, self.log, "AFK Mode: รอ 3 วิ แล้วเริ่ม (สลับไปเกม!)")
        time.sleep(3)
        self.root.after(0, self.log, "เริ่มทำงาน...")

        round_num = 0
        learning_rois = {}

        while self.running:
            try:
                frame = grab(region)
                if frame is None:
                    time.sleep(0.1); continue

                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                # ═══ 1. สแกนทุกช่องพร้อมกัน ═══
                found_keys = []
                confidences = []
                rois = {}

                for i in range(num):
                    slot_img = get_slot(gray, i, num)
                    roi = prep(slot_img)
                    rois[i] = roi

                    if roi is not None and templates:
                        ch, cf = match_tmpl(roi)
                        if ch and cf > 0.60:
                            found_keys.append(ch)
                            confidences.append(cf)
                        else:
                            found_keys.append(None)
                    else:
                        found_keys.append(None)

                valid_keys = [k for k in found_keys if k is not None]

                # ═══ 2. เจอครบ → กดรวดเดียว! ═══
                if len(valid_keys) == num:
                    round_num += 1
                    seq_str = " ".join(k.upper() for k in valid_keys)
                    self.root.after(0, self.log, f"Round {round_num}: {seq_str}")

                    for key in valid_keys:
                        if not self.running: break
                        press_key(key)
                        self.session_keys += 1
                        time.sleep(kd + random.uniform(0.01, 0.05))

                    avg_cf = sum(confidences) / len(confidences) if confidences else 0
                    self.root.after(0, self._update_count, valid_keys[-1], avg_cf)
                    self.root.after(0, self.log, "กดครบ! รอดึงปลา...")

                    # ═══ 3. รอตัวละครดึงปลา ═══
                    time.sleep(5.0 + random.uniform(0.2, 0.8))

                    # ═══ 4. โยนเบ็ดใหม่ ═══
                    self.root.after(0, self.log, "โยนเบ็ดใหม่...")
                    press_key('e')  # *** แก้เป็นปุ่มโยนเบ็ดของเซิร์ฟคุณ ***
                    time.sleep(2.0 + random.uniform(0.3, 0.7))
                    continue

                # ═══ LEARN MODE: เจอไม่ครบ → ให้ user กดเอง ═══
                if not templates:
                    # ยังไม่มี template เลย → หา slot ที่มี content
                    learn_slot = -1
                    for i in range(num):
                        if rois.get(i) is not None:
                            learn_slot = i; break

                    if learn_slot >= 0:
                        if learn_slot not in learning_rois:
                            learning_rois[learn_slot] = rois[learn_slot]
                            self.root.after(0, self.log,
                                f"  Slot {learn_slot+1}: ? → กดปุ่มตามเกม!")

                        if self.user_pressed:
                            roi_save = learning_rois.get(learn_slot)
                            if roi_save is None:
                                roi_save = rois.get(learn_slot)
                            if roi_save is not None:
                                save_template(self.user_pressed, roi_save)
                                self.root.after(0, self.log,
                                    f"  จำ: {self.user_pressed.upper()} ✓ ({sum(len(v) for v in templates.values())} total)")
                                self.root.after(0, self._update_template_label)
                                learning_rois.pop(learn_slot, None)
                            self.user_pressed = None

                # มินิเกมยังไม่มา → รอ
                time.sleep(0.1)

            except Exception as e:
                self.root.after(0, self.log, f"Error: {e}")
                time.sleep(0.5)

        self.root.after(0, self._reset_ui)

    def _update_count(self, key, conf):
        self.lbl_count.config(text=str(self.session_keys))
        spd = self.session_keys / max(time.time()-self.session_start, 0.1)
        self.lbl_speed.config(text=f"{spd:.1f}/s  last:{key.upper()}")
        self._update_template_label()

    def _reset_ui(self):
        self.btn_start.config(text="  START  ", bg="#10b981")
        self.status_label.config(text="OFF", fg="#5a6a7e")
        self.log("Stopped")

    # ═══ Preview ═══
    def _preview_loop(self):
        if region:
            try:
                img = ImageGrab.grab(bbox=(region["left"],region["top"],
                    region["left"]+region["width"],region["top"]+region["height"]))
                w,h = img.size
                if w > 390: img = img.resize((390, max(int(h*390/w),1)), Image.NEAREST)
                photo = ImageTk.PhotoImage(img)
                self.preview_label.config(image=photo, text="")
                self.preview_label._photo = photo
            except: pass
        self.root.after(500, self._preview_loop)

    # ═══ Log ═══
    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        lines = int(self.log_text.index("end-1c").split(".")[0])
        if lines > 100: self.log_text.delete("1.0", f"{lines-100}.0")
        self.log_text.config(state=tk.DISABLED)

    def _quit(self):
        self.running = False
        time.sleep(0.1)
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
