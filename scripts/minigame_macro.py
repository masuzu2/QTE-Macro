"""
QTE Macro - No Template Edition
ไม่ต้อง setup ไม่ต้อง learn ไม่ต้อง template
สร้างตัวอักษร Q W E A S D จาก font ในเครื่อง + เลือกตัวที่ match สูงสุด
"""

import time, sys, os, json, threading, ctypes, random
import tkinter as tk
from tkinter import messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

try:
    import keyboard as kb
    import mss
    import cv2
    import numpy as np
    from PIL import Image, ImageTk, ImageGrab, ImageDraw, ImageFont
except ImportError as e:
    print(f"Missing: {e}\npip install keyboard mss opencv-python numpy Pillow")
    sys.exit(1)

# ═══════════════════════════════════
# SendInput (Windows game key press)
# ═══════════════════════════════════
SCAN = {'q':0x10,'w':0x11,'e':0x12,'r':0x13,'a':0x1E,'s':0x1F,'d':0x20,'f':0x21,
        'z':0x2C,'x':0x2D,'c':0x2E,'v':0x2F,'1':0x02,'2':0x03,'3':0x04,'4':0x05,
        'space':0x39}

def press_key(key):
    key = key.lower()
    sc = SCAN.get(key)
    if sc is None: return False
    if sys.platform == 'win32':
        try:
            PUL = ctypes.POINTER(ctypes.c_ulong)
            class KI(ctypes.Structure):
                _fields_ = [("wVk",ctypes.c_ushort),("wScan",ctypes.c_ushort),
                            ("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
            class HI(ctypes.Structure):
                _fields_ = [("uMsg",ctypes.c_ulong),("wParamL",ctypes.c_short),("wParamH",ctypes.c_ushort)]
            class MI(ctypes.Structure):
                _fields_ = [("dx",ctypes.c_long),("dy",ctypes.c_long),("mouseData",ctypes.c_ulong),
                            ("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
            class IU(ctypes.Union):
                _fields_ = [("ki",KI),("mi",MI),("hi",HI)]
            class INP(ctypes.Structure):
                _fields_ = [("type",ctypes.c_ulong),("ii",IU)]
            ex = ctypes.c_ulong(0)
            ii = IU(); ii.ki = KI(0, sc, 0x0008, 0, ctypes.pointer(ex))
            x = INP(ctypes.c_ulong(1), ii)
            ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))
            time.sleep(0.03)
            ii2 = IU(); ii2.ki = KI(0, sc, 0x0008|0x0002, 0, ctypes.pointer(ex))
            x2 = INP(ctypes.c_ulong(1), ii2)
            ctypes.windll.user32.SendInput(1, ctypes.pointer(x2), ctypes.sizeof(x2))
            return True
        except: pass
    try:
        import pyautogui; pyautogui.PAUSE = 0; pyautogui.press(key); return True
    except: pass
    try: kb.press_and_release(key); return True
    except: pass
    return False

# ═══════════════════════════════════
# Screen Capture
# ═══════════════════════════════════
grab_error = ""
def grab(r):
    global grab_error
    left,top,w,h = int(r["left"]),int(r["top"]),int(r["width"]),int(r["height"])
    try:
        with mss.mss() as s:
            arr = np.array(s.grab({"left":left,"top":top,"width":w,"height":h}))
            if arr.size > 0: return cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)
    except Exception as e: grab_error = str(e)
    try:
        img = ImageGrab.grab(bbox=(left,top,left+w,top+h))
        if img: return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    except Exception as e: grab_error = str(e)
    return None

# ═══════════════════════════════════
# Auto Key Detector - ไม่ต้อง template!
# สร้างตัวอักษรจาก font ในเครื่อง + เลือก best match
# ═══════════════════════════════════
class KeyDetector:
    def __init__(self, chars="qweasd"):
        self.chars = chars
        self.templates = {}  # char -> [gray images normalized h=48]
        self.TH = 48
        self._build()

    def _build(self):
        """สร้าง template จากทุก font ที่มี หลายขนาด หลาย weight"""
        fonts = self._find_fonts()
        for ch in self.chars:
            self.templates[ch] = []
            for fp in fonts:
                for sz in [28, 36, 44, 52, 64, 76, 88]:
                    for bold in [False, True]:
                        for case in [ch.upper(), ch.lower()]:
                            img = self._render(case, fp, sz, bold)
                            if img is not None:
                                h, w = img.shape[:2]
                                nw = max(int(w * self.TH / h), 3)
                                self.templates[ch].append(cv2.resize(img, (nw, self.TH)))

    def _find_fonts(self):
        fonts = []
        for n in ["arialbd.ttf","arial.ttf","impact.ttf","calibrib.ttf","segoeuib.ttf",
                   "segoeui.ttf","tahomabd.ttf","tahoma.ttf","verdanab.ttf","consolab.ttf",
                   "trebucbd.ttf","courbd.ttf"]:
            p = os.path.join(os.environ.get("WINDIR","C:\\Windows"),"Fonts",n)
            if os.path.exists(p): fonts.append(p)
        return fonts[:8] if fonts else [None]

    def _render(self, ch, fp, sz, bold=False):
        try:
            csz = sz + 16
            img = Image.new("L",(csz,csz),0)
            d = ImageDraw.Draw(img)
            f = ImageFont.truetype(fp, sz) if fp else ImageFont.load_default()
            bb = d.textbbox((0,0),ch,font=f); tw,th = bb[2]-bb[0],bb[3]-bb[1]
            d.text(((csz-tw)//2-bb[0],(csz-th)//2-bb[1]),ch,fill=255,font=f)
            a = np.array(img)
            if bold:
                a = cv2.dilate(a, cv2.getStructuringElement(cv2.MORPH_ELLIPSE,(3,3)))
            cs = np.argwhere(a > 40)
            if len(cs) < 5: return None
            y0,x0 = cs.min(0); y1,x1 = cs.max(0)
            cr = a[max(0,y0-1):y1+2, max(0,x0-1):x1+2]
            if cr.shape[0]<4 or cr.shape[1]<4: return None
            _, b = cv2.threshold(cr, 80, 255, cv2.THRESH_BINARY)
            return b
        except: return None

    def detect(self, gray_slot):
        """ตรวจจับตัวอักษร: เลือกตัวที่ match สูงสุด (ไม่ใช่ threshold)
        return: (char, confidence) หรือ (None, 0)"""
        roi = self._prep(gray_slot)
        if roi is None: return None, 0

        h,w = roi.shape[:2]
        if h < 3 or w < 3: return None, 0
        nw = max(int(w * self.TH / h), 3)
        roi_n = cv2.resize(roi, (nw, self.TH))

        # หาตัวที่ score สูงสุดจากทุกตัวอักษร
        best_char, best_score = None, -1

        for ch, tmpls in self.templates.items():
            for t in tmpls:
                tw = t.shape[1]
                if tw > roi_n.shape[1] * 1.8 or tw < roi_n.shape[1] * 0.3:
                    continue

                target = roi_n
                if tw > roi_n.shape[1] or t.shape[0] > roi_n.shape[0]:
                    px = max(0, (tw - roi_n.shape[1])//2 + 4)
                    py = max(0, (t.shape[0] - roi_n.shape[0])//2 + 4)
                    target = cv2.copyMakeBorder(roi_n, py,py,px,px, cv2.BORDER_CONSTANT, value=0)

                try:
                    res = cv2.matchTemplate(target, t, cv2.TM_CCOEFF_NORMED)
                    _, mx, _, _ = cv2.minMaxLoc(res)
                    if mx > best_score:
                        best_char, best_score = ch, mx
                except: continue

        # ต้อง score อย่างน้อย 0.25 ถึงจะถือว่าเจอตัวอักษร (ป้องกันจอว่าง)
        if best_score < 0.25:
            return None, 0

        return best_char, best_score

    def detect_lane(self, gray, num_keys):
        """อ่านทั้ง lane ทุก slot พร้อมกัน"""
        h, w = gray.shape[:2]
        sw = w // num_keys
        results = []
        for i in range(num_keys):
            pad = max(sw // 10, 2)
            slot = gray[:, max(0, i*sw - pad):min(w, (i+1)*sw + pad)]
            ch, cf = self.detect(slot)
            results.append((ch, cf))
        return results

    def _prep(self, gray):
        """Threshold + crop ตัวอักษร (ลองหลายวิธี เลือกดีสุด)"""
        if gray is None or gray.size < 20: return None
        h, w = gray.shape[:2]
        if h < 20 or w < 10:
            s = max(30/max(h,1), 30/max(w,1), 1)
            gray = cv2.resize(gray, None, fx=s, fy=s, interpolation=cv2.INTER_CUBIC)

        best, best_s = None, 0
        for tval in [0, 100, 140, 180, 220]:
            try:
                if tval == 0:
                    _, t = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                else:
                    _, t = cv2.threshold(gray, tval, 255, cv2.THRESH_BINARY)
                ti = cv2.bitwise_not(t)
                r = t if np.count_nonzero(t) < np.count_nonzero(ti) else ti
                cs = np.argwhere(r > 127)
                if len(cs) < 5: continue
                y0, x0 = cs.min(0); y1, x1 = cs.max(0)
                c = r[max(0,y0-1):y1+2, max(0,x0-1):x1+2]
                if c.shape[0] < 4 or c.shape[1] < 4: continue
                ratio = np.count_nonzero(c > 127) / max(c.size, 1)
                if 0.05 < ratio < 0.65 and c.size > best_s:
                    best = c; best_s = c.size
            except: continue
        return best

# ═══════════════════════════════════
# Config
# ═══════════════════════════════════
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

# สร้าง detector
detector = KeyDetector("qweasd")
n_templates = sum(len(v) for v in detector.templates.values())
print(f"Auto templates: {n_templates}")

# ═══════════════════════════════════
# App
# ═══════════════════════════════════
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QTE Macro")
        self.root.geometry("420x520")
        self.root.configure(bg="#0a0e17")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.running = False
        self.session_keys = 0
        self.session_start = 0
        self.num_keys = tk.IntVar(value=load_cfg().get("num_keys", 5))
        self.key_delay = tk.IntVar(value=50)
        self._build()
        self._preview_loop()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        BG, CARD, DIM = "#0a0e17", "#111a2b", "#5a6a7e"

        tk.Label(self.root, text="QTE Macro", bg=BG, fg="#22d3ee",
                 font=("Segoe UI",16,"bold")).pack(pady=(10,0))
        tk.Label(self.root, text=f"No Setup! Auto Detect Q W E A S D ({n_templates} templates)", bg=BG,
                 fg=DIM, font=("Segoe UI",8)).pack()

        self.preview_label = tk.Label(self.root, bg="#0c1220", text="กด F6 เลือก region",
                                       fg=DIM, font=("Segoe UI",8), height=3)
        self.preview_label.pack(fill=tk.X, padx=14, pady=(8,4))

        self.status_label = tk.Label(self.root, text="OFF", bg=BG, fg=DIM, font=("Consolas",10,"bold"))
        self.status_label.pack()

        self.btn = tk.Button(self.root, text="  START (F5)  ", bg="#10b981", fg="white",
            font=("Segoe UI",14,"bold"), relief="flat", pady=6, command=self.toggle)
        self.btn.pack(fill=tk.X, padx=14, pady=6)

        f1 = tk.Frame(self.root, bg=CARD); f1.pack(fill=tk.X, padx=14, pady=2)
        self.lbl_count = tk.Label(f1, text="0", bg=CARD, fg="#10b981", font=("Consolas",24,"bold"))
        self.lbl_count.pack(side=tk.LEFT, padx=10, pady=6)
        self.lbl_info = tk.Label(f1, text="", bg=CARD, fg=DIM, font=("Consolas",10))
        self.lbl_info.pack(side=tk.LEFT)

        tk.Label(self.root, text="REGION", bg=BG, fg=DIM, font=("Segoe UI",8,"bold")).pack(fill=tk.X, padx=16, pady=(8,2))
        f2 = tk.Frame(self.root, bg=CARD); f2.pack(fill=tk.X, padx=14, pady=2)
        tk.Button(f2, text="Select (F6)", bg="#1e3a5f", fg="#22d3ee", font=("Segoe UI",9,"bold"),
            relief="flat", padx=10, pady=4, command=self.pick_region).pack(side=tk.LEFT, padx=6, pady=6)
        tk.Button(f2, text="Test Key", bg="#3b1764", fg="#a78bfa", font=("Segoe UI",9,"bold"),
            relief="flat", padx=8, pady=4, command=self.test_press).pack(side=tk.LEFT, padx=2, pady=6)
        tk.Button(f2, text="Test Read", bg="#1a3320", fg="#22ff88", font=("Segoe UI",9,"bold"),
            relief="flat", padx=8, pady=4, command=self.test_read).pack(side=tk.LEFT, padx=2, pady=6)
        self.lbl_region = tk.Label(f2, text="not set", bg=CARD, fg="#f97316", font=("Consolas",9))
        self.lbl_region.pack(side=tk.RIGHT, padx=8)

        tk.Label(self.root, text="SETTINGS", bg=BG, fg=DIM, font=("Segoe UI",8,"bold")).pack(fill=tk.X, padx=16, pady=(8,2))
        sf = tk.Frame(self.root, bg=CARD); sf.pack(fill=tk.X, padx=14, pady=2)
        self._slider(sf, "Keys/Lane", self.num_keys, 2, 10)
        self._slider(sf, "Key Delay ms", self.key_delay, 20, 200)

        tk.Label(self.root, text="LOG", bg=BG, fg=DIM, font=("Segoe UI",8,"bold")).pack(fill=tk.X, padx=16, pady=(8,2))
        self.log_text = tk.Text(self.root, bg="#070c18", fg="#22ff88", font=("Consolas",9),
            height=7, relief="flat", wrap=tk.WORD, state=tk.DISABLED, padx=8, pady=4)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=14, pady=(2,10))

        self.root.bind("<F5>", lambda e: self.toggle())
        self.root.bind("<F6>", lambda e: self.pick_region())
        self.log("Ready! ไม่ต้อง setup แค่ Select Region → Start")

    def _slider(self, p, label, var, lo, hi):
        f = tk.Frame(p, bg="#111a2b"); f.pack(fill=tk.X, padx=8, pady=2)
        tk.Label(f, text=label+":", bg="#111a2b", fg="#eee", font=("Segoe UI",9)).pack(side=tk.LEFT)
        vl = tk.Label(f, text=str(var.get()), bg="#111a2b", fg="#22d3ee", font=("Consolas",9,"bold"), width=5)
        vl.pack(side=tk.RIGHT)
        tk.Scale(f, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var,
                 bg="#111a2b", fg="#111a2b", troughcolor="#0a0e17", highlightthickness=0,
                 showvalue=False, length=140, sliderlength=14,
                 command=lambda v, l=vl: l.config(text=str(int(float(v))))).pack(side=tk.RIGHT, padx=4)

    def pick_region(self):
        global region
        if self.running: self.log("Stop first!"); return
        self.root.iconify(); time.sleep(0.3)
        sel = tk.Toplevel(self.root)
        sel.overrideredirect(True); sel.attributes("-topmost",True); sel.attributes("-alpha",0.25)
        sw,sh = sel.winfo_screenwidth(), sel.winfo_screenheight()
        sel.geometry(f"{sw}x{sh}+0+0"); sel.configure(bg="black")
        c = tk.Canvas(sel, bg="black", highlightthickness=0, cursor="cross")
        c.pack(fill=tk.BOTH, expand=True)
        c.create_rectangle(sw//2-200,15,sw//2+200,75, fill="black", outline="#00ff88", width=2)
        c.create_text(sw//2,35, text="ลากครอบแถวตัวอักษร (ไม่รวม counter)", fill="#00ff88", font=("Segoe UI",14,"bold"))
        c.create_text(sw//2,60, text="ESC = ยกเลิก", fill="#aaa", font=("Segoe UI",10))
        pos = c.create_text(sw//2,95, text="", fill="#00ff88", font=("Consolas",12,"bold"))
        st = {"sx":0,"sy":0,"r":None}
        c.bind("<Motion>", lambda e: c.itemconfig(pos, text=f"X:{e.x} Y:{e.y}"))
        def _p(e):
            st["sx"],st["sy"]=e.x,e.y
            if st["r"]: c.delete(st["r"])
            st["r"]=c.create_rectangle(e.x,e.y,e.x,e.y, outline="#00ff88", width=3)
        def _d(e):
            c.coords(st["r"],st["sx"],st["sy"],e.x,e.y); c.delete("sz")
            c.create_text((st["sx"]+e.x)//2,min(st["sy"],e.y)-14,
                text=f"{abs(e.x-st['sx'])}x{abs(e.y-st['sy'])}", fill="#00ff88", font=("Consolas",13,"bold"), tags="sz")
        def _r(e):
            global region
            x1,y1=min(st["sx"],e.x),min(st["sy"],e.y); x2,y2=max(st["sx"],e.x),max(st["sy"],e.y)
            if (x2-x1)>10 and (y2-y1)>10:
                region={"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}; save_cfg()
                self.lbl_region.config(text=f"{x2-x1}x{y2-y1}", fg="#10b981")
                self.log(f"Region: {x2-x1}x{y2-y1} @ ({x1},{y1})")
            sel.destroy(); self.root.deiconify()
        c.bind("<ButtonPress-1>",_p); c.bind("<B1-Motion>",_d); c.bind("<ButtonRelease-1>",_r)
        sel.bind("<Escape>", lambda e: (sel.destroy(), self.root.deiconify()))
        sel.after(50, sel.focus_force)

    def test_press(self):
        self.log("กด Q ใน 3 วิ... (คลิกเกมก่อน!)")
        def _do():
            time.sleep(3)
            ok = press_key('q')
            self.root.after(0, self.log, f"Press Q: {'OK' if ok else 'FAIL'}")
        threading.Thread(target=_do, daemon=True).start()

    def test_read(self):
        if not region: self.log("เลือก region ก่อน!"); return
        frame = grab(region)
        if frame is None: self.log(f"จับจอไม่ได้! {grab_error}"); return
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        results = detector.detect_lane(gray, self.num_keys.get())
        parts = []
        for i, (ch, cf) in enumerate(results):
            parts.append(f"[{ch.upper() if ch else '?'}:{cf:.0%}]")
        self.log(f"Read: {' '.join(parts)}")

    def toggle(self):
        if self.running:
            self.running = False
        else:
            if not region: messagebox.showwarning("QTE","กด F6 เลือก Region!"); return
            self.running = True
            save_cfg(num_keys=self.num_keys.get())
            self.btn.config(text="  STOP  ", bg="#ef4444")
            self.status_label.config(text="RUNNING", fg="#10b981")
            self.log("AFK Mode: สลับไปเกมใน 3 วินาที!")
            threading.Thread(target=self._run, daemon=True).start()

    # ═══ AFK Loop ═══
    def _run(self):
        num = self.num_keys.get()
        kd = self.key_delay.get() / 1000
        self.session_keys = 0
        self.session_start = time.time()

        test = grab(region)
        if test is None:
            self.root.after(0, self.log, f"จับจอไม่ได้! {grab_error}")
            self.running = False
            self.root.after(0, self._reset_ui); return

        time.sleep(3)
        self.root.after(0, self.log, "เริ่ม!")
        round_num = 0
        last_seq = ""

        while self.running:
            try:
                frame = grab(region)
                if frame is None: time.sleep(0.1); continue
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                # สแกนทุก slot
                results = detector.detect_lane(gray, num)
                found = [ch for ch, cf in results if ch is not None]

                # เจอครบ → กดรวด!
                if len(found) == num:
                    seq = "".join(found)
                    if seq != last_seq:
                        round_num += 1
                        self.root.after(0, self.log, f"R{round_num}: {' '.join(k.upper() for k in found)}")

                        for key in found:
                            if not self.running: break
                            press_key(key)
                            self.session_keys += 1
                            time.sleep(kd + random.uniform(0.01, 0.05))

                        self.root.after(0, self.lbl_count.config, {"text": str(self.session_keys)})
                        self.root.after(0, self.lbl_info.config,
                            {"text": f"{self.session_keys/(time.time()-self.session_start):.1f}/s"})
                        last_seq = seq

                        # รอดึงปลา
                        self.root.after(0, self.log, "รอดึงปลา...")
                        time.sleep(5.0 + random.uniform(0.2, 0.8))

                        # โยนเบ็ดใหม่
                        if self.running:
                            self.root.after(0, self.log, "โยนเบ็ดใหม่")
                            press_key('e')  # *** แก้ปุ่มตามเซิร์ฟ ***
                            time.sleep(2.0 + random.uniform(0.3, 0.7))
                            last_seq = ""
                    else:
                        time.sleep(0.05)
                else:
                    time.sleep(0.1)

            except Exception as e:
                self.root.after(0, self.log, f"Error: {e}")
                time.sleep(0.5)

        self.root.after(0, self._reset_ui)

    def _reset_ui(self):
        self.btn.config(text="  START (F5)  ", bg="#10b981")
        self.status_label.config(text="OFF", fg="#5a6a7e")
        self.log("Stopped")

    def _preview_loop(self):
        if region:
            try:
                img = ImageGrab.grab(bbox=(region["left"],region["top"],
                    region["left"]+region["width"],region["top"]+region["height"]))
                w,h = img.size
                if w > 390: img = img.resize((390,max(int(h*390/w),1)), Image.NEAREST)
                photo = ImageTk.PhotoImage(img)
                self.preview_label.config(image=photo, text="")
                self.preview_label._photo = photo
            except: pass
        self.root.after(500, self._preview_loop)

    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        lines = int(self.log_text.index("end-1c").split(".")[0])
        if lines > 100: self.log_text.delete("1.0", f"{lines-100}.0")
        self.log_text.config(state=tk.DISABLED)

    def _quit(self):
        self.running = False; time.sleep(0.1); self.root.destroy()

    def run(self): self.root.mainloop()

if __name__ == "__main__":
    App().run()
