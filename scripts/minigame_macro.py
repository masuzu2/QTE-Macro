"""
AutoFish by Herlove
FiveM Fishing Macro - OCR + SendInput Scancode
"""

import time, sys, os, json, threading, ctypes, random
import tkinter as tk
from tkinter import messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ═══ บังคับ Run as Administrator ═══
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

# ═══ Libs ═══
try:
    import keyboard as kb
    import mss
    import cv2
    import numpy as np
    from PIL import Image, ImageTk, ImageGrab
    import pytesseract
except ImportError as e:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("AutoFish",
        f"Missing: {e}\n\npip install keyboard mss opencv-python numpy Pillow pytesseract")
    sys.exit(1)

for p in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
          r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]:
    if os.path.exists(p): pytesseract.pytesseract.tesseract_cmd = p; break

# ═══ SendInput Scancode (Hardware level - FiveM) ═══
SCAN = {'q':0x10,'w':0x11,'e':0x12,'r':0x13,'a':0x1E,'s':0x1F,'d':0x20,'f':0x21,'space':0x39}

PUL = ctypes.POINTER(ctypes.c_ulong) if sys.platform == 'win32' else None
if sys.platform == 'win32':
    class KEYBDINPUT(ctypes.Structure):
        _fields_=[("wVk",ctypes.c_ushort),("wScan",ctypes.c_ushort),
                  ("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
    class HARDWAREINPUT(ctypes.Structure):
        _fields_=[("uMsg",ctypes.c_ulong),("wParamL",ctypes.c_short),("wParamH",ctypes.c_ushort)]
    class MOUSEINPUT(ctypes.Structure):
        _fields_=[("dx",ctypes.c_long),("dy",ctypes.c_long),("mouseData",ctypes.c_ulong),
                  ("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
    class INPUT_UNION(ctypes.Union):
        _fields_=[("ki",KEYBDINPUT),("mi",MOUSEINPUT),("hi",HARDWAREINPUT)]
    class INPUT(ctypes.Structure):
        _fields_=[("type",ctypes.c_ulong),("ii",INPUT_UNION)]

def press_key(key):
    sc = SCAN.get(key.lower())
    if not sc: return False
    if sys.platform != 'win32': return False
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

# ═══ Fast Capture (cache mss) ═══
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
        except:
            pass
    try:
        img = ImageGrab.grab(bbox=(left,top,left+w,top+h))
        if img: return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    except: pass
    return None

# ═══ Fast OCR ═══
VALID = set("qweasd")

def read_all(gray, num_keys):
    """อ่านทั้งแถวครั้งเดียว - เร็ว + แม่น"""
    if gray is None or gray.size < 100: return None

    # Sharpen + denoise ก่อน OCR (แก้แสงเกมเปลี่ยน + font เพี้ยน)
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

            if np.count_nonzero(t) > t.size // 2:
                t = cv2.bitwise_not(t)

            text = pytesseract.image_to_string(t,
                config='--psm 7 -c tessedit_char_whitelist=QWEASDqweasd').strip().lower()
            chars = [c for c in text if c in VALID]
            # Validate: ครบ + ทุกตัวต้องเป็น QWEASD จริง
            if len(chars) == num_keys and all(c in VALID for c in chars):
                return chars
        except: continue
    return None

def make_debug(frame, keys, num):
    d = frame.copy()
    h,w = d.shape[:2]; sw = w // num
    for i in range(num):
        x1, x2 = i*sw, min((i+1)*sw, w)
        if keys and i < len(keys) and keys[i]:
            cv2.rectangle(d, (x1+1,1), (x2-1,h-1), (0,255,0), 2)
            cv2.putText(d, keys[i].upper(), (x1+5,22),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0,255,0), 2)
        else:
            cv2.rectangle(d, (x1+1,1), (x2-1,h-1), (0,0,255), 1)
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

# ═══ Colors ═══
BG = "#0b1120"
CARD = "#111d35"
ACCENT = "#38bdf8"
GREEN = "#22c55e"
RED = "#ef4444"
DIM = "#475569"
WHITE = "#f1f5f9"

# ═══ App ═══
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AutoFish by Herlove")
        self.root.geometry("480x660")
        self.root.configure(bg=BG)
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        # icon
        try:
            ico = os.path.join(SCRIPT_DIR, "logo.png")
            if os.path.exists(ico):
                img = ImageTk.PhotoImage(Image.open(ico).resize((32,32)))
                self.root.iconphoto(True, img)
                self._ico = img
        except: pass

        self.running = False
        self.session_keys = 0
        self.session_start = 0
        self.debug_img = None
        self.fish_count = 0
        self.num_keys = tk.IntVar(value=load_cfg().get("num_keys", 5))
        self.key_delay = tk.IntVar(value=60)

        self._build()
        self._preview_loop()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        # ═══ Header with logo ═══
        hdr = tk.Frame(self.root, bg=BG)
        hdr.pack(fill=tk.X, padx=16, pady=(10,0))

        try:
            logo_path = os.path.join(SCRIPT_DIR, "logo.png")
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path).resize((48, 48), Image.LANCZOS)
                self._logo = ImageTk.PhotoImage(logo_img)
                tk.Label(hdr, image=self._logo, bg=BG).pack(side=tk.LEFT, padx=(0,10))
        except: pass

        title_f = tk.Frame(hdr, bg=BG)
        title_f.pack(side=tk.LEFT)
        tk.Label(title_f, text="AutoFish", bg=BG, fg=ACCENT,
                 font=("Segoe UI", 20, "bold")).pack(anchor="w")
        tk.Label(title_f, text="by Herlove  |  FiveM Fishing Macro", bg=BG, fg=DIM,
                 font=("Segoe UI", 9)).pack(anchor="w")

        self.lbl_st = tk.Label(hdr, text="OFF", bg=BG, fg=DIM,
                                font=("Consolas", 11, "bold"))
        self.lbl_st.pack(side=tk.RIGHT)

        # Gradient line
        line = tk.Canvas(self.root, bg=BG, height=3, highlightthickness=0)
        line.pack(fill=tk.X, padx=16, pady=(8,0))
        for i in range(448):
            t = i/448
            r = int(56+(34-56)*t); g = int(189+(211-189)*t); b = int(248+(238-248)*t)
            line.create_line(i+16, 0, i+16, 3, fill=f"#{r:02x}{g:02x}{b:02x}")

        # ═══ Preview (ใหญ่) ═══
        self.preview = tk.Label(self.root, bg="#060d1a", text="F6 เลือก Region",
                                 fg=DIM, font=("Segoe UI", 9), relief="flat")
        self.preview.pack(fill=tk.X, padx=16, pady=(10,6), ipady=35)

        # ═══ START ═══
        self.btn = tk.Button(self.root, text="START  (F5)", bg=GREEN, fg="white",
            font=("Segoe UI", 15, "bold"), relief="flat", pady=8, cursor="hand2",
            activebackground="#16a34a", command=self.toggle)
        self.btn.pack(fill=tk.X, padx=16, pady=(2,6))

        # ═══ Stats ═══
        sf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        sf.pack(fill=tk.X, padx=16, pady=(0,6))
        si = tk.Frame(sf, bg=CARD); si.pack(fill=tk.X, padx=12, pady=8)

        self.lbl_cnt = tk.Label(si, text="0", bg=CARD, fg=GREEN, font=("Consolas",28,"bold"))
        self.lbl_cnt.pack(side=tk.LEFT)
        tk.Label(si, text=" keys", bg=CARD, fg=DIM, font=("Segoe UI",10)).pack(side=tk.LEFT)

        self.lbl_fish = tk.Label(si, text="", bg=CARD, fg=ACCENT, font=("Consolas",11,"bold"))
        self.lbl_fish.pack(side=tk.RIGHT, padx=8)
        self.lbl_spd = tk.Label(si, text="", bg=CARD, fg=DIM, font=("Consolas",10))
        self.lbl_spd.pack(side=tk.RIGHT)

        # ═══ Region + Test ═══
        rf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        rf.pack(fill=tk.X, padx=16, pady=(0,6))
        ri = tk.Frame(rf, bg=CARD); ri.pack(fill=tk.X, padx=8, pady=6)

        tk.Button(ri, text="Select Region (F6)", bg="#1e3a5f", fg=ACCENT,
            font=("Segoe UI",9,"bold"), relief="flat", padx=10, pady=3,
            cursor="hand2", command=self.pick_region).pack(side=tk.LEFT, padx=2)
        tk.Button(ri, text="Test Read", bg="#14332a", fg=GREEN,
            font=("Segoe UI",9,"bold"), relief="flat", padx=8, pady=3,
            cursor="hand2", command=self.test_read).pack(side=tk.LEFT, padx=2)
        tk.Button(ri, text="Test Key", bg="#2d1b4e", fg="#a78bfa",
            font=("Segoe UI",9,"bold"), relief="flat", padx=8, pady=3,
            cursor="hand2", command=self.test_press).pack(side=tk.LEFT, padx=2)

        self.lbl_reg = tk.Label(ri, text="not set", bg=CARD, fg="#f97316", font=("Consolas",9))
        self.lbl_reg.pack(side=tk.RIGHT, padx=6)

        # ═══ Settings ═══
        stf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        stf.pack(fill=tk.X, padx=16, pady=(0,6))
        self._sl(stf, "Keys / Lane", self.num_keys, 2, 10, ACCENT)
        self._sl(stf, "Key Delay (ms)", self.key_delay, 30, 200, GREEN)

        # ═══ Log ═══
        tk.Label(self.root, text="LOG", bg=BG, fg=DIM, font=("Segoe UI",8,"bold"),
                 anchor="w").pack(fill=tk.X, padx=18, pady=(4,1))
        self.log_box = tk.Text(self.root, bg="#060d1a", fg=GREEN, font=("Consolas",9),
            height=6, relief="flat", wrap=tk.WORD, state=tk.DISABLED, padx=8, pady=4)
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0,10))

        # Footer
        tk.Label(self.root, text="F5 Start/Stop  |  F6 Region  |  Run as Admin",
                 bg=BG, fg="#1e293b", font=("Segoe UI",7)).pack(pady=(0,6))

        self.root.bind("<F5>", lambda e: self.toggle())
        self.root.bind("<F6>", lambda e: self.pick_region())

        # Status
        self.log("AutoFish by Herlove")
        self.log(f"Admin: {'Yes' if is_admin() else 'No!'}")
        try:
            v = pytesseract.get_tesseract_version()
            self.log(f"Tesseract {v}")
        except:
            self.log("ERROR: Tesseract not found!")

    def _sl(self, p, label, var, lo, hi, color):
        f = tk.Frame(p, bg=CARD); f.pack(fill=tk.X, padx=10, pady=2)
        tk.Label(f, text=label, bg=CARD, fg=WHITE, font=("Segoe UI",9)).pack(side=tk.LEFT)
        vl = tk.Label(f, text=str(var.get()), bg=CARD, fg=color, font=("Consolas",10,"bold"), width=4)
        vl.pack(side=tk.RIGHT)
        tk.Scale(f, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var, bg=CARD, fg=CARD,
            troughcolor=BG, highlightthickness=0, showvalue=False, length=130, sliderlength=14,
            activebackground=color,
            command=lambda v,l=vl:l.config(text=str(int(float(v))))).pack(side=tk.RIGHT,padx=4)

    # ═══ Region ═══
    def pick_region(self):
        global region
        if self.running: self.log("Stop first!"); return
        self.root.iconify(); time.sleep(0.3)
        sel = tk.Toplevel(self.root)
        sel.overrideredirect(True); sel.attributes("-topmost",True); sel.attributes("-alpha",0.25)
        sw,sh = sel.winfo_screenwidth(),sel.winfo_screenheight()
        sel.geometry(f"{sw}x{sh}+0+0"); sel.configure(bg="black")
        c = tk.Canvas(sel,bg="black",highlightthickness=0,cursor="cross")
        c.pack(fill=tk.BOTH,expand=True)
        c.create_rectangle(sw//2-220,15,sw//2+220,80,fill="black",outline="#38bdf8",width=2)
        c.create_text(sw//2,33,text="ลากครอบแค่ตัวอักษร (ไม่รวม counter)",fill="#38bdf8",font=("Segoe UI",14,"bold"))
        c.create_text(sw//2,58,text="ลากให้พอดีกรอบดำ  |  ESC ยกเลิก",fill="#94a3b8",font=("Segoe UI",10))
        pos = c.create_text(sw//2,100,text="",fill="#38bdf8",font=("Consolas",12,"bold"))
        st = {"sx":0,"sy":0,"r":None}
        c.bind("<Motion>",lambda e:c.itemconfig(pos,text=f"X:{e.x}  Y:{e.y}"))
        def _p(e):
            st["sx"],st["sy"]=e.x,e.y
            if st["r"]: c.delete(st["r"])
            st["r"]=c.create_rectangle(e.x,e.y,e.x,e.y,outline="#38bdf8",width=3)
        def _d(e):
            c.coords(st["r"],st["sx"],st["sy"],e.x,e.y); c.delete("sz")
            c.create_text((st["sx"]+e.x)//2,min(st["sy"],e.y)-14,
                text=f"{abs(e.x-st['sx'])}x{abs(e.y-st['sy'])}",fill="#38bdf8",font=("Consolas",13,"bold"),tags="sz")
        def _r(e):
            global region
            x1,y1=min(st["sx"],e.x),min(st["sy"],e.y); x2,y2=max(st["sx"],e.x),max(st["sy"],e.y)
            if (x2-x1)>10 and (y2-y1)>10:
                region={"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}; save_cfg()
                self.lbl_reg.config(text=f"{x2-x1}x{y2-y1}",fg=GREEN)
                self.log(f"Region: {x2-x1}x{y2-y1} @ ({x1},{y1})")
            sel.destroy(); self.root.deiconify()
        c.bind("<ButtonPress-1>",_p); c.bind("<B1-Motion>",_d); c.bind("<ButtonRelease-1>",_r)
        sel.bind("<Escape>",lambda e:(sel.destroy(),self.root.deiconify()))
        sel.after(50,sel.focus_force)

    def test_read(self):
        if not region: self.log("เลือก region ก่อน!"); return
        t0 = time.time()
        frame = grab(region)
        if frame is None: self.log("จับจอไม่ได้!"); return
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        keys = read_all(gray, self.num_keys.get())
        ms = int((time.time()-t0)*1000)
        self.debug_img = make_debug(frame, keys, self.num_keys.get())
        self._show_debug()
        if keys:
            self.log(f"Read ({ms}ms): {' '.join(k.upper() for k in keys)}")
        else:
            self.log(f"อ่านไม่ครบ ({ms}ms) ลองลาก region ใหม่")

    def test_press(self):
        self.log("กด D+S ใน 3 วิ (คลิกเกมก่อน!)")
        def _do():
            time.sleep(3)
            press_key('d'); time.sleep(0.2); press_key('s')
            self.root.after(0,self.log,"Pressed D S!")
        threading.Thread(target=_do,daemon=True).start()

    # ═══ Toggle ═══
    def toggle(self):
        if self.running: self.running = False
        else:
            if not region: messagebox.showwarning("AutoFish","กด F6 เลือก Region!"); return
            self.running = True; save_cfg(num_keys=self.num_keys.get())
            self.btn.config(text="STOP  (F5)", bg=RED, activebackground="#dc2626")
            self.lbl_st.config(text="ON", fg=GREEN)
            self.log("สลับไปเกมใน 3 วิ!")
            threading.Thread(target=self._run, daemon=True).start()

    # ═══ Main Loop ═══
    def _run(self):
        num = self.num_keys.get()
        kd = self.key_delay.get() / 1000
        self.session_keys = 0
        self.session_start = time.time()
        self.fish_count = 0

        test = grab(region)
        if test is None:
            self.root.after(0,self.log,"จับจอไม่ได้!")
            self.running = False; self.root.after(0,self._reset); return

        time.sleep(3)
        self.root.after(0,self.log,"เริ่มสแกน...")
        last_seq = ""

        while self.running:
            try:
                frame = grab(region)
                if frame is None: time.sleep(0.05); continue
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                keys = read_all(gray, num)
                self.debug_img = make_debug(frame, keys, num)

                if keys and len(keys) == num:
                    seq = "".join(keys)
                    if seq != last_seq:
                        self.fish_count += 1
                        display = " ".join(k.upper() for k in keys)
                        self.root.after(0,self.log,f"#{self.fish_count} {display}")

                        for key in keys:
                            if not self.running: break
                            press_key(key)
                            self.session_keys += 1
                            time.sleep(kd + random.uniform(0.02, 0.08))

                        self.root.after(0, self._update_stats)
                        last_seq = seq
                        time.sleep(0.8 + random.uniform(0.1, 0.3))
                    else:
                        time.sleep(0.04)
                else:
                    time.sleep(0.06)

            except Exception as e:
                self.root.after(0,self.log,f"Error: {e}")
                time.sleep(0.5)

        self.root.after(0,self._reset)

    def _update_stats(self):
        self.lbl_cnt.config(text=str(self.session_keys))
        elapsed = max(time.time() - self.session_start, 0.1)
        self.lbl_spd.config(text=f"{self.session_keys/elapsed:.1f}/s")
        self.lbl_fish.config(text=f"Fish #{self.fish_count}")

    def _reset(self):
        self.btn.config(text="START  (F5)", bg=GREEN, activebackground="#16a34a")
        self.lbl_st.config(text="OFF", fg=DIM)
        self.log("Stopped")

    def _show_debug(self):
        if self.debug_img is None: return
        try:
            rgb = cv2.cvtColor(self.debug_img, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(rgb)
            w,h = img.size
            nw = 448; nh = max(int(h*nw/w), 20)
            img = img.resize((nw, nh), Image.NEAREST)
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
                    w,h = img.size
                    nw = 448; nh = max(int(h*nw/w), 20)
                    img = img.resize((nw,nh), Image.NEAREST)
                    photo = ImageTk.PhotoImage(img)
                    self.preview.config(image=photo, text="")
                    self.preview.image = photo
            except: pass
        self.root.after(150, self._preview_loop)

    def log(self,msg):
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END,f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_box.see(tk.END)
        lines = int(self.log_box.index("end-1c").split(".")[0])
        if lines > 80: self.log_box.delete("1.0",f"{lines-80}.0")
        self.log_box.config(state=tk.DISABLED)

    def _quit(self): self.running=False; time.sleep(0.1); self.root.destroy()
    def run(self): self.root.mainloop()

if __name__ == "__main__": App().run()
