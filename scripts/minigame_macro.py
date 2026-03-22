"""
QTE Macro - OCR Edition (Tesseract)
อ่านตัวอักษรจริงด้วย AI ไม่ต้อง template ไม่ต้อง learn
ติดตั้ง Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
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
    from PIL import Image, ImageTk, ImageGrab
    import pytesseract
except ImportError as e:
    print(f"Missing: {e}\npip install keyboard mss opencv-python numpy Pillow pytesseract")
    sys.exit(1)

# Tesseract path
for p in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
          r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
          r"C:\Users\{}\AppData\Local\Tesseract-OCR\tesseract.exe".format(os.environ.get("USERNAME",""))]:
    if os.path.exists(p):
        pytesseract.pytesseract.tesseract_cmd = p
        break

# ═══ SendInput ═══
SCAN = {'q':0x10,'w':0x11,'e':0x12,'r':0x13,'a':0x1E,'s':0x1F,'d':0x20,'f':0x21,
        'z':0x2C,'x':0x2D,'c':0x2E,'v':0x2F,'space':0x39}

def press_key(key):
    key = key.lower()
    sc = SCAN.get(key)
    if sc is None: return False
    if sys.platform == 'win32':
        try:
            PUL = ctypes.POINTER(ctypes.c_ulong)
            class KI(ctypes.Structure):
                _fields_=[("wVk",ctypes.c_ushort),("wScan",ctypes.c_ushort),
                          ("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
            class HI(ctypes.Structure):
                _fields_=[("uMsg",ctypes.c_ulong),("wParamL",ctypes.c_short),("wParamH",ctypes.c_ushort)]
            class MI(ctypes.Structure):
                _fields_=[("dx",ctypes.c_long),("dy",ctypes.c_long),("mouseData",ctypes.c_ulong),
                          ("dwFlags",ctypes.c_ulong),("time",ctypes.c_ulong),("dwExtraInfo",PUL)]
            class IU(ctypes.Union):
                _fields_=[("ki",KI),("mi",MI),("hi",HI)]
            class INP(ctypes.Structure):
                _fields_=[("type",ctypes.c_ulong),("ii",IU)]
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
    try: import pyautogui; pyautogui.PAUSE=0; pyautogui.press(key); return True
    except: pass
    try: kb.press_and_release(key); return True
    except: pass
    return False

# ═══ Capture ═══
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

# ═══ OCR อ่านตัวอักษร ═══
VALID_KEYS = set("qweasd")
OCR_CONFIG = '--psm 10 -c tessedit_char_whitelist=QWEASDqweasd'

def read_slot(gray_slot):
    """อ่าน 1 ตัวอักษรจาก slot ด้วย OCR"""
    if gray_slot is None or gray_slot.size < 20:
        return None

    # ขยาย 3 เท่า ให้ OCR อ่านง่าย
    big = cv2.resize(gray_slot, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)

    # ลองหลาย preprocessing แล้ว vote
    candidates = []

    for method in range(4):
        try:
            if method == 0:
                # Otsu
                _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            elif method == 1:
                # Otsu inverted
                _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                t = cv2.bitwise_not(t)
            elif method == 2:
                # High contrast + threshold
                c = cv2.convertScaleAbs(big, alpha=2.5, beta=-80)
                _, t = cv2.threshold(c, 150, 255, cv2.THRESH_BINARY)
            elif method == 3:
                # Adaptive
                t = cv2.adaptiveThreshold(big, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                           cv2.THRESH_BINARY, 15, 5)

            text = pytesseract.image_to_string(t, config=OCR_CONFIG).strip().lower()
            for ch in text:
                if ch in VALID_KEYS:
                    candidates.append(ch)
                    break
        except:
            continue

    if not candidates:
        return None

    # Vote: เลือกตัวที่เจอมากที่สุด
    from collections import Counter
    most = Counter(candidates).most_common(1)[0]
    return most[0]

def read_lane(gray, num_keys):
    """อ่านทั้ง lane ทุก slot"""
    h, w = gray.shape[:2]
    sw = w // num_keys
    results = []
    for i in range(num_keys):
        pad = max(sw // 10, 2)
        slot = gray[:, max(0, i*sw-pad):min(w, (i+1)*sw+pad)]
        ch = read_slot(slot)
        results.append(ch)
    return results

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

# ═══ App ═══
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QTE Macro (OCR)")
        self.root.geometry("420x500")
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
        BG,CARD,DIM = "#0a0e17","#111a2b","#5a6a7e"
        tk.Label(self.root,text="QTE Macro [OCR]",bg=BG,fg="#22d3ee",font=("Segoe UI",16,"bold")).pack(pady=(10,0))
        tk.Label(self.root,text="อ่านตัวอักษรด้วย AI ไม่ต้อง template",bg=BG,fg=DIM,font=("Segoe UI",8)).pack()
        self.preview = tk.Label(self.root,bg="#0c1220",text="กด F6 เลือก region",fg=DIM,font=("Segoe UI",8),height=3)
        self.preview.pack(fill=tk.X,padx=14,pady=(8,4))
        self.lbl_status = tk.Label(self.root,text="OFF",bg=BG,fg=DIM,font=("Consolas",10,"bold"))
        self.lbl_status.pack()
        self.btn = tk.Button(self.root,text="  START (F5)  ",bg="#10b981",fg="white",
            font=("Segoe UI",14,"bold"),relief="flat",pady=6,command=self.toggle)
        self.btn.pack(fill=tk.X,padx=14,pady=6)
        f1 = tk.Frame(self.root,bg=CARD); f1.pack(fill=tk.X,padx=14,pady=2)
        self.lbl_count = tk.Label(f1,text="0",bg=CARD,fg="#10b981",font=("Consolas",24,"bold"))
        self.lbl_count.pack(side=tk.LEFT,padx=10,pady=6)
        self.lbl_info = tk.Label(f1,text="",bg=CARD,fg=DIM,font=("Consolas",10))
        self.lbl_info.pack(side=tk.LEFT)

        tk.Label(self.root,text="REGION",bg=BG,fg=DIM,font=("Segoe UI",8,"bold")).pack(fill=tk.X,padx=16,pady=(8,2))
        f2 = tk.Frame(self.root,bg=CARD); f2.pack(fill=tk.X,padx=14,pady=2)
        tk.Button(f2,text="Select (F6)",bg="#1e3a5f",fg="#22d3ee",font=("Segoe UI",9,"bold"),
            relief="flat",padx=10,pady=4,command=self.pick_region).pack(side=tk.LEFT,padx=6,pady=6)
        tk.Button(f2,text="Test Read",bg="#1a3320",fg="#22ff88",font=("Segoe UI",9,"bold"),
            relief="flat",padx=8,pady=4,command=self.test_read).pack(side=tk.LEFT,padx=2,pady=6)
        tk.Button(f2,text="Test Key",bg="#3b1764",fg="#a78bfa",font=("Segoe UI",9,"bold"),
            relief="flat",padx=8,pady=4,command=self.test_press).pack(side=tk.LEFT,padx=2,pady=6)
        self.lbl_region = tk.Label(f2,text="not set",bg=CARD,fg="#f97316",font=("Consolas",9))
        self.lbl_region.pack(side=tk.RIGHT,padx=8)

        tk.Label(self.root,text="SETTINGS",bg=BG,fg=DIM,font=("Segoe UI",8,"bold")).pack(fill=tk.X,padx=16,pady=(8,2))
        sf = tk.Frame(self.root,bg=CARD); sf.pack(fill=tk.X,padx=14,pady=2)
        self._slider(sf,"Keys/Lane",self.num_keys,2,10)
        self._slider(sf,"Key Delay ms",self.key_delay,20,200)

        tk.Label(self.root,text="LOG",bg=BG,fg=DIM,font=("Segoe UI",8,"bold")).pack(fill=tk.X,padx=16,pady=(8,2))
        self.log_text = tk.Text(self.root,bg="#070c18",fg="#22ff88",font=("Consolas",9),
            height=6,relief="flat",wrap=tk.WORD,state=tk.DISABLED,padx=8,pady=4)
        self.log_text.pack(fill=tk.BOTH,expand=True,padx=14,pady=(2,10))

        self.root.bind("<F5>",lambda e:self.toggle())
        self.root.bind("<F6>",lambda e:self.pick_region())

        # เช็ค Tesseract
        try:
            pytesseract.get_tesseract_version()
            self.log("Tesseract OK!")
        except:
            self.log("ERROR: Tesseract ไม่เจอ!")
            self.log("ดาวน์โหลด: github.com/UB-Mannheim/tesseract/wiki")
            self.log("ติดตั้งที่ C:\\Program Files\\Tesseract-OCR")

    def _slider(self,p,label,var,lo,hi):
        f = tk.Frame(p,bg="#111a2b"); f.pack(fill=tk.X,padx=8,pady=2)
        tk.Label(f,text=label+":",bg="#111a2b",fg="#eee",font=("Segoe UI",9)).pack(side=tk.LEFT)
        vl = tk.Label(f,text=str(var.get()),bg="#111a2b",fg="#22d3ee",font=("Consolas",9,"bold"),width=5)
        vl.pack(side=tk.RIGHT)
        tk.Scale(f,from_=lo,to=hi,orient=tk.HORIZONTAL,variable=var,bg="#111a2b",fg="#111a2b",
            troughcolor="#0a0e17",highlightthickness=0,showvalue=False,length=140,sliderlength=14,
            command=lambda v,l=vl:l.config(text=str(int(float(v))))).pack(side=tk.RIGHT,padx=4)

    def pick_region(self):
        global region
        if self.running: self.log("Stop first!"); return
        self.root.iconify(); time.sleep(0.3)
        sel = tk.Toplevel(self.root)
        sel.overrideredirect(True); sel.attributes("-topmost",True); sel.attributes("-alpha",0.25)
        sw,sh = sel.winfo_screenwidth(), sel.winfo_screenheight()
        sel.geometry(f"{sw}x{sh}+0+0"); sel.configure(bg="black")
        c = tk.Canvas(sel,bg="black",highlightthickness=0,cursor="cross")
        c.pack(fill=tk.BOTH,expand=True)
        c.create_rectangle(sw//2-200,15,sw//2+200,75,fill="black",outline="#00ff88",width=2)
        c.create_text(sw//2,35,text="ลากครอบแถวตัวอักษร (ไม่รวม counter)",fill="#00ff88",font=("Segoe UI",14,"bold"))
        c.create_text(sw//2,60,text="ESC = ยกเลิก",fill="#aaa",font=("Segoe UI",10))
        pos = c.create_text(sw//2,95,text="",fill="#00ff88",font=("Consolas",12,"bold"))
        st = {"sx":0,"sy":0,"r":None}
        c.bind("<Motion>",lambda e:c.itemconfig(pos,text=f"X:{e.x} Y:{e.y}"))
        def _p(e):
            st["sx"],st["sy"]=e.x,e.y
            if st["r"]: c.delete(st["r"])
            st["r"]=c.create_rectangle(e.x,e.y,e.x,e.y,outline="#00ff88",width=3)
        def _d(e):
            c.coords(st["r"],st["sx"],st["sy"],e.x,e.y); c.delete("sz")
            c.create_text((st["sx"]+e.x)//2,min(st["sy"],e.y)-14,
                text=f"{abs(e.x-st['sx'])}x{abs(e.y-st['sy'])}",fill="#00ff88",font=("Consolas",13,"bold"),tags="sz")
        def _r(e):
            global region
            x1,y1=min(st["sx"],e.x),min(st["sy"],e.y); x2,y2=max(st["sx"],e.x),max(st["sy"],e.y)
            if (x2-x1)>10 and (y2-y1)>10:
                region={"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}; save_cfg()
                self.lbl_region.config(text=f"{x2-x1}x{y2-y1}",fg="#10b981")
                self.log(f"Region: {x2-x1}x{y2-y1} @ ({x1},{y1})")
            sel.destroy(); self.root.deiconify()
        c.bind("<ButtonPress-1>",_p); c.bind("<B1-Motion>",_d); c.bind("<ButtonRelease-1>",_r)
        sel.bind("<Escape>",lambda e:(sel.destroy(),self.root.deiconify()))
        sel.after(50,sel.focus_force)

    def test_read(self):
        if not region: self.log("เลือก region ก่อน!"); return
        frame = grab(region)
        if frame is None: self.log(f"จับจอไม่ได้! {grab_error}"); return
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        num = self.num_keys.get()
        keys = read_lane(gray, num)
        parts = [k.upper() if k else "?" for k in keys]
        self.log(f"Read: [{'] ['.join(parts)}]")

    def test_press(self):
        self.log("กด Q ใน 3 วิ... (คลิกเกมก่อน!)")
        def _do():
            time.sleep(3); press_key('q')
            self.root.after(0,self.log,"Press Q done!")
        threading.Thread(target=_do,daemon=True).start()

    def toggle(self):
        if self.running: self.running = False
        else:
            if not region: messagebox.showwarning("QTE","กด F6 เลือก Region!"); return
            self.running = True; save_cfg(num_keys=self.num_keys.get())
            self.btn.config(text="  STOP  ",bg="#ef4444")
            self.lbl_status.config(text="RUNNING",fg="#10b981")
            self.log("AFK: สลับไปเกมใน 3 วิ!")
            threading.Thread(target=self._run,daemon=True).start()

    def _run(self):
        num = self.num_keys.get()
        kd = self.key_delay.get() / 1000
        self.session_keys = 0
        self.session_start = time.time()

        test = grab(region)
        if test is None:
            self.root.after(0,self.log,f"จับจอไม่ได้! {grab_error}")
            self.running = False; self.root.after(0,self._reset_ui); return

        time.sleep(3)
        self.root.after(0,self.log,"เริ่ม!")
        last_seq = ""

        while self.running:
            try:
                frame = grab(region)
                if frame is None: time.sleep(0.1); continue
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                # อ่านทุก slot ด้วย OCR
                keys = read_lane(gray, num)
                valid = [k for k in keys if k is not None]

                # เจอครบ → กดรวด!
                if len(valid) == num:
                    seq = "".join(valid)
                    if seq != last_seq:
                        self.root.after(0,self.log,f"อ่านได้: {' '.join(k.upper() for k in valid)}")
                        self.root.after(0,self.log,f"กำลังกด {num} ปุ่ม...")

                        for idx, key in enumerate(valid):
                            if not self.running: break
                            ok = press_key(key)
                            self.session_keys += 1
                            self.root.after(0,self.log,
                                f"  กด {key.upper()} {'OK' if ok else 'FAIL!'}")
                            time.sleep(kd + random.uniform(0.01, 0.05))

                        self.root.after(0,self.lbl_count.config,{"text":str(self.session_keys)})
                        self.root.after(0,self.log,"กดครบ!")
                        last_seq = seq

                        # รอดึงปลา (~15 วิ)
                        self.root.after(0,self.log,"รอดึงปลา 15 วิ...")
                        time.sleep(15.0 + random.uniform(0.5, 2.0))

                        # โยนเบ็ดใหม่
                        if self.running:
                            self.root.after(0,self.log,"โยนเบ็ดใหม่")
                            press_key('e')  # *** แก้ปุ่มตามเซิร์ฟ ***
                            # รอมินิเกมขึ้น (~15 วิ)
                            self.root.after(0,self.log,"รอมินิเกม 15 วิ...")
                            time.sleep(15.0 + random.uniform(0.5, 2.0))
                            last_seq = ""
                    else:
                        time.sleep(0.05)
                else:
                    time.sleep(0.1)

            except Exception as e:
                self.root.after(0,self.log,f"Error: {e}")
                time.sleep(0.5)

        self.root.after(0,self._reset_ui)

    def _reset_ui(self):
        self.btn.config(text="  START (F5)  ",bg="#10b981")
        self.lbl_status.config(text="OFF",fg="#5a6a7e")
        self.log("Stopped")

    def _preview_loop(self):
        if region:
            try:
                img = ImageGrab.grab(bbox=(region["left"],region["top"],
                    region["left"]+region["width"],region["top"]+region["height"]))
                w,h = img.size
                if w > 390: img = img.resize((390,max(int(h*390/w),1)),Image.NEAREST)
                photo = ImageTk.PhotoImage(img)
                self.preview.config(image=photo,text=""); self.preview._photo = photo
            except: pass
        self.root.after(500,self._preview_loop)

    def log(self,msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END,f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        lines = int(self.log_text.index("end-1c").split(".")[0])
        if lines > 100: self.log_text.delete("1.0",f"{lines-100}.0")
        self.log_text.config(state=tk.DISABLED)

    def _quit(self): self.running=False; time.sleep(0.1); self.root.destroy()
    def run(self): self.root.mainloop()

if __name__ == "__main__": App().run()
