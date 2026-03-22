"""
QTE Macro - FiveM Edition
OCR + PostMessage (ส่ง key ตรงเข้าหน้าต่างเกม)
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

for p in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
          r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]:
    if os.path.exists(p): pytesseract.pytesseract.tesseract_cmd = p; break

# ═══════════════════════════════════
# Key Press - ลองทุกวิธีที่มี
# ═══════════════════════════════════
VK = {'q':0x51,'w':0x57,'e':0x45,'r':0x52,'a':0x41,'s':0x53,'d':0x44,'f':0x46,'space':0x20}
SCAN = {'q':0x10,'w':0x11,'e':0x12,'r':0x13,'a':0x1E,'s':0x1F,'d':0x20,'f':0x21,'space':0x39}
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101

press_method = "all"  # จะใช้ทุกวิธีพร้อมกัน

def press_key(key):
    """กดปุ่ม - ยิงทุกวิธีพร้อมกัน ตัวไหนเข้าก็เข้า"""
    key = key.lower()
    vk = VK.get(key)
    sc = SCAN.get(key)
    if not vk: return False

    ok = False

    if sys.platform == 'win32':
        # วิธี 1: PostMessage ตรงเข้าหน้าต่างที่ active
        try:
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            if hwnd:
                lparam_down = (sc or 0) << 16 | 1
                lparam_up = (sc or 0) << 16 | 1 | (1 << 30) | (1 << 31)
                ctypes.windll.user32.PostMessageW(hwnd, WM_KEYDOWN, vk, lparam_down)
                time.sleep(0.05)
                ctypes.windll.user32.PostMessageW(hwnd, WM_KEYUP, vk, lparam_up)
                ok = True
        except: pass

        # วิธี 2: keybd_event
        try:
            ctypes.windll.user32.keybd_event(vk, sc or 0, 0, 0)
            time.sleep(0.05)
            ctypes.windll.user32.keybd_event(vk, sc or 0, 0x0002, 0)
            ok = True
        except: pass

        # วิธี 3: SendInput scan code
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
            # กดด้วยทั้ง VK + scan code
            ii = IU(); ii.ki = KI(vk, sc or 0, 0, 0, ctypes.pointer(ex))
            x = INP(ctypes.c_ulong(1), ii)
            ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))
            time.sleep(0.05)
            ii2 = IU(); ii2.ki = KI(vk, sc or 0, 0x0002, 0, ctypes.pointer(ex))
            x2 = INP(ctypes.c_ulong(1), ii2)
            ctypes.windll.user32.SendInput(1, ctypes.pointer(x2), ctypes.sizeof(x2))
            ok = True
        except: pass

    # วิธี 4: keyboard library
    try:
        kb.press(key); time.sleep(0.05); kb.release(key)
        ok = True
    except: pass

    return ok

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

# ═══ OCR ═══
VALID = set("qweasd")
OCR_CFG = '--psm 10 -c tessedit_char_whitelist=QWEASDqweasd'

def read_slot(gray_slot):
    if gray_slot is None or gray_slot.size < 20: return None
    big = cv2.resize(gray_slot, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
    candidates = []
    for method in range(4):
        try:
            if method == 0:
                _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
            elif method == 1:
                _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
                t = cv2.bitwise_not(t)
            elif method == 2:
                c = cv2.convertScaleAbs(big, alpha=2.5, beta=-80)
                _, t = cv2.threshold(c, 150, 255, cv2.THRESH_BINARY)
            else:
                t = cv2.adaptiveThreshold(big, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                           cv2.THRESH_BINARY, 15, 5)
            text = pytesseract.image_to_string(t, config=OCR_CFG).strip().lower()
            for ch in text:
                if ch in VALID: candidates.append(ch); break
        except: continue
    if not candidates: return None
    from collections import Counter
    return Counter(candidates).most_common(1)[0][0]

def read_lane(gray, num):
    h,w = gray.shape[:2]; sw = w // num
    results = []
    for i in range(num):
        pad = max(sw//10, 2)
        slot = gray[:, max(0,i*sw-pad):min(w,(i+1)*sw+pad)]
        results.append(read_slot(slot))
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
        self.root.title("QTE Macro")
        self.root.geometry("420x480")
        self.root.configure(bg="#0a0e17")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.running = False
        self.session_keys = 0
        self.session_start = 0
        self.num_keys = tk.IntVar(value=load_cfg().get("num_keys", 5))
        self.key_delay = tk.IntVar(value=60)
        self._build()
        self._preview_loop()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        BG,CARD,DIM = "#0a0e17","#111a2b","#5a6a7e"
        tk.Label(self.root,text="QTE Macro",bg=BG,fg="#22d3ee",font=("Segoe UI",16,"bold")).pack(pady=(10,0))
        tk.Label(self.root,text="OCR + FiveM Key Inject",bg=BG,fg=DIM,font=("Segoe UI",8)).pack()
        self.preview = tk.Label(self.root,bg="#0c1220",text="F6 เลือก region",fg=DIM,font=("Segoe UI",8),height=3)
        self.preview.pack(fill=tk.X,padx=14,pady=(8,4))
        self.lbl_st = tk.Label(self.root,text="OFF",bg=BG,fg=DIM,font=("Consolas",10,"bold"))
        self.lbl_st.pack()
        self.btn = tk.Button(self.root,text="  START (F5)  ",bg="#10b981",fg="white",
            font=("Segoe UI",14,"bold"),relief="flat",pady=6,command=self.toggle)
        self.btn.pack(fill=tk.X,padx=14,pady=6)
        f1 = tk.Frame(self.root,bg=CARD); f1.pack(fill=tk.X,padx=14,pady=2)
        self.lbl_cnt = tk.Label(f1,text="0",bg=CARD,fg="#10b981",font=("Consolas",24,"bold"))
        self.lbl_cnt.pack(side=tk.LEFT,padx=10,pady=6)
        self.lbl_info = tk.Label(f1,text="",bg=CARD,fg=DIM,font=("Consolas",10))
        self.lbl_info.pack(side=tk.LEFT)

        tk.Label(self.root,text="REGION + TEST",bg=BG,fg=DIM,font=("Segoe UI",8,"bold")).pack(fill=tk.X,padx=16,pady=(8,2))
        f2 = tk.Frame(self.root,bg=CARD); f2.pack(fill=tk.X,padx=14,pady=2)
        tk.Button(f2,text="Select (F6)",bg="#1e3a5f",fg="#22d3ee",font=("Segoe UI",9,"bold"),
            relief="flat",padx=8,pady=4,command=self.pick_region).pack(side=tk.LEFT,padx=4,pady=6)
        tk.Button(f2,text="Test Read",bg="#1a3320",fg="#22ff88",font=("Segoe UI",9,"bold"),
            relief="flat",padx=8,pady=4,command=self.test_read).pack(side=tk.LEFT,padx=2,pady=6)
        tk.Button(f2,text="Test Press",bg="#3b1764",fg="#a78bfa",font=("Segoe UI",9,"bold"),
            relief="flat",padx=8,pady=4,command=self.test_press).pack(side=tk.LEFT,padx=2,pady=6)
        self.lbl_reg = tk.Label(f2,text="not set",bg=CARD,fg="#f97316",font=("Consolas",9))
        self.lbl_reg.pack(side=tk.RIGHT,padx=8)

        tk.Label(self.root,text="SETTINGS",bg=BG,fg=DIM,font=("Segoe UI",8,"bold")).pack(fill=tk.X,padx=16,pady=(8,2))
        sf = tk.Frame(self.root,bg=CARD); sf.pack(fill=tk.X,padx=14,pady=2)
        self._sl(sf,"Keys/Lane",self.num_keys,2,10)
        self._sl(sf,"Key Delay ms",self.key_delay,30,200)

        tk.Label(self.root,text="LOG",bg=BG,fg=DIM,font=("Segoe UI",8,"bold")).pack(fill=tk.X,padx=16,pady=(8,2))
        self.log_box = tk.Text(self.root,bg="#070c18",fg="#22ff88",font=("Consolas",9),
            height=6,relief="flat",wrap=tk.WORD,state=tk.DISABLED,padx=8,pady=4)
        self.log_box.pack(fill=tk.BOTH,expand=True,padx=14,pady=(2,10))
        self.root.bind("<F5>",lambda e:self.toggle())
        self.root.bind("<F6>",lambda e:self.pick_region())

        try:
            pytesseract.get_tesseract_version()
            self.log("Tesseract OK")
        except:
            self.log("ERROR: ติดตั้ง Tesseract ก่อน!")

    def _sl(self,p,label,var,lo,hi):
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
                self.lbl_reg.config(text=f"{x2-x1}x{y2-y1}",fg="#10b981")
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
        keys = read_lane(gray, self.num_keys.get())
        self.log(f"Read: [{'] ['.join(k.upper() if k else '?' for k in keys)}]")

    def test_press(self):
        self.log("กด D ใน 3 วิ (คลิกเกมก่อน!)")
        def _do():
            time.sleep(3)
            ok = press_key('d')
            self.root.after(0,self.log,f"Press D: {'OK' if ok else 'FAIL'}")
        threading.Thread(target=_do,daemon=True).start()

    def toggle(self):
        if self.running:
            self.running = False
        else:
            if not region: messagebox.showwarning("QTE","กด F6 เลือก Region!"); return
            self.running = True; save_cfg(num_keys=self.num_keys.get())
            self.btn.config(text="  STOP  ",bg="#ef4444")
            self.lbl_st.config(text="RUNNING",fg="#10b981")
            self.log("สลับไปเกมใน 3 วิ!")
            threading.Thread(target=self._run,daemon=True).start()

    def _run(self):
        """Loop ง่ายสุด: สแกน → เจอ → กด → วน"""
        num = self.num_keys.get()
        kd = self.key_delay.get() / 1000
        self.session_keys = 0
        self.session_start = time.time()

        test = grab(region)
        if test is None:
            self.root.after(0,self.log,f"จับจอไม่ได้! {grab_error}")
            self.running = False; self.root.after(0,self._reset); return

        time.sleep(3)
        self.root.after(0,self.log,"เริ่ม! สแกนไปเรื่อยๆ...")
        last_seq = ""

        while self.running:
            try:
                frame = grab(region)
                if frame is None: time.sleep(0.1); continue
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                keys = read_lane(gray, num)
                valid = [k for k in keys if k is not None]

                if len(valid) == num:
                    seq = "".join(valid)
                    if seq != last_seq:
                        self.root.after(0,self.log,f"เจอ: {' '.join(k.upper() for k in valid)}")

                        for key in valid:
                            if not self.running: break
                            press_key(key)
                            self.session_keys += 1
                            self.root.after(0,self.log,f"  กด {key.upper()}")
                            time.sleep(kd + random.uniform(0.01, 0.05))

                        self.root.after(0,self.lbl_cnt.config,{"text":str(self.session_keys)})
                        last_seq = seq
                        # รอสั้นๆ ป้องกันกดซ้ำ
                        time.sleep(1.0 + random.uniform(0.1, 0.3))
                    else:
                        time.sleep(0.05)
                else:
                    time.sleep(0.1)

            except Exception as e:
                self.root.after(0,self.log,f"Error: {e}")
                time.sleep(0.5)

        self.root.after(0,self._reset)

    def _reset(self):
        self.btn.config(text="  START (F5)  ",bg="#10b981")
        self.lbl_st.config(text="OFF",fg="#5a6a7e")
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
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END,f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_box.see(tk.END)
        lines = int(self.log_box.index("end-1c").split(".")[0])
        if lines > 100: self.log_box.delete("1.0",f"{lines-100}.0")
        self.log_box.config(state=tk.DISABLED)

    def _quit(self): self.running=False; time.sleep(0.1); self.root.destroy()
    def run(self): self.root.mainloop()

if __name__ == "__main__": App().run()
