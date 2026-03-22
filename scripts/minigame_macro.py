"""
QTE Macro - Final FiveM Edition
OCR ทั้งแถวครั้งเดียว + Visual Debug + All Key Methods
จอ 1920x1080 DX11 Borderless
"""

import time, sys, os, json, threading, ctypes, random, re
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
    print(f"pip install keyboard mss opencv-python numpy Pillow pytesseract")
    sys.exit(1)

for p in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
          r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]:
    if os.path.exists(p): pytesseract.pytesseract.tesseract_cmd = p; break

# ═══ Key Press (ทุกวิธีพร้อมกัน) ═══
VK = {'q':0x51,'w':0x57,'e':0x45,'r':0x52,'a':0x41,'s':0x53,'d':0x44,'f':0x46,'space':0x20}
SC = {'q':0x10,'w':0x11,'e':0x12,'r':0x13,'a':0x1E,'s':0x1F,'d':0x20,'f':0x21,'space':0x39}

def press_key(key):
    key = key.lower()
    vk, sc = VK.get(key), SC.get(key)
    if not vk: return False
    ok = False
    if sys.platform == 'win32':
        try:
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            lp_down = (sc or 0) << 16 | 1
            lp_up = (sc or 0) << 16 | 1 | (1 << 30) | (1 << 31)
            ctypes.windll.user32.PostMessageW(hwnd, 0x0100, vk, lp_down)
            time.sleep(0.06)
            ctypes.windll.user32.PostMessageW(hwnd, 0x0101, vk, lp_up)
            ok = True
        except: pass
        try:
            ctypes.windll.user32.keybd_event(vk, sc or 0, 0, 0)
            time.sleep(0.06)
            ctypes.windll.user32.keybd_event(vk, sc or 0, 2, 0)
            ok = True
        except: pass
    try: kb.press(key); time.sleep(0.06); kb.release(key); ok = True
    except: pass
    return ok

# ═══ Capture ═══
def grab(r):
    left,top,w,h = int(r["left"]),int(r["top"]),int(r["width"]),int(r["height"])
    try:
        with mss.mss() as s:
            arr = np.array(s.grab({"left":left,"top":top,"width":w,"height":h}))
            if arr.size > 0: return cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)
    except: pass
    try:
        img = ImageGrab.grab(bbox=(left,top,left+w,top+h))
        if img: return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    except: pass
    return None

# ═══ OCR - อ่านทั้งแถวครั้งเดียว ═══
VALID = set("qweasd")

def read_all(gray, num_keys):
    """อ่านทั้ง lane ครั้งเดียวด้วย OCR แล้วแยกทีละตัว
    return: list of chars เช่น ['d','s','w','a','w'] หรือ None ถ้าอ่านไม่ครบ"""
    if gray is None or gray.size < 100: return None

    # ขยาย 3 เท่า
    big = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)

    best_result = None

    # ลองหลาย threshold
    for method in range(5):
        try:
            if method == 0:
                _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            elif method == 1:
                _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                t = cv2.bitwise_not(t)
            elif method == 2:
                _, t = cv2.threshold(big, 180, 255, cv2.THRESH_BINARY)
            elif method == 3:
                _, t = cv2.threshold(big, 150, 255, cv2.THRESH_BINARY)
            elif method == 4:
                c = cv2.convertScaleAbs(big, alpha=2.0, beta=-50)
                _, t = cv2.threshold(c, 127, 255, cv2.THRESH_BINARY)

            # ให้ตัวอักษรเป็นสีดำ พื้นเป็นสีขาว (Tesseract ชอบแบบนี้)
            if np.count_nonzero(t) > t.size / 2:
                t = cv2.bitwise_not(t)

            # OCR ทั้งแถว (PSM 7 = single text line)
            config = '--psm 7 -c tessedit_char_whitelist=QWEASDqweasd'
            text = pytesseract.image_to_string(t, config=config).strip().lower()

            # กรองเอาเฉพาะตัวที่ valid
            chars = [c for c in text if c in VALID]

            if len(chars) == num_keys:
                return chars  # เจอครบ!

            # เก็บผลที่ได้มากที่สุดไว้
            if best_result is None or len(chars) > len(best_result):
                best_result = chars

        except: continue

    # ถ้าไม่เจอครบ ลอง PSM 6 (block of text)
    try:
        _, t = cv2.threshold(big, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        if np.count_nonzero(t) > t.size / 2: t = cv2.bitwise_not(t)
        config = '--psm 6 -c tessedit_char_whitelist=QWEASDqweasd'
        text = pytesseract.image_to_string(t, config=config).strip().lower()
        chars = [c for c in text if c in VALID]
        if len(chars) == num_keys: return chars
        if best_result is None or len(chars) > len(best_result):
            best_result = chars
    except: pass

    return best_result if best_result and len(best_result) == num_keys else None

def make_debug_image(frame, keys_found, num_keys):
    """สร้างภาพ debug มี box + ตัวอักษรที่อ่านได้"""
    debug = frame.copy()
    h, w = debug.shape[:2]
    sw = w // num_keys

    for i in range(num_keys):
        x1, x2 = i * sw, min((i+1) * sw, w)

        if keys_found and i < len(keys_found) and keys_found[i]:
            # เจอตัวอักษร → กรอบเขียว
            cv2.rectangle(debug, (x1+2, 2), (x2-2, h-2), (0, 255, 0), 2)
            cv2.putText(debug, keys_found[i].upper(), (x1+8, 25),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
        else:
            # ไม่เจอ → กรอบแดง
            cv2.rectangle(debug, (x1+2, 2), (x2-2, h-2), (0, 0, 255), 1)
            cv2.putText(debug, "?", (x1+8, 25),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)

    return debug

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
        self.root.geometry("500x600")
        self.root.configure(bg="#0a0e17")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.running = False
        self.session_keys = 0
        self.session_start = 0
        self.debug_img = None
        self.num_keys = tk.IntVar(value=load_cfg().get("num_keys", 5))
        self.key_delay = tk.IntVar(value=60)
        self._build()
        self._preview_loop()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        BG,CARD,DIM = "#0a0e17","#111a2b","#5a6a7e"

        tk.Label(self.root,text="QTE Macro",bg=BG,fg="#22d3ee",font=("Segoe UI",16,"bold")).pack(pady=(8,0))
        tk.Label(self.root,text="OCR + Visual Debug | 1920x1080 DX11",bg=BG,fg=DIM,font=("Segoe UI",8)).pack()

        # Preview ใหญ่ขึ้น
        self.preview = tk.Label(self.root, bg="#0c1220", text="F6 เลือก region",
                                 fg=DIM, font=("Segoe UI",9))
        self.preview.pack(fill=tk.X, padx=14, pady=(8,4), ipady=30)

        self.lbl_st = tk.Label(self.root,text="OFF",bg=BG,fg=DIM,font=("Consolas",10,"bold"))
        self.lbl_st.pack()

        self.btn = tk.Button(self.root,text="  START (F5)  ",bg="#10b981",fg="white",
            font=("Segoe UI",14,"bold"),relief="flat",pady=6,command=self.toggle)
        self.btn.pack(fill=tk.X,padx=14,pady=4)

        f1 = tk.Frame(self.root,bg=CARD); f1.pack(fill=tk.X,padx=14,pady=2)
        self.lbl_cnt = tk.Label(f1,text="0",bg=CARD,fg="#10b981",font=("Consolas",24,"bold"))
        self.lbl_cnt.pack(side=tk.LEFT,padx=10,pady=4)
        self.lbl_info = tk.Label(f1,text="",bg=CARD,fg=DIM,font=("Consolas",10))
        self.lbl_info.pack(side=tk.LEFT)

        f2 = tk.Frame(self.root,bg=CARD); f2.pack(fill=tk.X,padx=14,pady=2)
        tk.Button(f2,text="Select (F6)",bg="#1e3a5f",fg="#22d3ee",font=("Segoe UI",9,"bold"),
            relief="flat",padx=8,pady=3,command=self.pick_region).pack(side=tk.LEFT,padx=4,pady=4)
        tk.Button(f2,text="Test Read",bg="#1a3320",fg="#22ff88",font=("Segoe UI",9,"bold"),
            relief="flat",padx=8,pady=3,command=self.test_read).pack(side=tk.LEFT,padx=2,pady=4)
        tk.Button(f2,text="Test Press",bg="#3b1764",fg="#a78bfa",font=("Segoe UI",9,"bold"),
            relief="flat",padx=8,pady=3,command=self.test_press).pack(side=tk.LEFT,padx=2,pady=4)
        self.lbl_reg = tk.Label(f2,text="not set",bg=CARD,fg="#f97316",font=("Consolas",9))
        self.lbl_reg.pack(side=tk.RIGHT,padx=8)

        sf = tk.Frame(self.root,bg=CARD); sf.pack(fill=tk.X,padx=14,pady=2)
        self._sl(sf,"Keys/Lane",self.num_keys,2,10)
        self._sl(sf,"Key Delay ms",self.key_delay,30,200)

        self.log_box = tk.Text(self.root,bg="#070c18",fg="#22ff88",font=("Consolas",9),
            height=7,relief="flat",wrap=tk.WORD,state=tk.DISABLED,padx=8,pady=4)
        self.log_box.pack(fill=tk.BOTH,expand=True,padx=14,pady=(4,10))

        self.root.bind("<F5>",lambda e:self.toggle())
        self.root.bind("<F6>",lambda e:self.pick_region())

        try:
            v = pytesseract.get_tesseract_version()
            self.log(f"Tesseract {v} OK")
        except:
            self.log("ERROR: ติดตั้ง Tesseract ก่อน!")

    def _sl(self,p,label,var,lo,hi):
        f = tk.Frame(p,bg="#111a2b"); f.pack(fill=tk.X,padx=8,pady=1)
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
        sw,sh = sel.winfo_screenwidth(),sel.winfo_screenheight()
        sel.geometry(f"{sw}x{sh}+0+0"); sel.configure(bg="black")
        c = tk.Canvas(sel,bg="black",highlightthickness=0,cursor="cross")
        c.pack(fill=tk.BOTH,expand=True)
        c.create_rectangle(sw//2-220,15,sw//2+220,80,fill="black",outline="#00ff88",width=2)
        c.create_text(sw//2,33,text="ลากครอบแค่ตัวอักษร ไม่รวม counter/วงกลม",fill="#00ff88",font=("Segoe UI",13,"bold"))
        c.create_text(sw//2,55,text="ลากให้พอดีกรอบดำของตัวอักษร | ESC ยกเลิก",fill="#aaa",font=("Segoe UI",10))
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
        if frame is None: self.log("จับจอไม่ได้!"); return
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        num = self.num_keys.get()
        keys = read_all(gray, num)
        # สร้าง debug image
        self.debug_img = make_debug_image(frame, keys, num)
        self._show_debug()
        if keys:
            self.log(f"อ่านได้: {' '.join(k.upper() for k in keys)}")
        else:
            self.log("อ่านไม่ครบ! ลอง: ลากregionให้พอดีตัวอักษร")

    def test_press(self):
        self.log("กด D ใน 3 วิ (คลิกเกมก่อน!)")
        def _do():
            time.sleep(3)
            ok = press_key('d')
            self.root.after(0,self.log,f"Press D: {'OK' if ok else 'FAIL'}")
            time.sleep(0.3)
            ok2 = press_key('s')
            self.root.after(0,self.log,f"Press S: {'OK' if ok2 else 'FAIL'}")
        threading.Thread(target=_do,daemon=True).start()

    def toggle(self):
        if self.running: self.running = False
        else:
            if not region: messagebox.showwarning("QTE","กด F6 เลือก Region!"); return
            self.running = True; save_cfg(num_keys=self.num_keys.get())
            self.btn.config(text="  STOP  ",bg="#ef4444")
            self.lbl_st.config(text="RUNNING",fg="#10b981")
            self.log("สลับไปเกมใน 3 วิ!")
            threading.Thread(target=self._run,daemon=True).start()

    def _run(self):
        num = self.num_keys.get()
        kd = self.key_delay.get() / 1000
        self.session_keys = 0
        self.session_start = time.time()

        test = grab(region)
        if test is None:
            self.root.after(0,self.log,"จับจอไม่ได้!")
            self.running = False; self.root.after(0,self._reset); return

        time.sleep(3)
        self.root.after(0,self.log,"เริ่ม! สแกนไปเรื่อยๆ...")
        last_seq = ""

        while self.running:
            try:
                frame = grab(region)
                if frame is None: time.sleep(0.1); continue
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                # OCR ทั้งแถว
                keys = read_all(gray, num)

                # สร้าง debug image
                self.debug_img = make_debug_image(frame, keys, num)

                if keys and len(keys) == num:
                    seq = "".join(keys)
                    if seq != last_seq:
                        self.root.after(0,self.log,f"เจอ: {' '.join(k.upper() for k in keys)}")

                        for key in keys:
                            if not self.running: break
                            press_key(key)
                            self.session_keys += 1
                            time.sleep(kd + random.uniform(0.01, 0.04))

                        self.root.after(0,self.lbl_cnt.config,{"text":str(self.session_keys)})
                        self.root.after(0,self.lbl_info.config,
                            {"text":f"{self.session_keys/(time.time()-self.session_start):.1f}/s"})
                        last_seq = seq
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

    def _show_debug(self):
        """แสดง debug image ใน preview"""
        if self.debug_img is not None:
            try:
                rgb = cv2.cvtColor(self.debug_img, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(rgb)
                w,h = img.size
                # ขยายให้เต็ม preview (กว้าง ~470px)
                nw = 470
                nh = max(int(h * nw / w), 20)
                img = img.resize((nw, nh), Image.NEAREST)
                photo = ImageTk.PhotoImage(img)
                self.preview.config(image=photo, text="")
                self.preview._photo = photo
            except: pass

    def _preview_loop(self):
        if self.running and self.debug_img is not None:
            self._show_debug()
        elif region:
            try:
                frame = grab(region)
                if frame is not None:
                    num = self.num_keys.get()
                    debug = make_debug_image(frame, None, num)
                    rgb = cv2.cvtColor(debug, cv2.COLOR_BGR2RGB)
                    img = Image.fromarray(rgb)
                    w,h = img.size
                    nw = 470; nh = max(int(h*nw/w), 20)
                    img = img.resize((nw,nh), Image.NEAREST)
                    photo = ImageTk.PhotoImage(img)
                    self.preview.config(image=photo, text="")
                    self.preview._photo = photo
            except: pass
        self.root.after(200, self._preview_loop)

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
