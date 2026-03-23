"""
AutoPet by Herlove v3.0
NO SETUP! ลากครอบเมนู 1 ครั้ง → Bot ทำทุกอย่างเอง
ตรวจสถานะด้วยสี ไม่ต้อง template
"""

import time, sys, os, json, threading, ctypes, random
import tkinter as tk
from tkinter import messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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
    import mss, cv2, numpy as np
    from PIL import Image, ImageTk, ImageGrab
except ImportError as e:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("AutoPet", f"pip install mss opencv-python numpy Pillow")
    sys.exit(1)

# ═══ Click ═══
def click_at(x, y):
    if sys.platform != 'win32': return False
    try:
        ctypes.windll.user32.SetCursorPos(int(x), int(y))
        time.sleep(0.08)
        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
        time.sleep(0.06)
        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)
        return True
    except: return False

# ═══ Capture ═══
_sct = None
def grab(r):
    global _sct
    if _sct is None:
        try: _sct = mss.mss()
        except: pass
    left,top,w,h = int(r["left"]),int(r["top"]),int(r["width"]),int(r["height"])
    if _sct:
        try:
            arr = np.array(_sct.grab({"left":left,"top":top,"width":w,"height":h}))
            if arr.size > 0: return cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)
        except: pass
    try:
        img = ImageGrab.grab(bbox=(left,top,left+w,top+h))
        if img: return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    except: pass
    return None

# ═══ State Detection (สี ไม่ต้อง template) ═══
def detect_state(card_img):
    """ตรวจสถานะจากสีของ card
    return: 'hungry' | 'ready' | 'empty' | 'growing'"""
    if card_img is None or card_img.size < 100:
        return "unknown"

    h, w = card_img.shape[:2]
    # ดูแค่ส่วนกลางของ card (40-80% ของความกว้าง/สูง)
    cy1, cy2 = int(h*0.15), int(h*0.65)
    cx1, cx2 = int(w*0.15), int(w*0.85)
    center = card_img[cy1:cy2, cx1:cx2]

    if center.size < 50:
        return "unknown"

    gray = cv2.cvtColor(center, cv2.COLOR_BGR2GRAY)
    hsv = cv2.cvtColor(center, cv2.COLOR_BGR2HSV)

    avg_brightness = np.mean(gray)
    bright_pixels = np.count_nonzero(gray > 200) / max(gray.size, 1)
    dark_pixels = np.count_nonzero(gray < 40) / max(gray.size, 1)

    # Hue channels
    h_ch = hsv[:,:,0]
    s_ch = hsv[:,:,1]

    # ═══ READY (checkmark ✓) ═══
    # checkmark = เยอะ pixel สว่างมาก (ขาว/เขียวอ่อน) > 15%
    if bright_pixels > 0.15:
        return "ready"

    # ═══ EMPTY ═══
    # ช่องว่าง = มืดมาก (>70% ของ pixel มืด)
    if dark_pixels > 0.70 and avg_brightness < 35:
        return "empty"

    # ═══ HUNGRY (food bowl) ═══
    # ชามอาหาร = มีสีอุ่น (gold/brown/silver) H:10-30, S:30-180
    warm_mask = (h_ch >= 8) & (h_ch <= 35) & (s_ch >= 25) & (s_ch <= 200)
    warm_ratio = np.count_nonzero(warm_mask) / max(h_ch.size, 1)
    # ชามมี metallic silver ด้วย: brightness สูงปานกลาง + saturation ต่ำ
    silver_mask = (gray > 120) & (gray < 220) & (s_ch < 40)
    silver_ratio = np.count_nonzero(silver_mask) / max(gray.size, 1)

    if warm_ratio > 0.08 or silver_ratio > 0.15:
        # อาจเป็นชาม ตรวจเพิ่ม: ชามไม่มี pattern วัว (ขาวดำ contrast สูง)
        contrast = np.std(gray)
        if contrast < 65:  # ชามมี contrast ต่ำกว่าวัว
            return "hungry"

    # ═══ GROWING ═══
    return "growing"

# ═══ Config ═══
CFG = os.path.join(SCRIPT_DIR, "pet_config.json")
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

# ═══ Grid: แบ่ง region เป็น slot อัตโนมัติ ═══
def get_card_positions(region, rows, cols):
    """คำนวณตำแหน่ง center + ปุ่ม GET + ปุ่ม + ของแต่ละ card"""
    left, top = region["left"], region["top"]
    w, h = region["width"], region["height"]
    cw = w / cols  # card width
    ch = h / rows  # card height

    cards = []
    for r in range(rows):
        for c in range(cols):
            cx = left + c * cw + cw / 2
            cy = top + r * ch + ch / 2
            # ปุ่ม + อยู่ล่างซ้าย
            px = left + c * cw + cw * 0.2
            py = top + r * ch + ch * 0.85
            # ปุ่ม GET อยู่ล่างขวา
            gx = left + c * cw + cw * 0.75
            gy = top + r * ch + ch * 0.85
            # capture area
            cap = {
                "left": int(left + c * cw + 2),
                "top": int(top + r * ch + 2),
                "width": int(cw - 4),
                "height": int(ch - 4)
            }
            cards.append({
                "center": (int(cx), int(cy)),
                "plus": (int(px), int(py)),
                "get": (int(gx), int(gy)),
                "capture": cap,
                "row": r, "col": c
            })
    return cards

# ═══ Theme ═══
BG="#020617"; CARD="#0f172a"; ACCENT="#f472b6"; GREEN="#4ade80"
RED="#f87171"; DIM="#334155"; WHITE="#e2e8f0"

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AutoPet by Herlove")
        self.root.geometry("460x620")
        self.root.configure(bg=BG)
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        self.running = False
        self.feed_count = 0
        self.harvest_count = 0
        self.add_count = 0
        self.check_sec = tk.IntVar(value=load_cfg().get("check_sec", 30))
        self.rows = tk.IntVar(value=load_cfg().get("rows", 2))
        self.cols = tk.IntVar(value=load_cfg().get("cols", 3))
        self.debug_img = None

        try:
            lp = os.path.join(SCRIPT_DIR, "mascot.png")
            if not os.path.exists(lp): lp = os.path.join(SCRIPT_DIR, "logo.png")
            if os.path.exists(lp):
                img = ImageTk.PhotoImage(Image.open(lp).resize((32,32),Image.LANCZOS))
                self.root.iconphoto(True, img); self._ico = img
        except: pass

        self._build()
        self._preview_loop()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        hdr = tk.Frame(self.root, bg=BG); hdr.pack(fill=tk.X, padx=16, pady=(10,0))
        try:
            lp = os.path.join(SCRIPT_DIR, "mascot.png")
            if not os.path.exists(lp): lp = os.path.join(SCRIPT_DIR, "logo.png")
            if os.path.exists(lp):
                li = Image.open(lp).resize((48,48),Image.LANCZOS)
                self._logo = ImageTk.PhotoImage(li)
                tk.Label(hdr, image=self._logo, bg=BG).pack(side=tk.LEFT, padx=(0,10))
        except: pass
        tf = tk.Frame(hdr, bg=BG); tf.pack(side=tk.LEFT)
        tk.Label(tf, text="AutoPet", bg=BG, fg=ACCENT, font=("Segoe UI",20,"bold")).pack(anchor="w")
        tk.Label(tf, text="NO SETUP! ลาก 1 ครั้ง → Bot ทำเอง", bg=BG, fg="#ec4899",
                 font=("Consolas",8,"bold")).pack(anchor="w")
        self.lbl_st = tk.Label(hdr, text="○ OFF", bg=BG, fg=DIM, font=("Consolas",11,"bold"))
        self.lbl_st.pack(side=tk.RIGHT)

        line = tk.Canvas(self.root, bg=BG, height=3, highlightthickness=0)
        line.pack(fill=tk.X, padx=16, pady=(6,0))
        for i in range(428):
            t=i/428
            line.create_line(i+16,0,i+16,3, fill=f"#{int(244-100*t):02x}{int(114+50*t):02x}{int(182-50*t):02x}")

        # Preview
        pf = tk.Frame(self.root, bg=CARD, highlightbackground="#ec4899", highlightthickness=1)
        pf.pack(fill=tk.X, padx=16, pady=(8,4))
        self.preview = tk.Label(pf, bg="#030a1a", text="F6 ลากครอบเมนูสัตว์ทั้งหมด",
                                 fg=DIM, font=("Consolas",9))
        self.preview.pack(fill=tk.X, ipady=30)

        # START
        self.btn = tk.Button(self.root, text="START (F5)", bg=GREEN, fg="#020617",
            font=("Segoe UI",13,"bold"), relief="flat", pady=5, cursor="hand2", command=self.toggle)
        self.btn.pack(fill=tk.X, padx=16, pady=(2,4))

        # Stats
        sf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        sf.pack(fill=tk.X, padx=16, pady=(0,4))
        si = tk.Frame(sf, bg=CARD); si.pack(fill=tk.X, padx=10, pady=6)
        for attr,label,color in [("lbl_fed","FED",GREEN),("lbl_harv","HARVEST","#38bdf8"),("lbl_add","ADDED",ACCENT)]:
            f = tk.Frame(si, bg=CARD); f.pack(side=tk.LEFT, expand=True)
            lbl = tk.Label(f, text="0", bg=CARD, fg=color, font=("Consolas",22,"bold"))
            lbl.pack()
            tk.Label(f, text=label, bg=CARD, fg=DIM, font=("Consolas",7,"bold")).pack()
            setattr(self, attr, lbl)

        # Region + Test
        rf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        rf.pack(fill=tk.X, padx=16, pady=(0,4))
        ri = tk.Frame(rf, bg=CARD); ri.pack(fill=tk.X, padx=8, pady=6)
        tk.Button(ri, text="Select Menu (F6)", bg="#0c1629", fg=ACCENT, font=("Segoe UI",9,"bold"),
            relief="flat", padx=8, pady=3, cursor="hand2", command=self.pick_region).pack(side=tk.LEFT, padx=2)
        tk.Button(ri, text="Test Scan", bg="#0c1629", fg=GREEN, font=("Segoe UI",9,"bold"),
            relief="flat", padx=8, pady=3, cursor="hand2", command=self.test_scan).pack(side=tk.LEFT, padx=2)
        tk.Button(ri, text="Test Click", bg="#0c1629", fg="#a78bfa", font=("Segoe UI",9,"bold"),
            relief="flat", padx=8, pady=3, cursor="hand2", command=self.test_click).pack(side=tk.LEFT, padx=2)
        self.lbl_reg = tk.Label(ri, text="not set", bg=CARD, fg="#f97316", font=("Consolas",9))
        self.lbl_reg.pack(side=tk.RIGHT, padx=6)

        # Settings
        stf = tk.Frame(self.root, bg=CARD, highlightbackground="#1e3a5f", highlightthickness=1)
        stf.pack(fill=tk.X, padx=16, pady=(0,4))
        self._sl(stf, "Rows", self.rows, 1, 3, ACCENT)
        self._sl(stf, "Columns", self.cols, 1, 4, ACCENT)
        self._sl(stf, "Check (sec)", self.check_sec, 10, 120, GREEN)

        # Log
        tk.Label(self.root, text="> LOG", bg=BG, fg=DIM, font=("Consolas",8,"bold")).pack(fill=tk.X, padx=18, pady=(4,1))
        self.log_box = tk.Text(self.root, bg="#030a1a", fg=GREEN, font=("Consolas",9),
            height=6, relief="flat", wrap=tk.WORD, state=tk.DISABLED, padx=10, pady=4)
        self.log_box.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0,8))

        self.root.bind("<F5>", lambda e: self.toggle())
        self.root.bind("<F6>", lambda e: self.pick_region())

        self.log("> AutoPet v3.0 (No Setup)")
        self.log(f"> Admin: {'YES' if is_admin() else 'NO!'}")
        if region: self.log(f"> Region: {region['width']}x{region['height']}")

    def _sl(self, p, label, var, lo, hi, color):
        f = tk.Frame(p, bg=CARD); f.pack(fill=tk.X, padx=10, pady=2)
        tk.Label(f, text=label, bg=CARD, fg=WHITE, font=("Segoe UI",9)).pack(side=tk.LEFT)
        vl = tk.Label(f, text=str(var.get()), bg=CARD, fg=color, font=("Consolas",10,"bold"), width=3)
        vl.pack(side=tk.RIGHT)
        tk.Scale(f, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var, bg=CARD, fg=CARD,
            troughcolor=BG, highlightthickness=0, showvalue=False, length=110, sliderlength=14,
            activebackground=color,
            command=lambda v,l=vl:l.config(text=str(int(float(v))))).pack(side=tk.RIGHT, padx=4)

    def pick_region(self):
        global region
        if self.running: self.log("> Stop first!"); return
        self.root.iconify(); time.sleep(0.3)
        sel = tk.Toplevel(self.root)
        sel.overrideredirect(True); sel.attributes("-topmost",True); sel.attributes("-alpha",0.2)
        sw,sh = sel.winfo_screenwidth(),sel.winfo_screenheight()
        sel.geometry(f"{sw}x{sh}+0+0"); sel.configure(bg="black")
        c = tk.Canvas(sel, bg="black", highlightthickness=0, cursor="cross")
        c.pack(fill=tk.BOTH, expand=True)
        c.create_rectangle(sw//2-200,15,sw//2+200,65, fill="black", outline=ACCENT, width=2)
        c.create_text(sw//2,33, text="ลากครอบ slot สัตว์ทั้งหมด (ไม่รวม header)", fill=ACCENT, font=("Segoe UI",13,"bold"))
        c.create_text(sw//2,53, text="ESC = ยกเลิก", fill="#64748b", font=("Segoe UI",10))
        pos = c.create_text(sw//2,85, text="", fill=ACCENT, font=("Consolas",12,"bold"))
        st = {"sx":0,"sy":0,"r":None}
        c.bind("<Motion>", lambda e: c.itemconfig(pos, text=f"X:{e.x} Y:{e.y}"))
        def _p(e):
            st["sx"],st["sy"]=e.x,e.y
            if st["r"]: c.delete(st["r"])
            st["r"]=c.create_rectangle(e.x,e.y,e.x,e.y, outline=ACCENT, width=3)
        def _d(e):
            c.coords(st["r"],st["sx"],st["sy"],e.x,e.y); c.delete("sz")
            c.create_text((st["sx"]+e.x)//2,min(st["sy"],e.y)-14,
                text=f"{abs(e.x-st['sx'])}x{abs(e.y-st['sy'])}", fill=ACCENT, font=("Consolas",13,"bold"), tags="sz")
        def _r(e):
            global region
            x1,y1=min(st["sx"],e.x),min(st["sy"],e.y); x2,y2=max(st["sx"],e.x),max(st["sy"],e.y)
            if (x2-x1)>20 and (y2-y1)>20:
                region={"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}; save_cfg()
                self.lbl_reg.config(text=f"{x2-x1}x{y2-y1}", fg=GREEN)
                self.log(f"> Region: {x2-x1}x{y2-y1} @ ({x1},{y1})")
            sel.destroy(); self.root.deiconify()
        c.bind("<ButtonPress-1>",_p); c.bind("<B1-Motion>",_d); c.bind("<ButtonRelease-1>",_r)
        sel.bind("<Escape>", lambda e:(sel.destroy(),self.root.deiconify()))
        sel.after(50, sel.focus_force)

    def test_scan(self):
        if not region: self.log("> Select menu first!"); return
        cards = get_card_positions(region, self.rows.get(), self.cols.get())
        self.log(f"> Scanning {len(cards)} slots...")
        for i, card in enumerate(cards):
            frame = grab(card["capture"])
            if frame is None: self.log(f">   Slot {i+1}: capture failed"); continue
            state = detect_state(frame)
            emoji = {"hungry":"HUNGRY","ready":"READY","empty":"EMPTY","growing":"OK","unknown":"??"}
            self.log(f">   Slot {i+1}: {emoji.get(state, state)}")
        # Show debug
        full = grab(region)
        if full is not None:
            self._draw_debug(full, cards)

    def test_click(self):
        if not region: self.log("> Select menu first!"); return
        cards = get_card_positions(region, self.rows.get(), self.cols.get())
        if not cards: return
        cx, cy = cards[0]["center"]
        self.log(f"> Click slot 1 center ({cx},{cy}) in 3s...")
        def _do():
            time.sleep(3)
            click_at(cx, cy)
            self.root.after(0, self.log, "> Clicked!")
        threading.Thread(target=_do, daemon=True).start()

    def toggle(self):
        if self.running: self.running = False
        else:
            if not region: messagebox.showwarning("AutoPet","F6 เลือก region ก่อน!"); return
            self.running = True
            save_cfg(rows=self.rows.get(), cols=self.cols.get(), check_sec=self.check_sec.get())
            self.btn.config(text="STOP (F5)", bg=RED)
            self.lbl_st.config(text="● ON", fg=GREEN)
            self.log("> สลับไปเกมใน 3 วิ!")
            threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        self.feed_count=0; self.harvest_count=0; self.add_count=0
        interval = self.check_sec.get()
        rows, cols = self.rows.get(), self.cols.get()
        time.sleep(3)
        self.root.after(0, self.log, "> Bot running!")

        while self.running:
            try:
                cards = get_card_positions(region, rows, cols)

                for i, card in enumerate(cards):
                    if not self.running: break
                    frame = grab(card["capture"])
                    if frame is None: continue
                    state = detect_state(frame)
                    cx, cy = card["center"]

                    if state == "hungry":
                        self.feed_count += 1
                        self.root.after(0, self.log, f"> Slot {i+1}: HUNGRY → Feed!")
                        click_at(cx, cy)
                        time.sleep(0.5 + random.uniform(0.1, 0.3))
                        click_at(cx, cy)
                        time.sleep(0.3)
                        self.root.after(0, self.lbl_fed.config, {"text":str(self.feed_count)})
                        if sys.platform=='win32':
                            try: import winsound; winsound.Beep(600,80)
                            except: pass

                    elif state == "ready":
                        self.harvest_count += 1
                        gx, gy = card["get"]
                        self.root.after(0, self.log, f"> Slot {i+1}: READY → Harvest!")
                        click_at(gx, gy)
                        time.sleep(1.0 + random.uniform(0.2, 0.5))
                        self.root.after(0, self.lbl_harv.config, {"text":str(self.harvest_count)})
                        # หลัง harvest → เพิ่มใหม่
                        self.add_count += 1
                        px, py = card["plus"]
                        self.root.after(0, self.log, f"> Slot {i+1}: → Add new!")
                        click_at(px, py)
                        time.sleep(0.8 + random.uniform(0.1, 0.3))
                        self.root.after(0, self.lbl_add.config, {"text":str(self.add_count)})

                    elif state == "empty":
                        self.add_count += 1
                        px, py = card["plus"]
                        self.root.after(0, self.log, f"> Slot {i+1}: EMPTY → Add!")
                        click_at(px, py)
                        time.sleep(0.8 + random.uniform(0.1, 0.3))
                        self.root.after(0, self.lbl_add.config, {"text":str(self.add_count)})

                    time.sleep(0.2 + random.uniform(0.05, 0.1))

                # Debug preview
                full = grab(region)
                if full is not None:
                    self.root.after(0, self._draw_debug, full, cards)

                self.root.after(0, self.log, f"> Next check in {interval}s")
                for _ in range(interval*10):
                    if not self.running: break
                    time.sleep(0.1)

            except Exception as e:
                self.root.after(0, self.log, f"> Error: {e}")
                time.sleep(1)

        self.root.after(0, self._reset)

    def _draw_debug(self, full, cards):
        d = full.copy()
        h,w = d.shape[:2]
        for i, card in enumerate(cards):
            cap = card["capture"]
            rx = cap["left"] - region["left"]
            ry = cap["top"] - region["top"]
            rw, rh = cap["width"], cap["height"]
            # Detect state
            slot_img = full[max(0,ry):ry+rh, max(0,rx):rx+rw]
            if slot_img.size < 50: continue
            state = detect_state(slot_img)
            colors = {"hungry":(0,0,255),"ready":(0,255,0),"empty":(100,100,100),"growing":(255,200,0),"unknown":(80,80,80)}
            clr = colors.get(state, (80,80,80))
            cv2.rectangle(d, (rx,ry), (rx+rw,ry+rh), clr, 2)
            labels = {"hungry":"HUNGRY","ready":"GET!","empty":"EMPTY","growing":"OK","unknown":"??"}
            cv2.putText(d, f"{i+1}:{labels.get(state,'')}", (rx+4,ry+16),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.4, clr, 1)
        try:
            rgb = cv2.cvtColor(d, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(rgb)
            iw,ih = img.size; nw=420; nh=max(int(ih*nw/iw),20)
            img = img.resize((nw,nh), Image.NEAREST)
            photo = ImageTk.PhotoImage(img)
            self.preview.config(image=photo, text=""); self.preview.image = photo
        except: pass

    def _reset(self):
        self.btn.config(text="START (F5)", bg=GREEN)
        self.lbl_st.config(text="○ OFF", fg=DIM)
        self.log("> Stopped")

    def _preview_loop(self):
        if region and not self.running:
            try:
                full = grab(region)
                if full is not None:
                    cards = get_card_positions(region, self.rows.get(), self.cols.get())
                    self._draw_debug(full, cards)
            except: pass
        self.root.after(800, self._preview_loop)

    def log(self, msg):
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_box.see(tk.END)
        lines = int(self.log_box.index("end-1c").split(".")[0])
        if lines > 80: self.log_box.delete("1.0",f"{lines-80}.0")
        self.log_box.config(state=tk.DISABLED)

    def _quit(self): self.running=False; time.sleep(0.1); self.root.destroy()
    def run(self): self.root.mainloop()

if __name__ == "__main__": App().run()
