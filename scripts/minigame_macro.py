"""
QTE Minigame Macro v4 - Lane Mode
ตัวอักษรขึ้นทีเดียวหลายตัว เช่น A S D E W แล้วกดทีละตัวจนครบ
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
for lib, pkg in [("keyboard","keyboard"),("mss","mss"),
                  ("cv2","opencv-python"),("numpy","numpy"),("PIL","Pillow")]:
    try: __import__(lib)
    except ImportError: MISSING.append(pkg)

if MISSING:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing", f"pip install {' '.join(MISSING)}")
    sys.exit(1)

import keyboard as kb
import mss, mss.tools
import cv2, numpy as np
from PIL import Image, ImageTk, ImageGrab

# ============================================================
# Screen Capture
# ============================================================
class Capture:
    def __init__(self):
        self.method = None; self.sct = None; self._detect()

    def _detect(self):
        try:
            s = mss.mss(); t = s.grab(s.monitors[0])
            if t.size[0] > 0: self.method = "mss"; self.sct = s; return
        except: pass
        try:
            t = ImageGrab.grab(bbox=(0,0,50,50))
            if t: self.method = "pil"; return
        except: pass

    def grab(self, region):
        x,y,w,h = region["left"],region["top"],region["width"],region["height"]
        try:
            if self.method == "mss":
                s = self.sct.grab(region)
                return cv2.cvtColor(np.array(s), cv2.COLOR_BGRA2BGR)
            elif self.method == "pil":
                return cv2.cvtColor(np.array(ImageGrab.grab(bbox=(x,y,x+w,y+h))), cv2.COLOR_RGB2BGR)
        except: self._detect()
        return None

    def grab_pil(self, region):
        x,y,w,h = region["left"],region["top"],region["width"],region["height"]
        try:
            if self.method == "mss":
                s = self.sct.grab(region)
                return Image.frombytes("RGB", s.size, s.bgra, "raw", "BGRX")
            elif self.method == "pil":
                return ImageGrab.grab(bbox=(x,y,x+w,y+h))
        except: pass; return None

    def close(self):
        if self.sct:
            try: self.sct.close()
            except: pass


# ============================================================
# Template Engine - Pre-computed scales
# ============================================================
class Templates:
    def __init__(self, d):
        self.dir = d; self.raw = {}; self.scaled = {}
        os.makedirs(d, exist_ok=True); self.load()

    def load(self):
        self.raw.clear(); self.scaled.clear()
        if not os.path.isdir(self.dir): return
        for f in os.listdir(self.dir):
            if not f.lower().endswith(".png"): continue
            key = f.rsplit(".",1)[0].split("_")[0].lower()
            if len(key) != 1: continue
            img = cv2.imread(os.path.join(self.dir, f), cv2.IMREAD_GRAYSCALE)
            if img is not None:
                self.raw.setdefault(key, []).append(img)
        self._build_scales()

    def _build_scales(self):
        """Pre-compute ทุก scale ล่วงหน้า"""
        self.scaled.clear()
        for key, imgs in self.raw.items():
            self.scaled[key] = []
            for img in imgs:
                for sc in [1.0, 0.95, 1.05, 0.9, 1.1]:
                    tw, th = int(img.shape[1]*sc), int(img.shape[0]*sc)
                    if tw < 3 or th < 3: continue
                    self.scaled[key].append(cv2.resize(img, (tw, th)))

    def match_one(self, gray_roi, threshold=0.75):
        """Match 1 ตัวอักษรใน ROI เล็กๆ"""
        best_key, best_val = None, 0
        gh, gw = gray_roi.shape[:2]
        for key, tmps in self.scaled.items():
            for t in tmps:
                if t.shape[0] > gh or t.shape[1] > gw: continue
                res = cv2.matchTemplate(gray_roi, t, cv2.TM_CCOEFF_NORMED)
                _, mx, _, _ = cv2.minMaxLoc(res)
                if mx > threshold and mx > best_val:
                    best_key, best_val = key, mx
                    if mx > 0.93: return best_key, best_val
        return best_key, best_val

    def match_lane(self, gray_full, num_keys, threshold=0.75):
        """Match ทั้งแถว - แบ่ง lane เป็น N ช่อง แล้ว match ทีละช่อง
        return: list of (key, confidence) เช่น [('a',0.9), ('s',0.85), ('d',0.92)]"""
        h, w = gray_full.shape[:2]
        slot_w = w // num_keys
        results = []

        for i in range(num_keys):
            x1 = i * slot_w
            x2 = min(x1 + slot_w, w)
            # เพิ่ม padding เล็กน้อย เผื่อตัวอักษรไม่ตรงกลาง
            pad = slot_w // 6
            rx1 = max(0, x1 - pad)
            rx2 = min(w, x2 + pad)
            roi = gray_full[:, rx1:rx2]

            if roi.shape[1] < 5: continue
            key, conf = self.match_one(roi, threshold)
            results.append((key, conf))

        return results

    def save(self, key, frame_bgr):
        existing = [f for f in os.listdir(self.dir) if f.lower().startswith(key) and f.endswith(".png")]
        fname = f"{key}.png" if not existing else f"{key}_{len(existing)+1}.png"
        path = os.path.join(self.dir, fname)
        cv2.imwrite(path, frame_bgr)
        gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY)
        self.raw.setdefault(key, []).append(gray)
        self._build_scales()
        return fname

    @property
    def count(self): return sum(len(v) for v in self.raw.values())
    @property
    def keys_list(self): return sorted(self.raw.keys())


# ============================================================
# Lane Detector - หัวใจของ v4
# ============================================================
class LaneDetector:
    """
    จับทั้ง lane → อ่านทุกตัว → กดทีละตัว → รอ lane ใหม่

    Flow:
    1. Scan lane → ได้ sequence เช่น [a, s, d, e, w]
    2. กด a → รอ delay → กด s → รอ delay → ...
    3. กดครบ → รอ lane เปลี่ยน (sequence ใหม่)
    4. ถ้า lane เหมือนเดิม → ไม่กดซ้ำ
    """
    def __init__(self):
        self.last_sequence = []    # sequence ที่กดไปแล้ว
        self.current_idx = 0       # กดถึงตัวที่เท่าไหร่แล้ว
        self.pressing = False      # กำลังกด sequence อยู่
        self.last_lane_hash = None

    def get_lane_hash(self, sequence):
        """สร้าง hash จาก sequence เพื่อเทียบว่าเป็น lane เดิมหรือใหม่"""
        return "".join(k or "?" for k, c in sequence)

    def is_new_lane(self, sequence):
        """เช็คว่าเป็น lane ใหม่ (ต่างจากที่กดไปแล้ว)"""
        h = self.get_lane_hash(sequence)
        if h != self.last_lane_hash:
            return True
        return False

    def start_sequence(self, sequence):
        """เริ่มกด sequence ใหม่"""
        self.last_sequence = sequence
        self.last_lane_hash = self.get_lane_hash(sequence)
        self.current_idx = 0
        self.pressing = True

    def next_key(self):
        """ดึง key ถัดไปที่ต้องกด"""
        if self.current_idx < len(self.last_sequence):
            key, conf = self.last_sequence[self.current_idx]
            self.current_idx += 1
            if self.current_idx >= len(self.last_sequence):
                self.pressing = False
            return key, conf
        self.pressing = False
        return None, 0

    def reset(self):
        self.last_sequence = []
        self.current_idx = 0
        self.pressing = False
        self.last_lane_hash = None


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

class Btn(tk.Canvas):
    def __init__(self, parent, text="", color="#00d4ff", w=200, h=44, fs=12, cmd=None, **kw):
        super().__init__(parent, width=w, height=h, bg=C["bg"], highlightthickness=0, **kw)
        self._w,self._h,self.color,self.text,self.cmd,self.fs = w,h,color,text,cmd,fs
        self._hover = False
        self.bind("<Enter>", lambda e: self._set(True))
        self.bind("<Leave>", lambda e: self._set(False))
        self.bind("<Button-1>", lambda e: self.cmd() if self.cmd else None)
        self._draw()
    def _blend(self, c1, c2, t):
        r1,g1,b1=[int(c1[i:i+2],16) for i in (1,3,5)]
        r2,g2,b2=[int(c2[i:i+2],16) for i in (1,3,5)]
        return f"#{int(r1+(r2-r1)*t):02x}{int(g1+(g2-g1)*t):02x}{int(b1+(b2-b1)*t):02x}"
    def _draw(self):
        self.delete("all"); w,h,r=self._w,self._h,10; a=0.4 if self._hover else 0.0
        fill=self._blend(C["card"],self.color,0.15+a*0.15)
        brd=self._blend(self.color,"#ffffff",a*0.3)
        pts=[2+r,2,w-2-r,2,w-2,2,w-2,2+r,w-2,h-2-r,w-2,h-2,w-2-r,h-2,2+r,h-2,2,h-2,2,h-2-r,2,2+r,2,2]
        self.create_polygon(pts,smooth=True,fill=fill,outline=brd,width=2)
        tc=self._blend(self.color,"#ffffff",a*0.5)
        self.create_text(w//2,h//2,text=self.text,fill=tc,font=("Segoe UI",self.fs,"bold"))
    def _set(self,h): self._hover=h; self._draw()
    def set_text(self,t): self.text=t; self._draw()
    def set_color(self,c): self.color=c; self._draw()

class RegionSelector(tk.Toplevel):
    def __init__(self, master, callback):
        super().__init__(master); self.callback=callback; self.sx=self.sy=0; self.rect=None
        self.overrideredirect(True); self.attributes("-topmost",True)
        try: self.attributes("-alpha",0.25)
        except: pass
        sw,sh=self.winfo_screenwidth(),self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0"); self.configure(bg="black")
        self.c=tk.Canvas(self,bg="black",highlightthickness=0,cursor="cross")
        self.c.pack(fill=tk.BOTH,expand=True)
        self.c.create_text(sw//2,36,text="DRAG TO SELECT",fill="#00ffaa",font=("Segoe UI",22,"bold"))
        self.c.create_text(sw//2,68,text="ลากเลือก ทั้งแถว ที่ตัวอักษรขึ้น  |  ESC = ยกเลิก",fill="#ccc",font=("Segoe UI",12))
        self.c.bind("<ButtonPress-1>",self._p); self.c.bind("<B1-Motion>",self._d)
        self.c.bind("<ButtonRelease-1>",self._r); self.bind("<Escape>",lambda e:self.destroy())
        self.after(100,self.focus_force)
    def _p(self,e):
        self.sx,self.sy=e.x,e.y
        if self.rect: self.c.delete(self.rect)
        self.rect=self.c.create_rectangle(e.x,e.y,e.x,e.y,outline="#00ffaa",width=2,dash=(6,3))
    def _d(self,e):
        self.c.coords(self.rect,self.sx,self.sy,e.x,e.y); self.c.delete("sz")
        self.c.create_text((self.sx+e.x)//2,min(self.sy,e.y)-18,
            text=f"{abs(e.x-self.sx)} x {abs(e.y-self.sy)}",fill="#00ffaa",font=("Consolas",13,"bold"),tags="sz")
    def _r(self,e):
        x1,y1=min(self.sx,e.x),min(self.sy,e.y); x2,y2=max(self.sx,e.x),max(self.sy,e.y)
        if (x2-x1)>5 and (y2-y1)>5:
            self.callback({"left":x1,"top":y1,"width":x2-x1,"height":y2-y1})
        self.destroy()


# ============================================================
# MAIN APP
# ============================================================
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QTE Macro v4 - Lane Mode")
        self.root.geometry("560x800")
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)

        self.cfg = self._load_cfg()
        self.running = False
        self.region = self.cfg.get("region")
        self.session_keys = 0
        self.session_start = None
        self.key_history = []

        self.cap = Capture()
        self.tmpl = Templates(os.path.join(get_base_path(), "templates"))
        self.lane = LaneDetector()

        self.v_num_keys = tk.IntVar(value=self.cfg.get("num_keys", 6))
        self.v_scan = tk.IntVar(value=self.cfg.get("scan_ms", 30))
        self.v_thresh = tk.IntVar(value=int(self.cfg.get("match_thresh", 0.75)*100))
        self.v_key_delay = tk.IntVar(value=self.cfg.get("key_delay_ms", 50))
        self.v_lane_delay = tk.IntVar(value=self.cfg.get("lane_delay_ms", 200))
        self.v_ontop = tk.BooleanVar(value=self.cfg.get("ontop", True))
        self.v_sound = tk.BooleanVar(value=self.cfg.get("sound", True))

        self._build()
        self.root.attributes("-topmost", self.v_ontop.get())
        self._update_region()
        self._update_tmpl()
        self._preview_loop()
        self._timer_loop()

        self.root.bind_all("<F5>", lambda e: self.toggle())
        self.root.bind_all("<F6>", lambda e: self.pick_region())
        self.root.bind_all("<Escape>", lambda e: self._stop())
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _cfg_path(self): return os.path.join(get_base_path(), "macro_config.json")
    def _load_cfg(self):
        try:
            with open(self._cfg_path(),"r",encoding="utf-8") as f: return json.load(f)
        except: return {}
    def _save_cfg(self):
        self.cfg.update({
            "num_keys":self.v_num_keys.get(),"scan_ms":self.v_scan.get(),
            "match_thresh":self.v_thresh.get()/100,"key_delay_ms":self.v_key_delay.get(),
            "lane_delay_ms":self.v_lane_delay.get(),"region":self.region,
            "ontop":self.v_ontop.get(),"sound":self.v_sound.get(),
        })
        try:
            with open(self._cfg_path(),"w",encoding="utf-8") as f:
                json.dump(self.cfg,f,indent=2,ensure_ascii=False)
        except: pass

    # ══════════════════════════════════════════
    # UI
    # ══════════════════════════════════════════
    def _build(self):
        # Header
        hdr = tk.Canvas(self.root, bg=C["bg"], height=70, highlightthickness=0)
        hdr.pack(fill=tk.X)
        for i in range(560):
            t=i/560; r,g,b=int(168*t),int(212+(85-212)*t),int(255+(247-255)*t)
            hdr.create_line(i,0,i,3,fill=f"#{r:02x}{g:02x}{b:02x}")
        hdr.create_text(24,28,text="QTE MACRO v4",anchor="w",fill=C["blue"],font=("Segoe UI",22,"bold"))
        hdr.create_text(24,50,text="Lane Mode - กดทีละตัวจนครบแถว",anchor="w",fill=C["dim"],font=("Segoe UI",9))
        cap_ok = self.cap.method is not None
        self.lbl_status=hdr.create_text(500,28,text="OFF",anchor="w",fill=C["dim"],font=("Segoe UI",10,"bold"))
        self.lbl_timer=hdr.create_text(536,50,text="00:00",anchor="e",fill=C["dim2"],font=("Consolas",9))
        self.hdr = hdr

        main = tk.Frame(self.root, bg=C["bg"])
        main.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0,8))

        # START
        self.btn_start = Btn(main,text="START (F5)",color=C["green"],w=528,h=52,fs=17,cmd=self.toggle)
        self.btn_start.pack(pady=(4,8))

        # ── Stats ──
        sf = tk.Frame(main,bg=C["card"],highlightbackground=C["border"],highlightthickness=1)
        sf.pack(fill=tk.X,pady=(0,8))
        si = tk.Frame(sf,bg=C["card"]); si.pack(fill=tk.X,padx=14,pady=10)

        left=tk.Frame(si,bg=C["card"]); left.pack(side=tk.LEFT)
        self.lbl_count=tk.Label(left,text="0",bg=C["card"],fg=C["green"],font=("Consolas",32,"bold"))
        self.lbl_count.pack(anchor="w")
        tk.Label(left,text="KEYS",bg=C["card"],fg=C["dim"],font=("Segoe UI",8,"bold")).pack(anchor="w")

        mid=tk.Frame(si,bg=C["card"]); mid.pack(side=tk.LEFT,padx=16)
        self.lbl_lane_status=tk.Label(mid,text="Idle",bg=C["card"],fg=C["dim"],font=("Segoe UI",11,"bold"))
        self.lbl_lane_status.pack(anchor="w")
        self.lbl_speed=tk.Label(mid,text="0.0 /sec",bg=C["card"],fg=C["blue"],font=("Consolas",10))
        self.lbl_speed.pack(anchor="w")
        self.lbl_debug=tk.Label(mid,text="",bg=C["card"],fg=C["dim"],font=("Consolas",9))
        self.lbl_debug.pack(anchor="w")

        # Lane sequence display
        self.lbl_lane=tk.Label(si,text="",bg=C["card"],fg=C["cyan"],
                                font=("Consolas",16,"bold"),anchor="e")
        self.lbl_lane.pack(side=tk.RIGHT,padx=8)

        # Preview
        pf=tk.Frame(sf,bg=C["bg2"],highlightbackground=C["border"],highlightthickness=1,height=55)
        pf.pack(fill=tk.X,padx=14,pady=(0,10)); pf.pack_propagate(False)
        self.preview=tk.Label(pf,bg=C["bg2"],text="Preview - select region first",fg=C["dim2"],font=("Segoe UI",8))
        self.preview.pack(expand=True)

        # History
        self.lbl_history=tk.Label(sf,text="",bg=C["card"],fg=C["dim"],font=("Consolas",9),anchor="w")
        self.lbl_history.pack(fill=tk.X,padx=14,pady=(0,8))

        # ── STEP 1: Region ──
        self._sec(main,"STEP 1 — SELECT REGION (ลากเลือก ทั้งแถว ที่ตัวอักษรขึ้น)")
        rc=tk.Frame(main,bg=C["card"],highlightbackground=C["border"],highlightthickness=1)
        rc.pack(fill=tk.X,pady=(0,6))
        ri=tk.Frame(rc,bg=C["card"]); ri.pack(fill=tk.X,padx=12,pady=8)
        Btn(ri,text="Select Region (F6)",color=C["blue"],w=170,h=32,fs=10,cmd=self.pick_region).pack(side=tk.LEFT)
        Btn(ri,text="Test",color=C["orange"],w=60,h=32,fs=9,cmd=self._test).pack(side=tk.LEFT,padx=6)
        self.lbl_region=tk.Label(ri,text="Not set",bg=C["card"],fg=C["orange"],font=("Consolas",9))
        self.lbl_region.pack(side=tk.RIGHT)

        # ── STEP 2: Templates ──
        self._sec(main,"STEP 2 — SETUP KEYS (จับภาพ ทีละตัว จากเกม)")
        tc=tk.Frame(main,bg=C["card"],highlightbackground=C["border"],highlightthickness=1)
        tc.pack(fill=tk.X,pady=(0,6))
        ti=tk.Frame(tc,bg=C["card"]); ti.pack(fill=tk.X,padx=12,pady=8)
        Btn(ti,text="Quick Setup",color=C["green"],w=130,h=34,fs=10,cmd=self._quick_setup).pack(side=tk.LEFT)
        Btn(ti,text="+ 1 Key",color=C["purple"],w=90,h=34,fs=9,cmd=self._capture_one).pack(side=tk.LEFT,padx=6)
        Btn(ti,text="Clear",color=C["red"],w=70,h=34,fs=9,cmd=self._clear_tmpl).pack(side=tk.LEFT,padx=6)
        self.lbl_tmpls=tk.Label(ti,text="0",bg=C["card"],fg=C["dim"],font=("Segoe UI",9))
        self.lbl_tmpls.pack(side=tk.RIGHT)
        self.lbl_keys=tk.Label(tc,text="Keys: (none)",bg=C["card"],fg=C["cyan"],font=("Consolas",12,"bold"),anchor="w")
        self.lbl_keys.pack(padx=12,pady=(0,8))

        # ── STEP 3: Settings ──
        self._sec(main,"STEP 3 — SETTINGS")
        stf=tk.Frame(main,bg=C["card"],highlightbackground=C["border"],highlightthickness=1)
        stf.pack(fill=tk.X,pady=(0,6))
        sti=tk.Frame(stf,bg=C["card"]); sti.pack(fill=tk.X,padx=12,pady=8)

        self._slider(sti,"Keys per Lane","",self.v_num_keys,2,12,C["blue"])
        self._slider(sti,"Scan Speed","ms",self.v_scan,15,100,C["cyan"])
        self._slider(sti,"Match Threshold","%",self.v_thresh,50,95,C["orange"])
        self._slider(sti,"Delay between Keys","ms",self.v_key_delay,20,200,C["purple"])
        self._slider(sti,"Delay after Lane","ms",self.v_lane_delay,100,1000,C["green"])

        chk=tk.Frame(sti,bg=C["card"]); chk.pack(fill=tk.X,pady=(4,0))
        tk.Checkbutton(chk,text="Always on Top",variable=self.v_ontop,bg=C["card"],fg=C["white"],
                       selectcolor=C["bg"],activebackground=C["card"],font=("Segoe UI",9),
                       command=lambda:self.root.attributes("-topmost",self.v_ontop.get())).pack(side=tk.LEFT,padx=(0,12))
        tk.Checkbutton(chk,text="Sound",variable=self.v_sound,bg=C["card"],fg=C["white"],
                       selectcolor=C["bg"],activebackground=C["card"],font=("Segoe UI",9)).pack(side=tk.LEFT)

        # ── LOG ──
        self._sec(main,"LOG")
        lf=tk.Frame(main,bg=C["card"],highlightbackground=C["border"],highlightthickness=1)
        lf.pack(fill=tk.BOTH,expand=True,pady=(0,4))
        self.log_box=tk.Text(lf,bg=C["bg"],fg=C["green"],font=("Consolas",9),
                              height=4,relief="flat",wrap=tk.WORD,state=tk.DISABLED,padx=8,pady=6)
        self.log_box.pack(fill=tk.BOTH,expand=True,padx=1,pady=1)
        self.log_box.tag_configure("key",foreground=C["cyan"],font=("Consolas",9,"bold"))
        self.log_box.tag_configure("err",foreground=C["red"])
        self.log_box.tag_configure("warn",foreground=C["orange"])
        self.log_box.tag_configure("ok",foreground=C["green"])
        self.log_box.tag_configure("dim",foreground=C["dim"])
        self.log_box.tag_configure("lane",foreground=C["purple"],font=("Consolas",10,"bold"))

        ft=tk.Frame(self.root,bg=C["bg2"],height=26); ft.pack(fill=tk.X,side=tk.BOTTOM)
        ft.pack_propagate(False)
        tk.Label(ft,text="F5 Start/Stop   F6 Region   ESC Stop",bg=C["bg2"],fg=C["dim2"],font=("Segoe UI",8)).pack(pady=4)

        self.log(f"Capture: {self.cap.method or 'FAILED'}","ok" if self.cap.method else "err")
        if self.tmpl.count>0:
            self.log(f"Templates: {' '.join(k.upper() for k in self.tmpl.keys_list)}","ok")

    def _sec(self,p,t):
        f=tk.Frame(p,bg=C["bg"]); f.pack(fill=tk.X,pady=(6,2))
        tk.Label(f,text=t,bg=C["bg"],fg=C["dim2"],font=("Segoe UI",8,"bold")).pack(side=tk.LEFT,padx=2)
        s=tk.Canvas(f,bg=C["bg"],height=1,highlightthickness=0)
        s.pack(side=tk.LEFT,fill=tk.X,expand=True,padx=(6,0))
        s.create_line(0,0,400,0,fill=C["border"])

    def _slider(self,p,label,unit,var,lo,hi,color):
        row=tk.Frame(p,bg=C["card"]); row.pack(fill=tk.X,pady=2)
        tk.Label(row,text=f"{label}:",bg=C["card"],fg=C["white"],font=("Segoe UI",9)).pack(side=tk.LEFT)
        vl=tk.Label(row,text=f"{var.get()}{unit}",bg=C["card"],fg=color,
                     font=("Consolas",10,"bold"),width=6,anchor="e"); vl.pack(side=tk.RIGHT)
        tk.Scale(row,from_=lo,to=hi,orient=tk.HORIZONTAL,variable=var,
                 bg=C["card"],fg=C["card"],troughcolor=C["bg"],highlightthickness=0,
                 showvalue=False,length=160,sliderlength=14,activebackground=color,
                 command=lambda v,l=vl,u=unit:l.config(text=f"{int(float(v))}{u}")
                 ).pack(side=tk.RIGHT,padx=(4,2))

    # ══════════════════════════════════════════
    # Region
    # ══════════════════════════════════════════
    def pick_region(self):
        if self.running: self.log("Stop first!","err"); return
        self.root.iconify(); self.root.after(400,self._sel)
    def _sel(self):
        def done(r):
            self.region=r; self._save_cfg(); self.root.deiconify(); self._update_region()
            self.log(f"Region: {r['width']}x{r['height']} @ ({r['left']},{r['top']})","ok")
        s=RegionSelector(self.root,done)
        s.protocol("WM_DELETE_WINDOW",lambda:(s.destroy(),self.root.deiconify()))
    def _update_region(self):
        if self.region:
            r=self.region
            self.lbl_region.config(text=f"{r['width']}x{r['height']}  ({r['left']},{r['top']})",fg=C["green"])
        else: self.lbl_region.config(text="Not set",fg=C["orange"])
    def _update_tmpl(self):
        self.lbl_tmpls.config(text=f"{self.tmpl.count} templates")
        keys=self.tmpl.keys_list
        self.lbl_keys.config(text=f"Keys: {' '.join(k.upper() for k in keys)}" if keys else "Keys: (none)")

    def _test(self):
        if not self.region: messagebox.showwarning("Test","Select region first!"); return
        frame=self.cap.grab(self.region)
        if frame is None: self.log("Capture FAILED!","err"); return
        self.log(f"Capture OK  {frame.shape[1]}x{frame.shape[0]}","ok")
        if self.tmpl.count > 0:
            gray=cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY)
            results=self.tmpl.match_lane(gray,self.v_num_keys.get(),self.v_thresh.get()/100)
            seq_str=""
            for key,conf in results:
                if key: seq_str += f" {key.upper()}({conf:.0%})"
                else: seq_str += " ??"
            self.log(f"  Lane:{seq_str}","key")

    # ══════════════════════════════════════════
    # Templates
    # ══════════════════════════════════════════
    def _quick_setup(self):
        if not self.region: messagebox.showwarning("Setup","Select region first!"); return
        dlg=tk.Toplevel(self.root); dlg.title("Quick Setup"); dlg.geometry("420x350")
        dlg.configure(bg=C["bg"]); dlg.attributes("-topmost",True); dlg.resizable(False,False)

        tk.Label(dlg,text="Quick Setup",bg=C["bg"],fg=C["blue"],font=("Segoe UI",16,"bold")).pack(pady=(16,4))
        tk.Label(dlg,text="ให้ตัวอักษร 1 ตัว โชว์ในเกม (ครอป region ให้พอดี 1 ตัว)\nแล้วกด Capture ทำทีละตัวจนครบ",
                 bg=C["bg"],fg=C["dim"],font=("Segoe UI",10),justify="center").pack(pady=(0,8))

        ef=tk.Frame(dlg,bg=C["bg"]); ef.pack(pady=4)
        tk.Label(ef,text="Keys:",bg=C["bg"],fg=C["dim"],font=("Segoe UI",9)).pack(side=tk.LEFT)
        setup_entry=tk.Entry(ef,bg=C["input"],fg=C["cyan"],font=("Consolas",13),
                              relief="flat",width=14,justify="center")
        setup_entry.insert(0,"qweasd")
        setup_entry.pack(side=tk.LEFT,padx=6)

        idx_holder = [0]
        keys_holder = [list("qweasd")]

        def apply_keys():
            txt=setup_entry.get().strip().lower()
            if txt:
                keys_holder[0]=list(dict.fromkeys(txt))
                idx_holder[0]=0; update()

        Btn(ef,text="Set",color=C["blue"],w=50,h=28,fs=9,cmd=apply_keys).pack(side=tk.LEFT)

        cur_lbl=tk.Label(dlg,text="",bg=C["bg"],fg=C["green"],font=("Consolas",48,"bold"))
        cur_lbl.pack(pady=8)
        info_lbl=tk.Label(dlg,text="",bg=C["bg"],fg=C["dim"],font=("Segoe UI",10))
        info_lbl.pack()
        prev_lbl=tk.Label(dlg,bg=C["bg2"],width=10,height=3)
        prev_lbl.pack(pady=6)

        btn_f=tk.Frame(dlg,bg=C["bg"]); btn_f.pack(pady=6)

        def capture():
            if idx_holder[0]>=len(keys_holder[0]): return
            key=keys_holder[0][idx_holder[0]]
            frame=self.cap.grab(self.region)
            if frame is None: self.log("Capture failed!","err"); return
            fname=self.tmpl.save(key,frame)
            self.log(f"Saved: '{key.upper()}' -> {fname}","ok")
            self._update_tmpl()
            try:
                img=Image.fromarray(cv2.cvtColor(frame,cv2.COLOR_BGR2RGB)).resize((80,50),Image.LANCZOS)
                photo=ImageTk.PhotoImage(img); prev_lbl.config(image=photo); prev_lbl._photo=photo
            except: pass
            idx_holder[0]+=1; update()

        def skip():
            idx_holder[0]+=1; update()

        def update():
            if idx_holder[0]>=len(keys_holder[0]):
                cur_lbl.config(text="DONE!",fg=C["green"])
                info_lbl.config(text=f"จับครบ {len(keys_holder[0])} ตัว! ปิดได้เลย")
                return
            k=keys_holder[0][idx_holder[0]]
            cur_lbl.config(text=k.upper(),fg=C["green"])
            info_lbl.config(text=f"({idx_holder[0]+1}/{len(keys_holder[0])})  ให้ {k.upper()} โชว์ แล้วกด Capture")

        Btn(btn_f,text="Capture",color=C["green"],w=140,h=40,fs=13,cmd=capture).pack(side=tk.LEFT,padx=4)
        Btn(btn_f,text="Skip",color=C["dim"],w=80,h=40,fs=10,cmd=skip).pack(side=tk.LEFT,padx=4)
        update()

    def _capture_one(self):
        if not self.region: messagebox.showwarning("QTE","Select region first!"); return
        dlg=tk.Toplevel(self.root); dlg.title("Capture"); dlg.geometry("280x150")
        dlg.configure(bg=C["bg"]); dlg.attributes("-topmost",True)
        tk.Label(dlg,text="Key on screen?",bg=C["bg"],fg=C["white"],font=("Segoe UI",12)).pack(pady=(16,8))
        e=tk.Entry(dlg,bg=C["input"],fg=C["cyan"],font=("Consolas",22,"bold"),justify="center",width=4,relief="flat")
        e.pack(pady=4); e.focus_set()
        def do(event=None):
            k=e.get().strip().lower()
            if not k or len(k)!=1: return
            frame=self.cap.grab(self.region)
            if frame is None: self.log("Failed!","err"); return
            self.tmpl.save(k,frame); self._update_tmpl()
            self.log(f"Saved: '{k.upper()}'","ok"); dlg.destroy()
        e.bind("<Return>",do)
        Btn(dlg,text="Capture!",color=C["purple"],w=120,h=34,fs=11,cmd=do).pack(pady=8)

    def _clear_tmpl(self):
        if not messagebox.askyesno("Clear","Delete all templates?"): return
        import shutil
        if os.path.isdir(self.tmpl.dir): shutil.rmtree(self.tmpl.dir)
        os.makedirs(self.tmpl.dir,exist_ok=True)
        self.tmpl.load(); self._update_tmpl()
        self.log("Cleared all templates","warn")

    # ══════════════════════════════════════════
    # Macro
    # ══════════════════════════════════════════
    def toggle(self):
        if self.running: self._stop()
        else: self._start()

    def _start(self):
        if not self.region: messagebox.showwarning("QTE","Select region first! (F6)"); return
        if not self.cap.method: messagebox.showerror("QTE","Capture not working!"); return
        if self.tmpl.count==0: messagebox.showwarning("QTE","No templates! Quick Setup first."); return

        self.lane.reset()
        self.running=True; self.session_keys=0; self.session_start=time.time()
        self.key_history=[]
        self.btn_start.set_text("STOP (F5)"); self.btn_start.set_color(C["red"])
        self.hdr.itemconfig(self.lbl_status,text="ON",fill=C["green"])
        self.log("Macro STARTED","ok")
        self._save_cfg()
        threading.Thread(target=self._loop,daemon=True).start()

    def _stop(self):
        if not self.running: return
        self.running=False
        self.btn_start.set_text("START (F5)"); self.btn_start.set_color(C["green"])
        self.hdr.itemconfig(self.lbl_status,text="OFF",fill=C["dim"])
        e=time.time()-(self.session_start or time.time())
        self.log(f"Stopped ({self.session_keys} keys / {e:.1f}s)")

    def _loop(self):
        thresh = self.v_thresh.get()/100
        num_keys = self.v_num_keys.get()
        key_delay = self.v_key_delay.get()/1000.0
        lane_delay = self.v_lane_delay.get()/1000.0

        while self.running:
            try:
                frame = self.cap.grab(self.region)
                if frame is None: time.sleep(0.05); continue
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

                # 1. Scan ทั้ง lane
                results = self.tmpl.match_lane(gray, num_keys, thresh)

                # กรองเอาเฉพาะที่ match ได้
                valid = [(k,c) for k,c in results if k is not None]

                if len(valid) == 0:
                    time.sleep(self.v_scan.get()/1000.0)
                    continue

                # 2. เช็คว่าเป็น lane ใหม่หรือ lane เดิม
                if not self.lane.is_new_lane(results):
                    # Lane เดิม ที่กดไปแล้ว → รอ
                    time.sleep(self.v_scan.get()/1000.0)
                    continue

                # 3. Lane ใหม่! → กดทีละตัว
                self.lane.start_sequence(results)
                seq_text = " ".join(k.upper() if k else "?" for k,c in results)
                self.root.after(0, self._show_lane, seq_text)
                self.root.after(0, self.log, f"  Lane: {seq_text}", "lane")

                while self.lane.pressing and self.running:
                    key, conf = self.lane.next_key()
                    if key:
                        kb.press_and_release(key)
                        self.session_keys += 1
                        self.root.after(0, self._on_key, key, conf)

                        if self.v_sound.get():
                            try:
                                import winsound; winsound.Beep(900,20)
                            except: pass

                        time.sleep(key_delay)

                # 4. กดครบ → รอก่อนสแกน lane ถัดไป
                time.sleep(lane_delay)

            except Exception as e:
                self.root.after(0, self.log, str(e), "err")
                time.sleep(0.1)

    def _show_lane(self, text):
        self.lbl_lane.config(text=text)
        self.lbl_lane_status.config(text="Pressing...", fg=C["green"])

    def _on_key(self, key, conf):
        self.lbl_count.config(text=str(self.session_keys))
        self.lbl_debug.config(text=f"Last: {key.upper()} ({conf:.0%})",
                               fg=C["green"] if conf>0.85 else C["orange"])
        e=time.time()-(self.session_start or time.time())
        if e>0: self.lbl_speed.config(text=f"{self.session_keys/e:.1f} /sec")
        self.key_history.append(key.upper())
        if len(self.key_history)>30: self.key_history.pop(0)
        self.lbl_history.config(text="History: "+" ".join(self.key_history[-20:]))

    # ══════════════════════════════════════════
    # Loops
    # ══════════════════════════════════════════
    def _preview_loop(self):
        if self.region and self.cap.method:
            try:
                img=self.cap.grab_pil(self.region)
                if img:
                    w,h=img.size; nw=min(520,w); nh=int(h*(nw/w))
                    img=img.resize((nw,max(nh,20)),Image.LANCZOS)
                    photo=ImageTk.PhotoImage(img)
                    self.preview.config(image=photo,text=""); self.preview._photo=photo
            except: pass
        self.root.after(500,self._preview_loop)

    def _timer_loop(self):
        if self.running and self.session_start:
            e=int(time.time()-self.session_start); m,s=divmod(e,60)
            self.hdr.itemconfig(self.lbl_timer,text=f"{m:02d}:{s:02d}",fill=C["green"])
        else:
            self.hdr.itemconfig(self.lbl_timer,text="00:00",fill=C["dim2"])
            if not self.running:
                self.lbl_lane_status.config(text="Idle",fg=C["dim"])
        self.root.after(500,self._timer_loop)

    def log(self,msg,tag=None):
        def _do():
            self.log_box.config(state=tk.NORMAL)
            ts=datetime.now().strftime("%H:%M:%S")
            if tag:
                self.log_box.insert(tk.END,f"[{ts}] ","dim")
                self.log_box.insert(tk.END,msg+"\n",tag)
            else: self.log_box.insert(tk.END,f"[{ts}] {msg}\n")
            self.log_box.see(tk.END)
            lines=int(self.log_box.index("end-1c").split(".")[0])
            if lines>200: self.log_box.delete("1.0",f"{lines-200}.0")
            self.log_box.config(state=tk.DISABLED)
        self.root.after(0,_do)

    def _quit(self):
        self.running=False; self._save_cfg(); self.cap.close()
        time.sleep(0.1); self.root.destroy()

    def run(self): self.root.mainloop()

if __name__=="__main__":
    App().run()
