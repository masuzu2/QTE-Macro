"""
QTE Macro - Simple Edition
เห็นตัวอักษร → กดเลย ไม่ต้อง setup
"""

import time, sys, os, json, threading, base64, io
import eel

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

MISSING = []
for lib, pkg in [("keyboard","keyboard"),("mss","mss"),("cv2","opencv-python"),
                  ("numpy","numpy"),("PIL","Pillow"),("eel","eel")]:
    try: __import__(lib)
    except ImportError: MISSING.append(pkg)
if MISSING:
    print(f"pip install {' '.join(MISSING)}"); sys.exit(1)

import keyboard as kb
import mss
import cv2, numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageGrab


# ═══════════════════════════════════
# Capture
# ═══════════════════════════════════
class Cap:
    def __init__(self):
        try: self.sct = mss.mss(); self.sct.grab(self.sct.monitors[0]); self.ok = True
        except: self.sct = None; self.ok = False

    def grab(self, r):
        try:
            if self.sct:
                return cv2.cvtColor(np.array(self.sct.grab(r)), cv2.COLOR_BGRA2BGR)
            return cv2.cvtColor(np.array(ImageGrab.grab(
                bbox=(r["left"],r["top"],r["left"]+r["width"],r["top"]+r["height"]))), cv2.COLOR_RGB2BGR)
        except: return None

    def preview_b64(self, r):
        try:
            img = ImageGrab.grab(bbox=(r["left"],r["top"],r["left"]+r["width"],r["top"]+r["height"]))
            w,h = img.size
            if w > 460: img = img.resize((460, max(int(h*460/w),1)), Image.NEAREST)
            buf = io.BytesIO(); img.save(buf, format="JPEG", quality=50)
            return base64.b64encode(buf.getvalue()).decode()
        except: return None


# ═══════════════════════════════════
# Auto Templates - QWEASD เท่านั้น สร้างจาก font ในเครื่อง
# ═══════════════════════════════════
class AutoKeys:
    def __init__(self):
        self.keys = {}  # char -> [gray templates normalized to height 48]
        self.H = 48
        self._build()

    def _build(self):
        fonts = []
        for n in ["arialbd.ttf","arial.ttf","impact.ttf","calibrib.ttf",
                   "segoeuib.ttf","tahomabd.ttf","verdanab.ttf","consolab.ttf"]:
            p = os.path.join(os.environ.get("WINDIR","C:\\Windows"),"Fonts",n)
            if os.path.exists(p): fonts.append(p)
        if not fonts: fonts = [None]

        for ch in "qweasd":
            self.keys[ch] = []
            for fp in fonts[:6]:
                for sz in [36, 48, 60, 72]:
                    for case in [ch.upper(), ch.lower()]:
                        for bold in [False, True]:
                            t = self._render(case, fp, sz, bold)
                            if t is not None:
                                h,w = t.shape[:2]
                                self.keys[ch].append(cv2.resize(t, (max(int(w*self.H/h),3), self.H)))

    def _render(self, ch, fp, sz, bold):
        try:
            img = Image.new("L",(sz+16,sz+16),0)
            d = ImageDraw.Draw(img)
            f = ImageFont.truetype(fp, sz) if fp else ImageFont.load_default()
            bb = d.textbbox((0,0),ch,font=f); tw,th = bb[2]-bb[0],bb[3]-bb[1]
            d.text(((img.width-tw)//2-bb[0],(img.height-th)//2-bb[1]),ch,fill=255,font=f)
            a = np.array(img)
            if bold: a = cv2.dilate(a, cv2.getStructuringElement(cv2.MORPH_ELLIPSE,(3,3)))
            cs = np.argwhere(a>40)
            if len(cs)<5: return None
            y0,x0 = cs.min(0); y1,x1 = cs.max(0)
            cr = a[max(0,y0-1):y1+2, max(0,x0-1):x1+2]
            if cr.shape[0]<4 or cr.shape[1]<4: return None
            _, b = cv2.threshold(cr,80,255,cv2.THRESH_BINARY)
            return b
        except: return None

    def match(self, gray_roi, threshold=0.45):
        """อ่านตัวอักษรจาก ROI → return (char, confidence)"""
        if gray_roi is None or gray_roi.size < 20: return None, 0
        h,w = gray_roi.shape[:2]
        roi = cv2.resize(gray_roi, (max(int(w*self.H/h),3), self.H))
        best_c, best_v = None, 0
        for ch, tmpls in self.keys.items():
            for t in tmpls:
                tw = t.shape[1]
                if tw > roi.shape[1]*1.5 or tw < roi.shape[1]*0.3: continue
                target = roi
                if tw > roi.shape[1] or t.shape[0] > roi.shape[0]:
                    px = max(0,(tw-roi.shape[1])//2+3)
                    py = max(0,(t.shape[0]-roi.shape[0])//2+3)
                    target = cv2.copyMakeBorder(roi,py,py,px,px,cv2.BORDER_CONSTANT,value=0)
                try:
                    res = cv2.matchTemplate(target, t, cv2.TM_CCOEFF_NORMED)
                    _,mx,_,_ = cv2.minMaxLoc(res)
                    if mx > threshold and mx > best_v:
                        best_c, best_v = ch, mx
                        if mx > 0.80: return best_c, best_v
                except: continue
        return best_c, best_v


# ═══════════════════════════════════
# Core Logic - ง่ายที่สุดเท่าที่จะเป็นไปได้
# ═══════════════════════════════════
def get_slot(gray, i, n):
    """ตัด slot i จาก gray (มี n ช่อง)"""
    h,w = gray.shape[:2]
    sw = w // n
    pad = max(sw//10, 2)
    x1,x2 = max(0, i*sw-pad), min(w, (i+1)*sw+pad)
    return gray[:, x1:x2]

def is_slot_pink(frame_bgr, i, n):
    """เช็คว่า slot i เป็นสีชมพูหรือยัง"""
    h,w = frame_bgr.shape[:2]
    sw = w // n
    slot = frame_bgr[:, i*sw:min((i+1)*sw, w)]
    # Pink ใน BGR: B=150-200, G=100-170, R=180-255
    b,g,r = cv2.split(slot)
    pink = (r > 160) & (b > 120) & (g < 180) & (r > b)
    return np.count_nonzero(pink) / max(pink.size, 1) > 0.10

def prep(gray):
    """Threshold ให้ตัวอักษรเป็นขาวบนดำ + crop"""
    if gray is None or gray.size < 20: return None
    h,w = gray.shape[:2]
    if h < 20:
        s = max(30/h, 1)
        gray = cv2.resize(gray, None, fx=s, fy=s, interpolation=cv2.INTER_CUBIC)
    _, t = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
    ti = cv2.bitwise_not(t)
    r = t if np.count_nonzero(t) < np.count_nonzero(ti) else ti
    cs = np.argwhere(r > 127)
    if len(cs) < 5: return None
    y0,x0 = cs.min(0); y1,x1 = cs.max(0)
    c = r[max(0,y0-1):y1+2, max(0,x0-1):x1+2]
    return c if c.shape[0]>3 and c.shape[1]>3 else None


# ═══════════════════════════════════
# State
# ═══════════════════════════════════
cap = Cap()
keys = AutoKeys()
running = False
region = None
session_keys = 0
session_start = 0
cfg_path = os.path.join(get_base_path(), "macro_config.json")

def load_cfg():
    global region
    try:
        with open(cfg_path,"r") as f: c=json.load(f); region=c.get("region"); return c
    except: return {}
def save_cfg(**kw):
    c=load_cfg(); c.update(kw); c["region"]=region
    try:
        with open(cfg_path,"w") as f: json.dump(c,f,indent=2)
    except: pass
load_cfg()


# ═══════════════════════════════════
# EEL
# ═══════════════════════════════════
@eel.expose
def select_region():
    global region
    import tkinter as tk
    try: ss = ImageGrab.grab()
    except: eel.onLog("Screenshot failed","err"); return
    root = tk.Tk(); root.overrideredirect(True); root.attributes("-topmost",True)
    sw,sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{sw}x{sh}+0+0")
    try:
        from PIL import ImageTk; photo = ImageTk.PhotoImage(ss)
    except: photo = None
    c = tk.Canvas(root, highlightthickness=0, cursor="cross"); c.pack(fill=tk.BOTH, expand=True)
    if photo: c.create_image(0,0,anchor="nw",image=photo)
    c.create_rectangle(0,0,sw,sh,fill="black",stipple="gray25")
    c.create_rectangle(sw//2-180,16,sw//2+180,62,fill="#000",outline="#00ffaa")
    c.create_text(sw//2,30,text="ลากครอบแถวตัวอักษร (ไม่รวม counter)",fill="#00ffaa",font=("Segoe UI",12,"bold"))
    c.create_text(sw//2,50,text="ESC = ยกเลิก",fill="#aaa",font=("Segoe UI",9))
    st = {"sx":0,"sy":0,"r":None}
    def p(e):
        st["sx"],st["sy"]=e.x,e.y
        if st["r"]: c.delete(st["r"])
        st["r"]=c.create_rectangle(e.x,e.y,e.x,e.y,outline="#00ffaa",width=2)
    def d(e):
        c.coords(st["r"],st["sx"],st["sy"],e.x,e.y); c.delete("sz")
        c.create_text((st["sx"]+e.x)//2,min(st["sy"],e.y)-12,text=f"{abs(e.x-st['sx'])}x{abs(e.y-st['sy'])}",
            fill="#00ffaa",font=("Consolas",11,"bold"),tags="sz")
    def r(e):
        global region
        x1,y1=min(st["sx"],e.x),min(st["sy"],e.y); x2,y2=max(st["sx"],e.x),max(st["sy"],e.y)
        if (x2-x1)>10 and (y2-y1)>10:
            region={"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}; save_cfg()
            eel.onRegionSet(f"{x2-x1}x{y2-y1} ({x1},{y1})")
            eel.onLog(f"Region: {x2-x1}x{y2-y1}","ok")
        root.destroy()
    c.bind("<ButtonPress-1>",p); c.bind("<B1-Motion>",d); c.bind("<ButtonRelease-1>",r)
    root.bind("<Escape>",lambda e:root.destroy())
    root.after(50,root.focus_force); root.mainloop()

@eel.expose
def toggle(num_keys, scan_ms, key_delay_ms, lane_delay_ms):
    global running
    if running: running = False; return
    if not region: eel.onLog("Select region first! (F6)","warn"); return
    running = True
    save_cfg(num_keys=num_keys)
    eel.onStatus(True)
    eel.onLog("START","ok")
    threading.Thread(target=run_loop, args=(num_keys,scan_ms/1000,key_delay_ms/1000,lane_delay_ms/1000), daemon=True).start()

@eel.expose
def test(num_keys):
    if not region: eel.onLog("Select region first!","warn"); return
    frame = cap.grab(region)
    if frame is None: eel.onLog("Capture FAILED","err"); return
    b = cap.preview_b64(region)
    if b: eel.onPreview(b)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    parts = []
    for i in range(num_keys):
        pink = is_slot_pink(frame, i, num_keys)
        roi = prep(get_slot(gray, i, num_keys))
        ch, cf = keys.match(roi)
        mark = "PINK" if pink else ">>>"
        parts.append(f"{ch.upper() if ch else '??'}({cf:.0%}){mark}")
    eel.onLog("Test: " + " | ".join(parts), "key")


# ═══════════════════════════════════
# Main Loop - เห็นอะไร กดอันนั้น
# ═══════════════════════════════════
def run_loop(num_keys, sd, kd, ld):
    global running, session_keys, session_start
    session_keys = 0; session_start = time.time()

    while running:
        try:
            frame = cap.grab(region)
            if frame is None: time.sleep(0.03); continue
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

            pressed_any = False

            for i in range(num_keys):
                if not running: break

                # ข้ามตัวที่เป็นสีชมพูแล้ว (กดไปแล้ว)
                if is_slot_pink(frame, i, num_keys):
                    continue

                # อ่านตัวอักษร
                roi = prep(get_slot(gray, i, num_keys))
                ch, cf = keys.match(roi)

                if ch:
                    # เห็นตัวอักษร → กดเลย!
                    kb.press_and_release(ch)
                    session_keys += 1
                    elapsed = time.time() - session_start
                    spd = session_keys / elapsed if elapsed > 0 else 0
                    eel.onKeyPressed(ch, session_keys, spd)
                    eel.onLog(f"  [{session_keys}] {ch.upper()} (slot {i+1}, {cf:.0%})", "key")
                    pressed_any = True
                    time.sleep(kd)

                    # กดแล้ว break ออกมาสแกนใหม่ (เพราะสีจะเปลี่ยน)
                    break

            if not pressed_any:
                # ไม่เจอตัวที่ต้องกด (ทุกตัวชมพูหมด หรือว่างอยู่) → รอ
                time.sleep(ld if all(is_slot_pink(frame,i,num_keys) for i in range(num_keys)) else sd)
            else:
                time.sleep(sd)

        except Exception as e:
            eel.onLog(str(e), "err")
            time.sleep(0.1)

    eel.onStatus(False)
    eel.onLog("STOPPED","ok")


# ═══════════════════════════════════
# Background
# ═══════════════════════════════════
def bg():
    while True:
        try:
            if region:
                b = cap.preview_b64(region)
                if b: eel.onPreview(b)
            if running and session_start:
                e=int(time.time()-session_start); m,s=divmod(e,60)
                eel.onTimerUpdate(f"{m:02d}:{s:02d}")
            else: eel.onTimerUpdate("00:00")
        except: pass
        time.sleep(0.5)


# ═══════════════════════════════════
# Start
# ═══════════════════════════════════
if __name__ == "__main__":
    # หา web folder ให้ถูกต้อง ไม่ว่าจะรันจาก directory ไหน
    script_dir = os.path.dirname(os.path.abspath(__file__))
    web_dir = os.path.join(script_dir, "web")

    if not os.path.exists(web_dir):
        print(f"ERROR: web folder not found at {web_dir}")
        sys.exit(1)
    if not os.path.exists(os.path.join(web_dir, "index.html")):
        print(f"ERROR: index.html not found in {web_dir}")
        sys.exit(1)

    print(f"Web dir: {web_dir}")
    eel.init(web_dir)

    n = sum(len(v) for v in keys.keys.values())
    print(f"Templates: {n} | Capture: {'OK' if cap.ok else 'FAIL'}")
    threading.Thread(target=bg, daemon=True).start()

    for mode in ["chrome-app", "chrome", "edge", "default"]:
        try:
            eel.start("index.html", size=(480, 700), port=0, mode=mode)
            break
        except Exception as e:
            print(f"Mode '{mode}' failed: {e}")
            continue
    else:
        print("Cannot open browser! Install Chrome or Edge.")
