"""
QTE Macro - Auto Learn Edition
ครั้งแรก: เล่นเองปกติ → macro จำ font เกมเอง
ครั้งถัดไป: macro กดให้เอง 100%
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
# Game Templates - จำ font จากเกมจริง
# ═══════════════════════════════════
class GameKeys:
    def __init__(self):
        self.templates = {}  # key -> [gray images height=48]
        self.H = 48
        self.tmpl_dir = os.path.join(get_base_path(), "game_templates")
        os.makedirs(self.tmpl_dir, exist_ok=True)
        self._load()

    def _load(self):
        """โหลด template ที่เคยเรียนรู้ไว้"""
        self.templates.clear()
        for f in os.listdir(self.tmpl_dir):
            if not f.endswith(".png"): continue
            key = f.split("_")[0].split(".")[0].lower()
            if len(key) != 1: continue
            img = cv2.imread(os.path.join(self.tmpl_dir, f), cv2.IMREAD_GRAYSCALE)
            if img is not None:
                h,w = img.shape[:2]
                nw = max(int(w * self.H / h), 3)
                self.templates.setdefault(key, []).append(cv2.resize(img, (nw, self.H)))

    def learn(self, key, gray_slot):
        """เรียนรู้ตัวอักษรใหม่จากเกม"""
        key = key.lower()
        n = len([f for f in os.listdir(self.tmpl_dir) if f.startswith(key)])
        fname = f"{key}.png" if n == 0 else f"{key}_{n+1}.png"
        cv2.imwrite(os.path.join(self.tmpl_dir, fname), gray_slot)
        h,w = gray_slot.shape[:2]
        nw = max(int(w * self.H / h), 3)
        self.templates.setdefault(key, []).append(cv2.resize(gray_slot, (nw, self.H)))

    def match(self, gray_roi, threshold=0.55):
        """Match ROI กับ template ที่เรียนรู้มา"""
        if not self.templates or gray_roi is None or gray_roi.size < 20:
            return None, 0
        h,w = gray_roi.shape[:2]
        if h < 3 or w < 3: return None, 0
        nw = max(int(w * self.H / h), 3)
        roi = cv2.resize(gray_roi, (nw, self.H))

        best_c, best_v = None, 0
        for ch, tmpls in self.templates.items():
            for t in tmpls:
                tw = t.shape[1]
                if tw > roi.shape[1]*1.5 or tw < roi.shape[1]*0.4: continue
                target = roi
                if tw > roi.shape[1] or t.shape[0] > roi.shape[0]:
                    px = max(0, (tw-roi.shape[1])//2+3)
                    py = max(0, (t.shape[0]-roi.shape[0])//2+3)
                    target = cv2.copyMakeBorder(roi, py,py,px,px, cv2.BORDER_CONSTANT, value=0)
                try:
                    res = cv2.matchTemplate(target, t, cv2.TM_CCOEFF_NORMED)
                    _,mx,_,_ = cv2.minMaxLoc(res)
                    if mx > threshold and mx > best_v:
                        best_c, best_v = ch, mx
                        if mx > 0.85: return best_c, best_v
                except: continue
        return best_c, best_v

    @property
    def count(self): return sum(len(v) for v in self.templates.values())
    @property
    def learned(self): return sorted(self.templates.keys())


# ═══════════════════════════════════
# Utils
# ═══════════════════════════════════
def get_slot(gray, i, n):
    h,w = gray.shape[:2]; sw = w//n; pad = max(sw//10,2)
    return gray[:, max(0,i*sw-pad):min(w,(i+1)*sw+pad)]

def is_slot_done(orig_slot, curr_slot, thresh=20):
    """เช็คว่า slot เปลี่ยนจากตอนเริ่ม round (กดไปแล้ว)"""
    if orig_slot is None or curr_slot is None: return False
    diff = cv2.absdiff(orig_slot, curr_slot)
    if len(diff.shape) == 3:
        diff = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
    return (np.count_nonzero(diff > 30) / max(diff.size,1)) * 100 > thresh

def prep(gray):
    """Threshold → crop ตัวอักษร"""
    if gray is None or gray.size < 20: return None
    h,w = gray.shape[:2]
    if h < 20: gray = cv2.resize(gray, None, fx=max(30/h,1), fy=max(30/h,1), interpolation=cv2.INTER_CUBIC)
    _, t = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
    ti = cv2.bitwise_not(t)
    r = t if np.count_nonzero(t) < np.count_nonzero(ti) else ti
    cs = np.argwhere(r > 127)
    if len(cs) < 5: return None
    y0,x0 = cs.min(0); y1,x1 = cs.max(0)
    c = r[max(0,y0-1):y1+2, max(0,x0-1):x1+2]
    return c if c.shape[0]>3 and c.shape[1]>3 else None

def split_small(frame, n):
    """แบ่ง frame เป็น N slots เล็กๆ (สำหรับเทียบสี)"""
    h,w = frame.shape[:2]; sw = w//n
    return [cv2.resize(frame[:, i*sw:min((i+1)*sw,w)], (32,32)) for i in range(n)]


# ═══════════════════════════════════
# State
# ═══════════════════════════════════
cap = Cap()
game = GameKeys()
running = False
region = None
session_keys = 0
session_start = 0
learning_key = None   # key ที่ user กดตอน learn
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
# Keyboard Listener - จับว่า user กดปุ่มอะไร (ตอน learn)
# ═══════════════════════════════════
def on_user_key(event):
    global learning_key
    k = event.name.lower()
    if k in "qweasd" and running:
        learning_key = k

kb.on_press(on_user_key)


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
    c.create_rectangle(sw//2-180,16,sw//2+180,60,fill="#000",outline="#00ffaa")
    c.create_text(sw//2,30,text="ลากครอบแถวตัวอักษร (ไม่รวม counter)",fill="#00ffaa",font=("Segoe UI",12,"bold"))
    c.create_text(sw//2,50,text="ESC = ยกเลิก",fill="#aaa",font=("Segoe UI",9))
    st={"sx":0,"sy":0,"r":None}
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
            eel.onLog(f"Region set!","ok")
        root.destroy()
    c.bind("<ButtonPress-1>",p); c.bind("<B1-Motion>",d); c.bind("<ButtonRelease-1>",r)
    root.bind("<Escape>",lambda e:root.destroy())
    root.after(50,root.focus_force); root.mainloop()

@eel.expose
def toggle(num_keys, scan_ms, key_delay_ms, lane_delay_ms):
    global running
    if running: running = False; return
    if not region: eel.onLog("Select region first! (F6)","warn"); return
    running = True; save_cfg(num_keys=num_keys)
    eel.onStatus(True)

    if game.count > 0:
        eel.onLog(f"มี template: {' '.join(k.upper() for k in game.learned)} → AUTO","ok")
    else:
        eel.onLog("ยังไม่มี template → เล่นเองรอบแรก macro จะจำเอง","warn")

    threading.Thread(target=run_loop, args=(num_keys,scan_ms/1000,key_delay_ms/1000,lane_delay_ms/1000), daemon=True).start()

@eel.expose
def test(num_keys):
    if not region: eel.onLog("Select region first!","warn"); return
    frame = cap.grab(region)
    if frame is None: eel.onLog("Capture FAILED","err"); return
    b = cap.preview_b64(region)
    if b: eel.onPreview(b)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    for i in range(num_keys):
        roi = prep(get_slot(gray, i, num_keys))
        ch, cf = game.match(roi)
        eel.onLog(f"  Slot {i+1}: {ch.upper() if ch else '??'} ({cf:.0%})", "key" if ch else "warn")
    eel.onLog(f"Learned keys: {' '.join(k.upper() for k in game.learned) or 'none'}", "dim")

@eel.expose
def clear_learned():
    import shutil
    if os.path.isdir(game.tmpl_dir): shutil.rmtree(game.tmpl_dir)
    os.makedirs(game.tmpl_dir, exist_ok=True)
    game.templates.clear()
    eel.onLog("Cleared all templates","ok")


# ═══════════════════════════════════
# Main Loop
# ═══════════════════════════════════
def run_loop(num_keys, sd, kd, ld):
    """
    Simple flow:
    1. จับภาพ "ตอนเริ่ม round"
    2. หา slot แรกที่ยังไม่เปลี่ยน = active
    3. ถ้ามี template → match → กดเลย
    4. ถ้าไม่มี template → รอ user กดเอง → จำ
    5. slot เปลี่ยนสี → ไป slot ถัดไป
    6. ทุก slot เปลี่ยน → จบ round → รอ round ใหม่
    """
    global running, session_keys, session_start, learning_key
    session_keys = 0; session_start = time.time()
    learning_key = None

    orig_slots = None  # ภาพ slot ตอนเริ่ม round
    completed_small = None  # ภาพย่อตอนจบ round
    last_active = -1

    while running:
        try:
            frame = cap.grab(region)
            if frame is None: time.sleep(0.03); continue
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

            # ═══ ยังไม่มี round (รอ round ใหม่) ═══
            if orig_slots is None:
                # เช็คว่าหน้าจอเปลี่ยนจากตอนจบ round ก่อนหน้า
                if completed_small is not None:
                    small = cv2.resize(frame, (64,64))
                    diff = cv2.absdiff(completed_small, small)
                    change = (np.count_nonzero(cv2.cvtColor(diff,cv2.COLOR_BGR2GRAY)>30) / diff[:,:,0].size) * 100
                    if change < 15:
                        time.sleep(sd); continue  # ยังเหมือนเดิม รอต่อ

                # จับภาพเริ่ม round
                orig_slots = split_small(frame, num_keys)
                last_active = -1
                eel.onLog("Round start!", "lane")
                time.sleep(0.05)
                continue

            # ═══ มี round → หา slot active ═══
            curr_slots = split_small(frame, num_keys)
            active = -1
            for i in range(num_keys):
                if not is_slot_done(orig_slots[i], curr_slots[i]):
                    active = i; break

            # ส่ง debug
            try: eel.onDebug(active, num_keys)
            except: pass

            if active == -1:
                # ทุก slot เปลี่ยน → จบ round
                eel.onLog("Round done!","ok")
                completed_small = cv2.resize(frame, (64,64))
                orig_slots = None
                time.sleep(ld)
                continue

            # ═══ อ่านตัวอักษรจาก active slot ═══
            roi_gray = get_slot(gray, active, num_keys)
            roi_processed = prep(roi_gray)

            ch, cf = game.match(roi_processed)

            if ch and cf > 0.55:
                # ═══ MATCH! → กดเลย ═══
                kb.press_and_release(ch)
                session_keys += 1
                elapsed = time.time() - session_start
                spd = session_keys / elapsed if elapsed > 0 else 0
                eel.onKeyPressed(ch, session_keys, spd)
                eel.onLog(f"  [{session_keys}] {ch.upper()} ({cf:.0%})", "key")
                last_active = active
                time.sleep(kd)

            else:
                # ═══ ไม่รู้จักตัวนี้ → รอ user กดเอง → จำ ═══
                if active != last_active:
                    eel.onLog(f"  Slot {active+1}: อ่านไม่ออก → กดเองได้เลย (macro จะจำ)", "warn")
                    last_active = active
                    learning_key = None

                # เช็คว่า user กดปุ่มอะไร
                if learning_key and roi_processed is not None:
                    game.learn(learning_key, roi_processed)
                    eel.onLog(f"  Learned: {learning_key.upper()} ✓", "ok")
                    eel.onLearned(game.learned)

                    # กดปุ่มนั้นให้ด้วย (เผื่อ user กดไม่ทัน)
                    # ไม่ต้องกดซ้ำ เพราะ user กดเองแล้ว
                    session_keys += 1
                    elapsed = time.time() - session_start
                    spd = session_keys / elapsed if elapsed > 0 else 0
                    eel.onKeyPressed(learning_key, session_keys, spd)

                    last_active = active
                    learning_key = None
                    time.sleep(kd)
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
    script_dir = os.path.dirname(os.path.abspath(__file__))
    web_dir = os.path.join(script_dir, "web")
    if not os.path.exists(os.path.join(web_dir, "index.html")):
        print(f"ERROR: {web_dir}/index.html not found"); sys.exit(1)

    eel.init(web_dir)
    print(f"Game templates: {game.count} | Keys: {', '.join(k.upper() for k in game.learned) or 'none'}")
    print(f"Capture: {'OK' if cap.ok else 'FAIL'}")
    threading.Thread(target=bg, daemon=True).start()

    for mode in ["chrome-app","chrome","edge","default"]:
        try: eel.start("index.html", size=(480,700), port=0, mode=mode); break
        except: continue
    else: print("Need Chrome/Edge!")
