"""
QTE Macro - Final Simple
1. Select region ครอบแถวตัวอักษร
2. START → เล่นเองรอบแรก macro จะจำ font เกม
3. รอบถัดไป macro กดให้เอง
"""

import time, sys, os, json, threading, base64, io
import eel

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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
from PIL import Image, ImageGrab


# ═══════════════════════════════════
# Capture
# ═══════════════════════════════════
cap_sct = None
try: cap_sct = mss.mss()
except: pass

def grab(r):
    try:
        if cap_sct:
            return cv2.cvtColor(np.array(cap_sct.grab(r)), cv2.COLOR_BGRA2BGR)
        return cv2.cvtColor(np.array(ImageGrab.grab(
            bbox=(r["left"],r["top"],r["left"]+r["width"],r["top"]+r["height"]))), cv2.COLOR_RGB2BGR)
    except: return None

def grab_b64(r):
    try:
        img = ImageGrab.grab(bbox=(r["left"],r["top"],r["left"]+r["width"],r["top"]+r["height"]))
        w,h = img.size
        if w > 460: img = img.resize((460, max(int(h*460/w),1)), Image.NEAREST)
        buf = io.BytesIO(); img.save(buf, format="JPEG", quality=50)
        return base64.b64encode(buf.getvalue()).decode()
    except: return None


# ═══════════════════════════════════
# Game Templates
# ═══════════════════════════════════
TMPL_DIR = os.path.join(SCRIPT_DIR, "game_templates")
os.makedirs(TMPL_DIR, exist_ok=True)
templates = {}  # key -> [gray images h=48]
TH = 48

def load_templates():
    global templates
    templates.clear()
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

def match_template(gray_roi, threshold=0.55):
    if not templates or gray_roi is None or gray_roi.size < 20: return None, 0
    h,w = gray_roi.shape[:2]
    if h<3 or w<3: return None, 0
    roi = cv2.resize(gray_roi, (max(int(w*TH/h),3), TH))
    best_c, best_v = None, 0
    for ch, tmpls in templates.items():
        for t in tmpls:
            tw = t.shape[1]
            if tw > roi.shape[1]*1.5 or tw < roi.shape[1]*0.4: continue
            target = roi
            if tw > roi.shape[1] or t.shape[0] > roi.shape[0]:
                px,py = max(0,(tw-roi.shape[1])//2+3), max(0,(t.shape[0]-roi.shape[0])//2+3)
                target = cv2.copyMakeBorder(roi,py,py,px,px,cv2.BORDER_CONSTANT,value=0)
            try:
                res = cv2.matchTemplate(target, t, cv2.TM_CCOEFF_NORMED)
                _,mx,_,_ = cv2.minMaxLoc(res)
                if mx > threshold and mx > best_v:
                    best_c, best_v = ch, mx
                    if mx > 0.85: return best_c, best_v
            except: continue
    return best_c, best_v

def learned_keys(): return sorted(templates.keys())

load_templates()


# ═══════════════════════════════════
# Utils
# ═══════════════════════════════════
def get_slot(gray, i, n):
    h,w = gray.shape[:2]; sw = w//n
    pad = max(sw//10, 2)
    return gray[:, max(0,i*sw-pad):min(w,(i+1)*sw+pad)]

def prep(gray):
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

def slot_changed(orig_small, curr_small, thresh=20):
    if orig_small is None: return False
    diff = cv2.absdiff(orig_small, curr_small)
    if len(diff.shape)==3: diff = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
    return (np.count_nonzero(diff>30)/max(diff.size,1))*100 > thresh


# ═══════════════════════════════════
# State
# ═══════════════════════════════════
running = False
region = None
session_keys = 0
session_start = 0
user_pressed = None  # key ที่ user กดล่าสุด
CFG = os.path.join(SCRIPT_DIR, "config.json")

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

# Listen keyboard (สำหรับ learning)
def _on_press(e):
    global user_pressed
    k = e.name.lower()
    if k in "qweasd" and running:
        user_pressed = k
kb.on_press(_on_press)


# ═══════════════════════════════════
# EEL Functions
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
    c.create_rectangle(sw//2-180,16,sw//2+180,58,fill="#000",outline="#00ffaa")
    c.create_text(sw//2,28,text="ลากครอบแถวตัวอักษร",fill="#00ffaa",font=("Segoe UI",12,"bold"))
    c.create_text(sw//2,48,text="ไม่รวม counter · ESC ยกเลิก",fill="#aaa",font=("Segoe UI",9))
    st={"sx":0,"sy":0,"r":None}
    def _p(e):
        st["sx"],st["sy"]=e.x,e.y
        if st["r"]: c.delete(st["r"])
        st["r"]=c.create_rectangle(e.x,e.y,e.x,e.y,outline="#00ffaa",width=2)
    def _d(e):
        c.coords(st["r"],st["sx"],st["sy"],e.x,e.y); c.delete("sz")
        c.create_text((st["sx"]+e.x)//2,min(st["sy"],e.y)-12,
            text=f"{abs(e.x-st['sx'])}x{abs(e.y-st['sy'])}",fill="#00ffaa",font=("Consolas",11,"bold"),tags="sz")
    def _r(e):
        global region
        x1,y1=min(st["sx"],e.x),min(st["sy"],e.y); x2,y2=max(st["sx"],e.x),max(st["sy"],e.y)
        if (x2-x1)>10 and (y2-y1)>10:
            region={"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}; save_cfg()
            eel.onRegionSet(f"{x2-x1}x{y2-y1} ({x1},{y1})")
            eel.onLog("Region set!","ok")
        root.destroy()
    c.bind("<ButtonPress-1>",_p); c.bind("<B1-Motion>",_d); c.bind("<ButtonRelease-1>",_r)
    root.bind("<Escape>",lambda e:root.destroy())
    root.after(50,root.focus_force); root.mainloop()

@eel.expose
def toggle(num_keys, scan_ms, key_delay_ms, lane_delay_ms):
    global running
    if running: running = False; return
    if not region: eel.onLog("กด F6 เลือก region ก่อน!","warn"); return
    running = True; save_cfg(num_keys=num_keys)
    eel.onStatus(True)
    has = learned_keys()
    if has:
        eel.onLog(f"มี template: {' '.join(k.upper() for k in has)} → auto!","ok")
    else:
        eel.onLog("ยังไม่มี template → เล่นเองรอบแรก กด QWEASD ตามเกม","warn")
    threading.Thread(target=main_loop, args=(num_keys,scan_ms/1000,key_delay_ms/1000,lane_delay_ms/1000), daemon=True).start()

@eel.expose
def test_btn(num_keys):
    if not region: eel.onLog("เลือก region ก่อน!","warn"); return
    frame = grab(region)
    if frame is None: eel.onLog("จับหน้าจอไม่ได้!","err"); return
    b = grab_b64(region)
    if b: eel.onPreview(b)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    for i in range(num_keys):
        roi = prep(get_slot(gray, i, num_keys))
        ch, cf = match_template(roi)
        eel.onLog(f"  Slot {i+1}: {ch.upper() if ch else '??'} ({cf:.0%})", "key" if ch else "warn")
    eel.onLog(f"Learned: {' '.join(k.upper() for k in learned_keys()) or 'none yet'}","dim")

@eel.expose
def clear_all():
    import shutil
    if os.path.isdir(TMPL_DIR): shutil.rmtree(TMPL_DIR)
    os.makedirs(TMPL_DIR, exist_ok=True)
    templates.clear()
    eel.onLog("ลบ template ทั้งหมดแล้ว","ok")


# ═══════════════════════════════════
# Main Loop
# ═══════════════════════════════════
def main_loop(num_keys, sd, kd, ld):
    """
    ง่ายที่สุด:
    - สแกนทุก slot จากซ้ายไปขวา
    - slot ที่จับภาพเปลี่ยนจากตอนเริ่ม = กดไปแล้ว → ข้าม
    - slot ที่ยังไม่เปลี่ยน = ต้องกด → อ่านตัวอักษร → กด
    - ถ้าอ่านไม่ออก → รอ user กดเอง → จำ
    """
    global running, session_keys, session_start, user_pressed
    session_keys = 0; session_start = time.time()
    user_pressed = None

    # จับ reference frame ตอนเริ่ม
    time.sleep(0.3)  # รอให้เกม render
    ref_frame = grab(region)
    while ref_frame is None and running:
        time.sleep(0.1); ref_frame = grab(region)
    if not running: eel.onStatus(False); return

    ref_slots = []  # ภาพ slot ตอนเริ่ม (ย่อ 32x32)
    h,w = ref_frame.shape[:2]; sw = w // num_keys
    for i in range(num_keys):
        s = ref_frame[:, i*sw:min((i+1)*sw, w)]
        ref_slots.append(cv2.resize(s, (32,32)))

    eel.onLog(f"จับภาพเริ่มต้นแล้ว ({num_keys} slots)", "ok")
    last_pressed_slot = -1
    round_count = 0

    while running:
        try:
            frame = grab(region)
            if frame is None: time.sleep(0.03); continue
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

            # แบ่ง slot ปัจจุบัน
            curr_slots = []
            for i in range(num_keys):
                s = frame[:, i*sw:min((i+1)*sw, w)]
                curr_slots.append(cv2.resize(s, (32,32)))

            # หา slot แรกที่ยังไม่เปลี่ยน = active
            active = -1
            for i in range(num_keys):
                if not slot_changed(ref_slots[i], curr_slots[i]):
                    active = i; break

            # Debug UI
            try: eel.onDebug(active, num_keys)
            except: pass

            # ═══ ทุก slot เปลี่ยน = จบ round ═══
            if active == -1:
                round_count += 1
                eel.onLog(f"Round {round_count} done!", "ok")

                # รอ + จับ reference ใหม่สำหรับ round ถัดไป
                time.sleep(ld)

                # รอจนหน้าจอเปลี่ยนเป็น round ใหม่
                old_small = cv2.resize(frame, (64,64))
                waited = 0
                while running and waited < 10:
                    time.sleep(0.1); waited += 0.1
                    new_frame = grab(region)
                    if new_frame is None: continue
                    new_small = cv2.resize(new_frame, (64,64))
                    diff = cv2.absdiff(old_small, new_small)
                    change = (np.count_nonzero(cv2.cvtColor(diff,cv2.COLOR_BGR2GRAY)>30)/diff[:,:,0].size)*100
                    if change > 15:
                        # หน้าจอเปลี่ยน! จับ reference ใหม่
                        time.sleep(0.2)  # รอให้ render เสร็จ
                        ref_frame = grab(region)
                        if ref_frame is not None:
                            h2,w2 = ref_frame.shape[:2]; sw = w2//num_keys
                            ref_slots = [cv2.resize(ref_frame[:,i*sw:min((i+1)*sw,w2)],(32,32)) for i in range(num_keys)]
                            eel.onLog("Round ใหม่!", "lane")
                        break
                last_pressed_slot = -1
                continue

            # ═══ มี slot active → อ่าน + กด ═══
            roi = prep(get_slot(gray, active, num_keys))
            ch, cf = match_template(roi)

            if ch and cf > 0.55:
                # อ่านออก → กดเลย!
                kb.press_and_release(ch)
                session_keys += 1
                spd = session_keys / max(time.time()-session_start, 0.1)
                eel.onKeyPressed(ch, session_keys, spd)
                eel.onLog(f"  [{session_keys}] {ch.upper()} (slot {active+1}, {cf:.0%})", "key")
                last_pressed_slot = active
                user_pressed = None
                time.sleep(kd)

            else:
                # อ่านไม่ออก → รอ user กดเอง
                if active != last_pressed_slot:
                    eel.onLog(f"  Slot {active+1}: ไม่รู้จัก → กดเองเลย!", "warn")
                    last_pressed_slot = active
                    user_pressed = None

                # User กดปุ่ม → เรียนรู้!
                if user_pressed and roi is not None:
                    save_template(user_pressed, roi)
                    eel.onLog(f"  จำได้แล้ว: {user_pressed.upper()} ✓", "ok")
                    eel.onLearned(learned_keys())
                    session_keys += 1
                    spd = session_keys / max(time.time()-session_start, 0.1)
                    eel.onKeyPressed(user_pressed, session_keys, spd)
                    last_pressed_slot = active
                    user_pressed = None
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
                b = grab_b64(region)
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
    web_dir = os.path.join(SCRIPT_DIR, "web")
    if not os.path.exists(os.path.join(web_dir, "index.html")):
        print(f"ERROR: {web_dir}/index.html not found"); sys.exit(1)

    eel.init(web_dir)
    lk = learned_keys()
    print(f"Templates: {sum(len(v) for v in templates.values())} | Learned: {', '.join(k.upper() for k in lk) or 'none'}")
    threading.Thread(target=bg, daemon=True).start()

    for mode in ["chrome-app","chrome","edge","default"]:
        try: eel.start("index.html", size=(480,700), port=0, mode=mode); break
        except: continue
    else: print("Need Chrome or Edge!")
