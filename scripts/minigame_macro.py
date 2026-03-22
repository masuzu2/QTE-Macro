"""
QTE Macro - Final Working Version
ใช้ pyautogui กดปุ่ม (เข้าเกมได้จริง)
เล่นเองรอบแรก → macro จำ → รอบถัดไป auto
"""

import time, sys, os, json, threading, base64, io
import eel

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ═══ Check libs ═══
MISSING = []
for lib, pkg in [("pyautogui","pyautogui"),("keyboard","keyboard"),("mss","mss"),
                  ("cv2","opencv-python"),("numpy","numpy"),("PIL","Pillow"),("eel","eel")]:
    try: __import__(lib)
    except ImportError: MISSING.append(pkg)
if MISSING:
    print(f"pip install {' '.join(MISSING)}"); sys.exit(1)

import pyautogui
import keyboard as kb
import mss
import cv2, numpy as np
from PIL import Image, ImageGrab

# pyautogui settings
pyautogui.PAUSE = 0
pyautogui.FAILSAFE = False

# ลอง pydirectinput (สำหรับ DirectX games)
try:
    import pydirectinput
    pydirectinput.PAUSE = 0
    HAS_DIRECTINPUT = True
except ImportError:
    HAS_DIRECTINPUT = False


# ═══════════════════════════════════
# Key Sender - ลองหลายวิธีกดปุ่ม
# ═══════════════════════════════════
def press_key(key):
    """กดปุ่ม - ลอง 3 วิธี จนกว่าจะได้"""
    key = key.lower()
    try:
        if HAS_DIRECTINPUT:
            pydirectinput.press(key)
        else:
            pyautogui.press(key)
    except:
        try:
            kb.press_and_release(key)
        except:
            pass


# ═══════════════════════════════════
# Capture
# ═══════════════════════════════════
sct = None
try: sct = mss.mss()
except: pass

def grab(r):
    try:
        if sct:
            return cv2.cvtColor(np.array(sct.grab(r)), cv2.COLOR_BGRA2BGR)
        return cv2.cvtColor(np.array(ImageGrab.grab(
            bbox=(r["left"],r["top"],r["left"]+r["width"],r["top"]+r["height"]))), cv2.COLOR_RGB2BGR)
    except: return None

def grab_b64(r):
    try:
        img = ImageGrab.grab(bbox=(r["left"],r["top"],r["left"]+r["width"],r["top"]+r["height"]))
        w,h = img.size
        if w > 460: img = img.resize((460,max(int(h*460/w),1)),Image.NEAREST)
        buf = io.BytesIO(); img.save(buf,format="JPEG",quality=50)
        return base64.b64encode(buf.getvalue()).decode()
    except: return None


# ═══════════════════════════════════
# Templates - จำ font จากเกม
# ═══════════════════════════════════
TMPL_DIR = os.path.join(SCRIPT_DIR, "game_templates")
os.makedirs(TMPL_DIR, exist_ok=True)
templates = {}
TH = 48

def load_templates():
    global templates; templates.clear()
    for f in os.listdir(TMPL_DIR):
        if not f.endswith(".png"): continue
        key = f.split("_")[0].split(".")[0].lower()
        if len(key)!=1: continue
        img = cv2.imread(os.path.join(TMPL_DIR,f), cv2.IMREAD_GRAYSCALE)
        if img is not None:
            h,w = img.shape[:2]
            templates.setdefault(key,[]).append(cv2.resize(img,(max(int(w*TH/h),3),TH)))

def save_tmpl(key, gray_img):
    key = key.lower()
    n = len([f for f in os.listdir(TMPL_DIR) if f.startswith(key)])
    fname = f"{key}.png" if n==0 else f"{key}_{n+1}.png"
    cv2.imwrite(os.path.join(TMPL_DIR,fname), gray_img)
    h,w = gray_img.shape[:2]
    templates.setdefault(key,[]).append(cv2.resize(gray_img,(max(int(w*TH/h),3),TH)))

def match(gray_roi, threshold=0.55):
    if not templates or gray_roi is None or gray_roi.size<20: return None,0
    h,w = gray_roi.shape[:2]
    if h<3 or w<3: return None,0
    roi = cv2.resize(gray_roi,(max(int(w*TH/h),3),TH))
    best_c, best_v = None, 0
    for ch, tmpls in templates.items():
        for t in tmpls:
            tw = t.shape[1]
            if tw > roi.shape[1]*1.5 or tw < roi.shape[1]*0.3: continue
            target = roi
            if tw > roi.shape[1] or t.shape[0] > roi.shape[0]:
                px,py = max(0,(tw-roi.shape[1])//2+3), max(0,(t.shape[0]-roi.shape[0])//2+3)
                target = cv2.copyMakeBorder(roi,py,py,px,px,cv2.BORDER_CONSTANT,value=0)
            try:
                res = cv2.matchTemplate(target,t,cv2.TM_CCOEFF_NORMED)
                _,mx,_,_ = cv2.minMaxLoc(res)
                if mx>threshold and mx>best_v:
                    best_c,best_v = ch,mx
                    if mx>0.85: return best_c,best_v
            except: continue
    return best_c, best_v

load_templates()


# ═══════════════════════════════════
# Image Processing
# ═══════════════════════════════════
def get_slot(gray, i, n):
    h,w = gray.shape[:2]; sw = w//n
    pad = max(sw//10,2)
    return gray[:, max(0,i*sw-pad):min(w,(i+1)*sw+pad)]

def prep(gray):
    """ทำให้ตัวอักษรเป็นขาวบนดำ + crop"""
    if gray is None or gray.size<20: return None
    h,w = gray.shape[:2]
    # ขยายถ้าเล็กเกิน
    if h<25 or w<15:
        s = max(30/max(h,1), 30/max(w,1), 1)
        gray = cv2.resize(gray, None, fx=s, fy=s, interpolation=cv2.INTER_CUBIC)

    # ลอง threshold หลายวิธี เลือกอันที่ได้ content ดีที่สุด
    best = None
    best_score = 0

    for method in ["otsu", "adaptive", "fixed_127", "fixed_180"]:
        try:
            if method == "otsu":
                _, t = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
            elif method == "adaptive":
                t = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,15,5)
            elif method == "fixed_127":
                _, t = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
            elif method == "fixed_180":
                _, t = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)

            ti = cv2.bitwise_not(t)
            # เลือกแบบที่ content น้อยกว่า (ตัวอักษร = ส่วนน้อย)
            r = t if np.count_nonzero(t) < np.count_nonzero(ti) else ti

            # Crop
            cs = np.argwhere(r > 127)
            if len(cs) < 5: continue
            y0,x0 = cs.min(0); y1,x1 = cs.max(0)
            c = r[max(0,y0-1):y1+2, max(0,x0-1):x1+2]
            if c.shape[0]<4 or c.shape[1]<4: continue

            # Score: ตัวอักษรควรกิน 10-60% ของพื้นที่
            ratio = np.count_nonzero(c>127) / max(c.size,1)
            if 0.08 < ratio < 0.65:
                score = c.size  # ยิ่งใหญ่ยิ่งดี
                if score > best_score:
                    best = c; best_score = score
        except: continue

    return best

def slot_changed(orig_small, curr_small, thresh=20):
    """เช็คว่า slot เปลี่ยนจากตอนเริ่ม"""
    if orig_small is None: return False
    diff = cv2.absdiff(orig_small, curr_small)
    if len(diff.shape)==3: diff = cv2.cvtColor(diff,cv2.COLOR_BGR2GRAY)
    return (np.count_nonzero(diff>30)/max(diff.size,1))*100 > thresh


# ═══════════════════════════════════
# State
# ═══════════════════════════════════
running = False
region = None
session_keys = 0
session_start = 0
user_pressed = None
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

# Listen user keyboard
def _on_press(e):
    global user_pressed
    k = e.name.lower()
    if k in "qweasd" and running:
        user_pressed = k
kb.on_press(_on_press)


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
    c = tk.Canvas(root,highlightthickness=0,cursor="cross"); c.pack(fill=tk.BOTH,expand=True)
    if photo: c.create_image(0,0,anchor="nw",image=photo)
    c.create_rectangle(0,0,sw,sh,fill="black",stipple="gray25")
    c.create_rectangle(sw//2-180,16,sw//2+180,58,fill="#000",outline="#00ffaa")
    c.create_text(sw//2,28,text="ลากครอบแถวตัวอักษร (ไม่รวม counter)",fill="#00ffaa",font=("Segoe UI",12,"bold"))
    c.create_text(sw//2,48,text="ESC ยกเลิก",fill="#aaa",font=("Segoe UI",9))
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
            eel.onLog(f"Region set!","ok")
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

    # บอกสถานะ
    method = "pydirectinput" if HAS_DIRECTINPUT else "pyautogui"
    eel.onLog(f"Key method: {method}","dim")
    lk = sorted(templates.keys())
    if lk:
        eel.onLog(f"มี template: {' '.join(k.upper() for k in lk)} → AUTO!","ok")
    else:
        eel.onLog("ยังไม่มี template → เล่นเองรอบแรก","warn")
        eel.onLog("กด Q W E A S D ตามเกม → macro จะจำเอง","warn")

    threading.Thread(target=main_loop, args=(num_keys,scan_ms/1000,key_delay_ms/1000,lane_delay_ms/1000), daemon=True).start()

@eel.expose
def test_btn(num_keys):
    if not region: eel.onLog("เลือก region ก่อน!","warn"); return
    frame = grab(region)
    if frame is None: eel.onLog("จับหน้าจอไม่ได้!","err"); return
    eel.onLog(f"Capture OK: {frame.shape[1]}x{frame.shape[0]}","ok")
    b = grab_b64(region)
    if b: eel.onPreview(b)
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    for i in range(num_keys):
        slot_gray = get_slot(gray, i, num_keys)
        eel.onLog(f"  Slot {i+1} raw: {slot_gray.shape[1]}x{slot_gray.shape[0]} px","dim")
        processed = prep(slot_gray)
        if processed is not None:
            eel.onLog(f"  Slot {i+1} prep: {processed.shape[1]}x{processed.shape[0]} px","dim")
            ch, cf = match(processed)
            eel.onLog(f"  Slot {i+1}: {ch.upper() if ch else '??'} ({cf:.0%})", "key" if ch else "warn")
        else:
            eel.onLog(f"  Slot {i+1}: prep failed (can't extract text)","err")
    eel.onLog(f"Learned: {' '.join(k.upper() for k in sorted(templates.keys())) or 'none'}","dim")

@eel.expose
def clear_all():
    import shutil
    if os.path.isdir(TMPL_DIR): shutil.rmtree(TMPL_DIR)
    os.makedirs(TMPL_DIR, exist_ok=True)
    templates.clear()
    eel.onLog("ลบ template ทั้งหมดแล้ว","ok")

@eel.expose
def test_press(key):
    """ทดสอบกดปุ่ม"""
    press_key(key)
    eel.onLog(f"Test press: {key.upper()}","ok")


# ═══════════════════════════════════
# Main Loop
# ═══════════════════════════════════
def main_loop(num_keys, sd, kd, ld):
    global running, session_keys, session_start, user_pressed
    session_keys = 0; session_start = time.time()
    user_pressed = None

    # จับ reference
    eel.onLog("กำลังจับภาพเริ่มต้น...","dim")
    time.sleep(0.5)
    ref_frame = grab(region)
    retry = 0
    while ref_frame is None and running and retry < 20:
        time.sleep(0.1); ref_frame = grab(region); retry += 1
    if ref_frame is None or not running:
        eel.onLog("จับหน้าจอไม่ได้! ลองรัน Admin","err")
        eel.onStatus(False); return

    h,w = ref_frame.shape[:2]; sw = w//num_keys
    ref_slots = [cv2.resize(ref_frame[:,i*sw:min((i+1)*sw,w)],(32,32)) for i in range(num_keys)]
    eel.onLog(f"OK! {num_keys} slots, {sw}px each","ok")

    last_active = -1
    round_num = 0

    while running:
        try:
            frame = grab(region)
            if frame is None: time.sleep(0.05); continue
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

            # แบ่ง slot ปัจจุบัน
            curr_slots = [cv2.resize(frame[:,i*sw:min((i+1)*sw,w)],(32,32)) for i in range(num_keys)]

            # หา slot active (ยังไม่เปลี่ยนจาก ref)
            active = -1
            for i in range(num_keys):
                if not slot_changed(ref_slots[i], curr_slots[i]):
                    active = i; break

            try: eel.onDebug(active, num_keys)
            except: pass

            # ═══ ทุก slot เปลี่ยน = จบ round ═══
            if active == -1:
                round_num += 1
                eel.onLog(f"Round {round_num} done!","ok")
                time.sleep(ld)

                # รอจนหน้าจอเปลี่ยนเป็น round ใหม่
                old_frame = cv2.resize(frame,(64,64))
                for _ in range(100):
                    if not running: break
                    time.sleep(0.1)
                    new_frame = grab(region)
                    if new_frame is None: continue
                    new_small = cv2.resize(new_frame,(64,64))
                    diff = cv2.absdiff(old_frame, new_small)
                    change = (np.count_nonzero(cv2.cvtColor(diff,cv2.COLOR_BGR2GRAY)>30)/diff[:,:,0].size)*100
                    if change > 15:
                        time.sleep(0.3)
                        ref_frame = grab(region)
                        if ref_frame is not None:
                            h2,w2 = ref_frame.shape[:2]; sw = w2//num_keys
                            ref_slots = [cv2.resize(ref_frame[:,i*sw:min((i+1)*sw,w2)],(32,32)) for i in range(num_keys)]
                            eel.onLog("Round ใหม่!","lane")
                        break
                last_active = -1
                continue

            # ═══ มี slot active ═══
            roi = prep(get_slot(gray, active, num_keys))
            ch, cf = match(roi)

            if ch and cf > 0.50:
                # ═══ รู้จัก → กดเลย! ═══
                press_key(ch)
                session_keys += 1
                spd = session_keys / max(time.time()-session_start, 0.1)
                eel.onKeyPressed(ch, session_keys, spd)
                eel.onLog(f"  [{session_keys}] {ch.upper()} (slot {active+1}, {cf:.0%})","key")
                last_active = active
                user_pressed = None
                time.sleep(kd)

            else:
                # ═══ ไม่รู้จัก → รอ user กด → จำ ═══
                if active != last_active:
                    eel.onLog(f"  Slot {active+1}: ไม่รู้จัก → กดเองเลย!","warn")
                    last_active = active
                    user_pressed = None

                if user_pressed and roi is not None:
                    save_tmpl(user_pressed, roi)
                    eel.onLog(f"  จำได้: {user_pressed.upper()} ✓ (total: {sum(len(v) for v in templates.values())})","ok")
                    try: eel.onLearned(sorted(templates.keys()))
                    except: pass
                    session_keys += 1
                    spd = session_keys / max(time.time()-session_start, 0.1)
                    eel.onKeyPressed(user_pressed, session_keys, spd)
                    last_active = active
                    user_pressed = None
                    time.sleep(kd)
                else:
                    time.sleep(sd)

        except Exception as e:
            eel.onLog(f"Error: {e}","err")
            time.sleep(0.2)

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
    lk = sorted(templates.keys())
    print(f"Templates: {sum(len(v) for v in templates.values())} | Keys: {', '.join(k.upper() for k in lk) or 'none'}")
    print(f"Key method: {'pydirectinput' if HAS_DIRECTINPUT else 'pyautogui'}")
    threading.Thread(target=bg, daemon=True).start()

    for mode in ["chrome-app","chrome","edge","default"]:
        try: eel.start("index.html", size=(480,700), port=0, mode=mode); break
        except: continue
    else: print("Need Chrome or Edge!")
