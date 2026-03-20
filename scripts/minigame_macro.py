"""
QTE Macro v5 - Web UI Edition
HTML + Tailwind CSS + eel bridge
"""

import time, sys, os, json, threading, base64, io
import eel

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# ── Lib check ──
MISSING = []
for lib, pkg in [("keyboard","keyboard"),("mss","mss"),("cv2","opencv-python"),
                  ("numpy","numpy"),("PIL","Pillow"),("eel","eel")]:
    try: __import__(lib)
    except ImportError: MISSING.append(pkg)
if MISSING:
    print(f"Missing: pip install {' '.join(MISSING)}")
    sys.exit(1)

import keyboard as kb
import mss
import cv2, numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageGrab

# ═══════════════════════════════════════════════════
# Screen Capture
# ═══════════════════════════════════════════════════
class Cap:
    def __init__(self):
        self.sct = None
        try: self.sct = mss.mss(); self.sct.grab(self.sct.monitors[0]); self.ok = True
        except: self.ok = False

    def grab(self, region):
        try:
            if self.sct:
                return cv2.cvtColor(np.array(self.sct.grab(region)), cv2.COLOR_BGRA2BGR)
            else:
                img = ImageGrab.grab(bbox=(region["left"],region["top"],
                    region["left"]+region["width"],region["top"]+region["height"]))
                return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        except: return None

    def grab_b64(self, region, max_w=480):
        """Grab as base64 JPEG (fast) for web preview"""
        try:
            if self.sct:
                s = self.sct.grab(region)
                img = Image.frombytes("RGB", s.size, s.bgra, "raw", "BGRX")
            else:
                img = ImageGrab.grab(bbox=(region["left"],region["top"],
                    region["left"]+region["width"],region["top"]+region["height"]))
            # Resize ให้เล็กลง → ส่งเร็วขึ้น
            w, h = img.size
            if w > max_w:
                ratio = max_w / w
                img = img.resize((max_w, max(int(h * ratio), 1)), Image.NEAREST)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=60)  # JPEG เร็วกว่า PNG 5x
            return base64.b64encode(buf.getvalue()).decode()
        except: return None


# ═══════════════════════════════════════════════════
# Character Engine - Auto templates + multi-pipeline
# ═══════════════════════════════════════════════════
class CharEngine:
    def __init__(self):
        self.templates = {}
        self.target_h = 48
        self.total = 0
        self._generate()

    def _generate(self):
        chars = "qweasdzxcrtyfghvbn1234567890"
        fonts = self._find_fonts()
        sizes = [40, 56, 72]

        for ch in chars:
            self.templates[ch] = []
            for fp in fonts:
                for sz in sizes:
                    for bold in [False, True]:
                        for case in [ch.upper(), ch.lower()]:
                            img = self._render(case, fp, sz, bold)
                            if img is not None:
                                h, w = img.shape[:2]
                                nw = max(int(w * (self.target_h / h)), 3)
                                self.templates[ch].append(cv2.resize(img, (nw, self.target_h)))
                                self.total += 1

    def _find_fonts(self):
        fonts = []
        names = ["arialbd.ttf","arial.ttf","calibrib.ttf","segoeuib.ttf","segoeui.ttf",
                 "consolab.ttf","consola.ttf","tahomabd.ttf","tahoma.ttf","verdanab.ttf",
                 "verdana.ttf","impact.ttf","trebucbd.ttf","courbd.ttf"]
        win_dir = os.environ.get("WINDIR", "C:\\Windows")
        local = os.environ.get("LOCALAPPDATA", "")
        paths = [os.path.join(win_dir,"Fonts")]
        if local: paths.append(os.path.join(local,"Microsoft","Windows","Fonts"))
        for n in names:
            for p in paths:
                fp = os.path.join(p, n)
                if os.path.exists(fp): fonts.append(fp); break
        return fonts[:8] if fonts else [None]

    def _render(self, char, font_path, size, bold=False):
        try:
            csz = size + 20
            img = Image.new("L", (csz, csz), 0)
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype(font_path, size) if font_path else ImageFont.load_default()
            bbox = draw.textbbox((0,0), char, font=font)
            tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
            x, y = (csz-tw)//2-bbox[0], (csz-th)//2-bbox[1]
            draw.text((x,y), char, fill=255, font=font)
            arr = np.array(img)
            if bold:
                arr = cv2.dilate(arr, cv2.getStructuringElement(cv2.MORPH_ELLIPSE,(3,3)), iterations=1)
            coords = np.argwhere(arr > 40)
            if len(coords) < 8: return None
            y0,x0 = coords.min(axis=0); y1,x1 = coords.max(axis=0)
            cr = arr[max(0,y0-1):y1+2, max(0,x0-1):x1+2]
            if cr.shape[0]<4 or cr.shape[1]<4: return None
            _, b = cv2.threshold(cr, 80, 255, cv2.THRESH_BINARY)
            return b
        except: return None

    def match_lane(self, gray, num_keys, threshold=0.50):
        prepped = self._multi_preprocess(gray)
        all_res = [self._match_slots(p, num_keys, threshold) for p in prepped]

        final = []
        for i in range(num_keys):
            from collections import Counter
            votes = Counter()
            bc = 0
            for res in all_res:
                if i < len(res) and res[i][0]:
                    votes[res[i][0]] += 1
                    bc = max(bc, res[i][1])
            if votes:
                final.append((votes.most_common(1)[0][0], bc))
            else:
                final.append((None, 0))
        return final

    def _match_slots(self, gray, num_keys, threshold):
        h, w = gray.shape[:2]
        sw = w // num_keys
        results = []
        for i in range(num_keys):
            x1 = i * sw; x2 = min(x1+sw, w)
            pad = max(sw//10, 2)
            roi = gray[:, max(0,x1-pad):min(w,x2+pad)]
            if roi.shape[1] < 5: results.append((None,0)); continue
            roi = self._tight_crop(roi)
            if roi is None: results.append((None,0)); continue
            results.append(self._match_one(roi, threshold))
        return results

    def _tight_crop(self, binary):
        coords = np.argwhere(binary > 127)
        if len(coords) < 8: return None
        y0,x0 = coords.min(axis=0); y1,x1 = coords.max(axis=0)
        c = binary[max(0,y0-2):y1+3, max(0,x0-2):x1+3]
        return c if c.shape[0]>3 and c.shape[1]>3 else None

    def _match_one(self, roi, threshold):
        rh, rw = roi.shape[:2]
        nw = max(int(rw*(self.target_h/rh)),3)
        roi_n = cv2.resize(roi, (nw, self.target_h))
        best_c, best_v = None, 0

        for char, tmpls in self.templates.items():
            for tmpl in tmpls:
                th, tw = tmpl.shape[:2]
                if tw > roi_n.shape[1]+4 or tw < roi_n.shape[1]*0.3: continue
                padded = roi_n
                if tw > roi_n.shape[1] or th > roi_n.shape[0]:
                    px = max(0,(tw-roi_n.shape[1])//2+2)
                    py = max(0,(th-roi_n.shape[0])//2+2)
                    padded = cv2.copyMakeBorder(roi_n,py,py,px,px,cv2.BORDER_CONSTANT,value=0)
                try:
                    res = cv2.matchTemplate(padded, tmpl, cv2.TM_CCOEFF_NORMED)
                    _,mx,_,_ = cv2.minMaxLoc(res)
                    if mx > threshold and mx > best_v:
                        best_c, best_v = char, mx
                        if mx > 0.80: return best_c, best_v
                except: continue
        return best_c, best_v

    def _multi_preprocess(self, gray):
        results = []
        if gray.shape[0] < 40:
            s = max(50/gray.shape[0], 1)
            gray = cv2.resize(gray, None, fx=s, fy=s, interpolation=cv2.INTER_CUBIC)

        # Otsu
        _, t1 = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
        t1i = cv2.bitwise_not(t1)
        results.append(t1 if np.count_nonzero(t1)<np.count_nonzero(t1i) else t1i)

        # Adaptive
        t2 = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,15,5)
        t2i = cv2.bitwise_not(t2)
        results.append(t2 if np.count_nonzero(t2)<np.count_nonzero(t2i) else t2i)

        # Contrast
        c = cv2.convertScaleAbs(gray, alpha=2.5, beta=-80)
        _, t3 = cv2.threshold(c, 127, 255, cv2.THRESH_BINARY)
        t3i = cv2.bitwise_not(t3)
        results.append(t3 if np.count_nonzero(t3)<np.count_nonzero(t3i) else t3i)

        # Edges
        edges = cv2.Canny(gray, 50, 150)
        results.append(cv2.dilate(edges, cv2.getStructuringElement(cv2.MORPH_RECT,(2,2)), iterations=1))

        return results


# ═══════════════════════════════════════════════════
# Lane Detector
# ═══════════════════════════════════════════════════
class LaneDetect:
    def __init__(self):
        self.last_hash = None
    def is_new(self, results):
        h = "".join(k or "?" for k,_ in results)
        return h != self.last_hash and "?" not in h
    def record(self, results):
        self.last_hash = "".join(k or "?" for k,_ in results)
    def reset(self):
        self.last_hash = None


# ═══════════════════════════════════════════════════
# Global State
# ═══════════════════════════════════════════════════
cap = Cap()
engine = CharEngine()
lane = LaneDetect()
running = False
region = None
session_keys = 0
session_start = 0
cfg_path = os.path.join(get_base_path(), "macro_config.json")

def load_cfg():
    global region
    try:
        with open(cfg_path,"r") as f:
            c = json.load(f)
            region = c.get("region")
            return c
    except: return {}

def save_cfg(**kw):
    c = load_cfg()
    c.update(kw)
    c["region"] = region
    try:
        with open(cfg_path,"w") as f: json.dump(c, f, indent=2)
    except: pass

load_cfg()


# ═══════════════════════════════════════════════════
# EEL Exposed Functions (JS calls these)
# ═══════════════════════════════════════════════════

@eel.expose
def select_region():
    """เปิด region picker (screenshot + drag)"""
    global region
    import tkinter as tk

    # Take screenshot
    try:
        screenshot = ImageGrab.grab()
    except:
        eel.onLog("Screenshot failed!", "err")
        return

    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{sw}x{sh}+0+0")

    photo = None
    try:
        from PIL import ImageTk as ITk
        photo = ITk.PhotoImage(screenshot)
    except: pass

    canvas = tk.Canvas(root, highlightthickness=0, cursor="cross")
    canvas.pack(fill=tk.BOTH, expand=True)

    if photo:
        canvas.create_image(0, 0, anchor="nw", image=photo)
    canvas.create_rectangle(0, 0, sw, sh, fill="black", stipple="gray25")

    # Guide text
    canvas.create_rectangle(sw//2-200, 14, sw//2+200, 70, fill="#000", outline="#00ffaa")
    canvas.create_text(sw//2, 30, text="ลากเลือกบริเวณที่ตัวอักษรขึ้น", fill="#00ffaa", font=("Segoe UI",13,"bold"))
    canvas.create_text(sw//2, 54, text="ลากให้ครอบทั้งแถว  |  ESC = ยกเลิก", fill="#ccc", font=("Segoe UI",10))

    pos_txt = canvas.create_text(sw//2, 86, text="", fill="#00ffaa", font=("Consolas", 10))

    state = {"sx":0, "sy":0, "rect":None}

    def on_motion(e):
        canvas.itemconfig(pos_txt, text=f"X:{e.x}  Y:{e.y}")
    def on_press(e):
        state["sx"], state["sy"] = e.x, e.y
        if state["rect"]: canvas.delete(state["rect"])
        state["rect"] = canvas.create_rectangle(e.x,e.y,e.x,e.y, outline="#00ffaa", width=2)
    def on_drag(e):
        canvas.coords(state["rect"], state["sx"], state["sy"], e.x, e.y)
        canvas.delete("sz")
        w,h = abs(e.x-state["sx"]), abs(e.y-state["sy"])
        canvas.create_text((state["sx"]+e.x)//2, min(state["sy"],e.y)-14,
            text=f"{w} x {h}", fill="#00ffaa", font=("Consolas",12,"bold"), tags="sz")
    def on_release(e):
        global region
        x1,y1 = min(state["sx"],e.x), min(state["sy"],e.y)
        x2,y2 = max(state["sx"],e.x), max(state["sy"],e.y)
        if (x2-x1)>10 and (y2-y1)>10:
            region = {"left":x1,"top":y1,"width":x2-x1,"height":y2-y1}
            save_cfg()
            eel.onRegionSet(f"{x2-x1}x{y2-y1}  ({x1},{y1})")
            eel.onLog(f"Region: {x2-x1}x{y2-y1} @ ({x1},{y1})", "ok")
        root.destroy()
    def on_escape(e):
        root.destroy()

    canvas.bind("<Motion>", on_motion)
    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)
    root.bind("<Escape>", on_escape)
    root.after(50, root.focus_force)
    root.mainloop()


@eel.expose
def test_capture(num_keys):
    if not region:
        eel.onLog("Select region first!", "warn"); return
    frame = cap.grab(region)
    if frame is None:
        eel.onLog("Capture FAILED!", "err"); return

    # Preview
    b64 = cap.grab_b64(region)
    if b64: eel.onPreview(b64)

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    results = engine.match_lane(gray, num_keys)
    seq = " ".join(f"{k.upper()}({c:.0%})" if k else "??" for k,c in results)
    eel.onLog(f"Test: {seq}", "key")


@eel.expose
def toggle_macro(num_keys, scan_ms, key_delay_ms, lane_delay_ms):
    global running
    if running:
        running = False
    else:
        if not region:
            eel.onLog("Select region first! (F6)", "warn"); return
        save_cfg(num_keys=num_keys, scan_ms=scan_ms, key_delay=key_delay_ms, lane_delay=lane_delay_ms)
        running = True
        threading.Thread(target=macro_loop, args=(num_keys, scan_ms, key_delay_ms, lane_delay_ms), daemon=True).start()

    eel.onStatusChange(running)
    eel.onLog("Macro STARTED" if running else "Macro STOPPED", "ok")


@eel.expose
def set_on_top(on):
    # eel ไม่มี native always-on-top แต่ log ไว้
    pass


def macro_loop(num_keys, scan_ms, key_delay_ms, lane_delay_ms):
    global running, session_keys, session_start
    session_keys = 0
    session_start = time.time()
    lane.reset()
    kd = key_delay_ms / 1000.0
    ld = lane_delay_ms / 1000.0

    while running:
        try:
            frame = cap.grab(region)
            if frame is None: time.sleep(0.05); continue

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            results = engine.match_lane(gray, num_keys)

            if not all(k for k,_ in results):
                time.sleep(scan_ms / 1000.0); continue

            if not lane.is_new(results):
                time.sleep(scan_ms / 1000.0); continue

            # New lane!
            lane.record(results)
            seq = " ".join(k.upper() for k,_ in results)
            eel.onLaneDetected(seq)
            eel.onLog(f"Lane: {seq}", "lane")

            for key, conf in results:
                if not running: break
                if key:
                    kb.press_and_release(key)
                    session_keys += 1
                    elapsed = time.time() - session_start
                    speed = session_keys / elapsed if elapsed > 0 else 0
                    eel.onKeyPressed(key, session_keys, speed)
                    time.sleep(kd)

            time.sleep(ld)

        except Exception as e:
            eel.onLog(str(e), "err")
            time.sleep(0.1)

    eel.onStatusChange(False)


# Preview thread - fast, separate from timer
def preview_loop():
    """Preview จอ realtime ทุก 300ms ใช้ JPEG เล็ก ไม่ค้าง"""
    prev_cap = Cap()  # ใช้ capture แยก ไม่ชนกับ macro thread
    while True:
        try:
            if region:
                b64 = prev_cap.grab_b64(region)
                if b64:
                    eel.onPreview(b64)
        except: pass
        time.sleep(0.3)

# Timer thread
def timer_loop():
    while True:
        try:
            if running and session_start:
                e = int(time.time() - session_start)
                m, s = divmod(e, 60)
                eel.onTimerUpdate(f"{m:02d}:{s:02d}")
            else:
                eel.onTimerUpdate("00:00")
        except: pass
        time.sleep(0.5)


# ═══════════════════════════════════════════════════
# Start
# ═══════════════════════════════════════════════════
if __name__ == "__main__":
    web_dir = os.path.join(get_base_path(), "web")
    eel.init(web_dir)

    print(f"Templates: {engine.total}")
    print(f"Capture: {'OK' if cap.ok else 'FAIL'}")
    print(f"Region: {region}")

    # Start background threads
    threading.Thread(target=preview_loop, daemon=True).start()
    threading.Thread(target=timer_loop, daemon=True).start()

    # Start eel (opens browser window)
    try:
        eel.start("index.html", size=(520, 780), port=0,
                  mode="chrome-app" if os.name == "nt" else "chrome")
    except:
        try:
            eel.start("index.html", size=(520, 780), port=0, mode="edge")
        except:
            try:
                eel.start("index.html", size=(520, 780), port=0, mode="default")
            except:
                print("Cannot open browser! Install Chrome or Edge.")
