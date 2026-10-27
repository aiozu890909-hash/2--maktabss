#!/usr/bin/env python3
"""
SmartSchool Attendance — kamera + yuzni tanish.
O'rnatish:  pip install opencv-python face_recognition numpy requests
Ishga tushirish (Windows PowerShell misoli):
    $env:SS_API_URL="https://script.google.com/macros/s/XXXX/exec"; $env:SS_DEVICE_KEY="..."; python camera.py
"""
import os, sys, time, json, uuid, base64, pickle, sqlite3, threading, logging, socket
from collections import defaultdict
import cv2, numpy as np, requests, face_recognition

# ====== SHU IKKITA QIYMATNI TEKSHIRING ======
API_URL = os.getenv("SS_API_URL", "https://script.google.com/macros/s/AKfycbx1r3Fwj9upH77waJT5Vr-O79vXCw2EQDk96Myxif9TX7snplyUbQHxsvjS5Jhlh131/exec")
DEVICE_KEY = os.getenv("SS_DEVICE_KEY", "BU_YERGA_DEVICE_KEY_NI_YOZING")   # setup() Execution log'ga chiqargan kalit
_cam = os.getenv("SS_CAMERA", "0")
CAM = int(_cam) if _cam.isdigit() else _cam      # 0 yoki RTSP/URL
TOL = float(os.getenv("SS_TOL", "0.48"))         # kichikroq = qattiqroq
MARGIN = 0.04                                    # 1-chi va 2-chi eng yaqin odam orasidagi minimal farq
CONFIRM = 3                                      # ketma-ket nechta kadrda tasdiqlanadi
COOLDOWN = 120                                   # bir o'quvchiga qayta yuborish oralig'i (soniya)
UNK_COOLDOWN = 30
LIVENESS = os.getenv("SS_LIVENESS", "1") == "1"  # ko'z qisish (blink) tekshiruvi
SHOW = os.getenv("SS_SHOW", "1") == "1"
SYNC_EVERY = 300

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("ss")
KNOWN = {"ids": [], "enc": np.empty((0, 128)), "info": {}}
STAT = {"fps": 0.0, "cam": "offline"}
os.makedirs("unknown", exist_ok=True)


def post(payload):
    r = requests.post(API_URL, data=json.dumps({**payload, "key": DEVICE_KEY}),
                      headers={"Content-Type": "text/plain"}, timeout=40)
    j = r.json()
    if not j.get("ok"):
        raise RuntimeError(j.get("error"))
    return j


def db():
    c = sqlite3.connect("outbox.db", timeout=30)
    c.execute("create table if not exists outbox(id text primary key, payload text)")
    return c


def sender():
    """Offline navbat: aloqa bo'lmasa saqlanadi, tiklangach yuboriladi (server eventId bo'yicha takrorni o'tkazib yuboradi)."""
    while True:
        try:
            con = db()
            for eid, p in con.execute("select id,payload from outbox order by rowid limit 20").fetchall():
                post(json.loads(p))
                con.execute("delete from outbox where id=?", (eid,)); con.commit()
            con.close()
        except Exception as e:
            log.warning("sinxronlash kutilmoqda: %s", e)
        time.sleep(5)


def heartbeat():
    while True:
        try:
            post({"action": "heartbeat", "fps": round(STAT["fps"], 1), "cam": STAT["cam"], "host": socket.gethostname()})
        except Exception as e:
            log.warning("heartbeat: %s", e)
        time.sleep(30)


def sync_faces():
    cache = {}
    if os.path.exists("faces_cache.pkl"):
        try: cache = pickle.load(open("faces_cache.pkl", "rb"))
        except Exception: cache = {}
    new, fresh = {}, 0
    studs = post({"action": "faces"})["students"]
    for s in studs:
        c = cache.get(s["id"])
        if c and c["photoId"] == s["photoId"]:
            new[s["id"]] = c; continue
        raw = base64.b64decode(post({"action": "photo", "photoId": s["photoId"]})["data"]); fresh += 1
        img = cv2.cvtColor(cv2.imdecode(np.frombuffer(raw, np.uint8), cv2.IMREAD_COLOR), cv2.COLOR_BGR2RGB)
        encs = face_recognition.face_encodings(img)
        if len(encs) != 1:
            log.warning("%s: rasmda %d ta yuz — o'tkazib yuborildi", s["id"], len(encs)); continue
        new[s["id"]] = {"photoId": s["photoId"], "enc": encs[0], "name": f'{s["ism"]} {s["familiya"]}'}
    pickle.dump(new, open("faces_cache.pkl", "wb"))
    ids = list(new)
    KNOWN.update(ids=ids, enc=np.array([new[i]["enc"] for i in ids]) if ids else np.empty((0, 128)), info=new)
    log.info("Sheets: %d faol o'quvchi | Drive'dan yangi rasm: %d | tanishga tayyor: %d", len(studs), fresh, len(ids))


def syncer():
    while True:
        time.sleep(SYNC_EVERY)
        try: sync_faces()
        except Exception as e: log.warning("yuz bazasi: %s", e)


def match(enc):
    if not KNOWN["ids"]:
        return None
    d = face_recognition.face_distance(KNOWN["enc"], enc)
    o = np.argsort(d)
    if d[o[0]] > TOL: return None
    if len(o) > 1 and d[o[1]] - d[o[0]] < MARGIN: return None   # ikkilanish — taxmin qilmaymiz
    return KNOWN["ids"][o[0]]


def ear(p):
    p = np.array(p, float)
    return (np.linalg.norm(p[1] - p[5]) + np.linalg.norm(p[2] - p[4])) / (2 * np.linalg.norm(p[0] - p[3]))


def emit(sid, frame):
    ev = {"action": "logEvent", "eventId": str(uuid.uuid4()), "studentId": sid, "ts": int(time.time() * 1000)}
    con = db(); con.execute("insert into outbox values(?,?)", (ev["eventId"], json.dumps(ev))); con.commit(); con.close()
    if sid is None:
        cv2.imwrite(f"unknown/{ev['eventId']}.jpg", frame)
    log.info("hodisa: %s", KNOWN["info"].get(sid, {}).get("name", "NOMA'LUM" if sid is None else sid))


def main():
    if not API_URL or DEVICE_KEY.startswith("BU_YERGA"):
        raise SystemExit("camera.py boshidagi DEVICE_KEY ni yozing (yoki SS_DEVICE_KEY muhit o'zgaruvchisini bering)")
    try:
        sync_faces()   # boshlanishda Sheets va Drive'dan ma'lumotni darhol olamiz
    except Exception as e:
        log.error("Sheets/Drive'ga ulanib bo'lmadi: %s (kalit, havola va internetni tekshiring)", e)
        if "--test" in sys.argv: raise SystemExit(1)
    if "--test" in sys.argv:
        raise SystemExit("Test tugadi.")
    for f in (sender, heartbeat, syncer):
        threading.Thread(target=f, daemon=True).start()
    track = defaultdict(lambda: {"n": 0, "last": 0.0, "blink": 0.0, "closed": False})
    sent = {}
    cap, t0, frames = None, time.time(), 0
    while True:
        if cap is None:
            cap = cv2.VideoCapture(CAM)
            if not cap.isOpened():
                cap = None; STAT["cam"] = "offline"; log.error("kamera ochilmadi, qayta urinish..."); time.sleep(3); continue
            STAT["cam"] = "online"; log.info("kamera ulandi")
        ok, frame = cap.read()
        if not ok:
            STAT["cam"] = "offline"; log.error("kadr olinmadi, qayta ulanish"); cap.release(); cap = None; time.sleep(2); continue
        frames += 1
        if time.time() - t0 >= 5:
            STAT["fps"] = frames / (time.time() - t0); frames, t0 = 0, time.time()
        small = cv2.resize(frame, None, fx=0.5, fy=0.5)
        rgb = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
        locs = face_recognition.face_locations(rgb)
        encs = face_recognition.face_encodings(rgb, locs)
        marks = face_recognition.face_landmarks(rgb, locs) if LIVENESS else [None] * len(locs)
        now = time.time()
        for loc, enc, lm in zip(locs, encs, marks):
            sid = match(enc); key = sid or "unknown"; t = track[key]
            if now - t["last"] > 1.5: t.update(n=0, blink=0.0, closed=False)
            t["n"] += 1; t["last"] = now
            if lm:
                e = (ear(lm["left_eye"]) + ear(lm["right_eye"])) / 2
                if e < 0.20: t["closed"] = True
                elif t["closed"] and e > 0.25: t["blink"] = now; t["closed"] = False
            if t["n"] >= CONFIRM:
                if sid is None:
                    if now - sent.get("unknown", 0) > UNK_COOLDOWN:
                        emit(None, frame); sent["unknown"] = now
                elif now - sent.get(sid, 0) > COOLDOWN and (not LIVENESS or now - t["blink"] < 10):
                    emit(sid, frame); sent[sid] = now; t["blink"] = 0.0
            if SHOW:
                y1, x2, y2, x1 = [v * 2 for v in loc]
                col = (16, 185, 129) if sid else (68, 68, 239)
                cv2.rectangle(frame, (x1, y1), (x2, y2), col, 2)
                cv2.putText(frame, KNOWN["info"][sid]["name"] if sid else "Noma'lum", (x1, y1 - 8), cv2.FONT_HERSHEY_SIMPLEX, 0.6, col, 2)
        if SHOW:
            cv2.imshow("SmartSchool", frame)
            if cv2.waitKey(1) & 0xFF == ord("q"): break
    if cap: cap.release()
    cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
