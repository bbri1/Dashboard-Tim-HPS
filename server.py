"""Server lokal Dashboard HPS untuk jaringan kantor (tanpa internet).

Fitur:
  - menyajikan dashboard (index.html + data.js) ke seluruh PC di jaringan kantor
  - POST /api/upload   : terima file Excel (.xlsx) lalu bangun ulang data.js
  - POST /api/refresh  : bangun ulang data.js dari Excel yang ada di folder
  - GET  /api/status   : status server + info data terakhir
  - memantau folder: begitu ada file .xlsx yang berubah, data.js otomatis dibangun ulang

Jalankan:  python server.py [port]        (default 8080)
Token (opsional, untuk Proteksi upload):
  setx HPS_TOKEN "rahasia"  lalu set service-nya memakai env yang sama
"""
import sys, os, json, time, shutil, socket, threading, hashlib, datetime as dt
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("HPS_DATA_DIR", HERE)   # folder data.js + file Excel
sys.path.insert(0, HERE)
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("HPS_PORT", 8080))
TOKEN = os.environ.get("HPS_TOKEN", "")
CANON = "Progress_Report_HPS.xlsx"
LOG = os.path.join(HERE, "hps_server.log")
MAX_UPLOAD = 60 * 1024 * 1024          # 60 MB
_lock = threading.Lock()               # hanya satu proses build pada satu waktu
_state = {"generated": "", "rows": 0, "source": "", "last_error": "", "builds": 0, "started": dt.datetime.now().strftime("%Y-%m-%d %H:%M")}


def log(msg):
    line = "%s  %s" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line, flush=True)
    try:
        with open(LOG, "a", encoding="utf8") as f:
            f.write(line + "\n")
        if os.path.getsize(LOG) > 2 * 1024 * 1024:      # pangkas log yang membengkak
            with open(LOG, "r", encoding="utf8") as f:
                f.readlines()[500:]
            with open(LOG, "w", encoding="utf8") as f:
                f.writelines(f.readlines()[-500:])
    except OSError:
        pass


def read_meta():
    try:
        with open(os.path.join(DATA_DIR, "data.meta.json"), encoding="utf8") as f:
            return json.load(f)
    except Exception:
        return {}


STATE_FILE = os.path.join(DATA_DIR, "hps_state.json")


def load_state():
    """Tampilan tersimpan + jejak audit (dipakai bersama seluruh tim)."""
    try:
        with open(STATE_FILE, encoding="utf8") as f:
            s = json.load(f)
    except Exception:
        s = {}
    s.setdefault("views", [])
    s.setdefault("activity", [])
    return s


def save_state(s):
    try:
        with open(STATE_FILE, "w", encoding="utf8") as f:
            json.dump(s, f, ensure_ascii=False, indent=1)
    except OSError as e:
        log("gagal menyimpan state: %s" % e)


def add_activity(action, rows=0, source=""):
    s = load_state()
    s["activity"].append({
        "time": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "ip": _client_ip, "action": action, "rows": rows, "source": source,
    })
    s["activity"] = s["activity"][-30:]
    save_state(s)


_client_ip = "-"


def regenerate(src=None, reason="manual"):
    """Bangun ulang data.js dari Excel. Kembalikan (ok, info)."""
    with _lock:
        try:
            import update_data
            meta = update_data.generate(src)
            _state.update(generated=meta.get("generated", ""), rows=meta.get("rows", 0),
                          source=meta.get("source", ""), last_error="", builds=_state["builds"] + 1)
            log("data.js diperbarui (%s) -> %s baris, sumber: %s" % (reason, meta.get("rows"), meta.get("source")))
            return True, meta
        except Exception as e:
            _state["last_error"] = "%s: %s" % (type(e).__name__, e)
            log("GAGAL membangun data.js (%s) -> %s" % (reason, _state["last_error"]))
            return False, {"error": _state["last_error"]}


def file_sig(path):
    """Tanda file: (ukuran, waktu modif, hash isi) - hash supaya perubahan
    terdeteksi walau tool sync mempertahankan waktu modif."""
    try:
        st = os.stat(path)
        h = hashlib.md5()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return (st.st_size, int(st.st_mtime), h.hexdigest())
    except OSError:
        return None


def xlsx_signature(folder=None):
    """Tanda perubahan file Excel di folder (nama + tanda file)."""
    folder = folder or DATA_DIR
    sig = []
    if os.path.isfile(folder):                  # sumber berupa satu file .xlsx (hps_source.txt)
        s = file_sig(folder)
        return ((os.path.basename(folder),) + s,) if s else ()
    try:
        files = os.listdir(folder)
    except OSError:
        return ()
    for f in sorted(files):
        if f.lower().endswith(".xlsx") and not f.startswith("~$") and not f.startswith("."):
            p = os.path.join(folder, f)
            s = file_sig(p)
            if s:
                sig.append((f,) + s)
    return tuple(sig)


# ---------------- sumber data realtime (OneDrive / SharePoint) ----------------
SOURCE_FILE = os.path.join(HERE, "hps_source.txt")


def inbox_dir():
    """Folder tempat file Excel realtime (mis. hasil sync OneDrive).
    Diatur lewat file hps_source.txt (isi: path folder atau path file .xlsx)
    atau environment variable HPS_SOURCE."""
    p = os.environ.get("HPS_SOURCE", "").strip()
    if not p:
        try:
            with open(SOURCE_FILE, encoding="utf-8-sig") as f:
                lines = [l.strip() for l in f.read().splitlines()]
            p = next((l for l in lines if l and not l.startswith("#")), "")
        except OSError:
            p = ""
    if not p:
        return DATA_DIR
    p = os.path.expandvars(os.path.expanduser(p.strip('"')))
    return p if os.path.exists(p) else DATA_DIR


INBOX = inbox_dir()


def inbox_files():
    """Daftar file .xlsx di folder sumber (bisa folder atau satu file)."""
    if os.path.isfile(INBOX):
        return [INBOX] if INBOX.lower().endswith(".xlsx") else []
    out = []
    try:
        for f in os.listdir(INBOX):
            if f.lower().endswith(".xlsx") and not f.startswith("~$") and not f.startswith("."):
                out.append(os.path.join(INBOX, f))
    except OSError:
        pass
    return out


def stable_signature(path, wait=5, limit=40):
    """Tanda file bila ukurannya, waktu modif & hash isinya sudah tidak berubah
    selama `wait` detik (penting untuk OneDrive yg sedang sync / Excel yg sedang disimpan)."""
    t0, last = time.time(), None
    while time.time() - t0 < limit:
        cur = file_sig(path)
        if cur is None:
            return None
        if cur != last:
            last, t0 = cur, time.time()
        elif time.time() - t0 >= wait:
            return cur
        time.sleep(1)
    return last


def pull_from_inbox():
    """Salin file Excel terbaru dari folder sumber ke folder data, lalu bangun ulang."""
    files = inbox_files()
    if not files:
        return None
    newest = max(files, key=os.path.getmtime)
    sig1 = stable_signature(newest, 5)
    if sig1 is None:
        return None
    time.sleep(1.5)
    if stable_signature(newest, 3) != sig1:
        log("file sumber masih berubah (sync?), menunggu...")
        return None
    dest = os.path.join(DATA_DIR, CANON)
    if os.path.abspath(newest) != os.path.abspath(dest):
        try:
            shutil.copy2(newest, dest)
            log("Excel diambil dari sumber: %s" % newest)
        except OSError as e:
            log("gagal menyalin dari sumber: %s" % e)
            return None
    ok, info = regenerate(dest, "sumber realtime")
    if ok:
        add_activity("sync sumber realtime", info.get("rows", 0), os.path.basename(newest))
    return info if ok else None


def watcher():
    """Kalau ada Excel yang diganti (share jaringan / OneDrive), rebuild otomatis."""
    last = xlsx_signature(INBOX)
    while True:
        time.sleep(15)
        try:
            cur = xlsx_signature(INBOX)
            if cur != last:
                last = cur
                if os.path.abspath(INBOX) != os.path.abspath(DATA_DIR):
                    pull_from_inbox()
                else:
                    newest = max((os.path.join(DATA_DIR, f) for f, *_ in cur), key=os.path.getmtime) if cur else None
                    ok, info = regenerate(newest, "perubahan file Excel")
                    if ok:
                        add_activity("file Excel diganti", info.get("rows", 0), os.path.basename(newest or ""))
        except Exception as e:
            log("watcher error: %s" % e)


_BLOCK_NAMES = {"graph_config.json", "hps_state.json", "hps_source.txt", "server.py", "graph_sync.py",
                "update_data.py", "publish_data.py", "dockerfile", "docker-compose.yml"}
_BLOCK_EXT = (".py", ".log", ".bat", ".ps1", ".service", ".example", ".pyc")


def _blocked(path):
    """File rahasia/kode tidak boleh diunduh lewat web (graph_config.json berisi client_secret)."""
    from urllib.parse import unquote
    parts = [x for x in unquote(path).replace("\\", "/").split("/") if x]
    if any(x.startswith(".") for x in parts) or "__pycache__" in parts:   # .git, .github, .kilo, dll
        return True
    name = parts[-1].lower() if parts else ""
    return name in _BLOCK_NAMES or name.endswith(_BLOCK_EXT)


class Handler(SimpleHTTPRequestHandler):
    server_version = "HPS/1.0"

    def __init__(self, *a, **kw):
        super().__init__(*a, directory=HERE, **kw)

    def translate_path(self, path):
        p = super().translate_path(path)
        base = os.path.basename(p)
        if DATA_DIR != HERE and base in ("data.js", "data.meta.json"):
            cand = os.path.join(DATA_DIR, base)
            if os.path.exists(cand):
                return cand
        return p

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, fmt, *a):        # matikan log request per file
        pass

    # ---------- util ----------
    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _token_ok(self, q):
        if not TOKEN:
            return True
        got = self.headers.get("X-Token") or (q.get("token", [None])[0])
        return got == TOKEN

    def _need_token(self, q):
        if self._token_ok(q):
            return False
        self._json({"ok": False, "error": "token salah / tidak diisi"}, 401)
        return True

    # ---------- API ----------
    def do_GET(self):
        global _client_ip
        _client_ip = self.client_address[0]
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if _blocked(u.path):
            self._json({"ok": False, "error": "tidak ditemukan"}, 404)
            return
        if u.path == "/api/status":
            meta = read_meta()
            st = load_state()
            self._json({"ok": True, "server": "HPS LAN", "started": _state["started"],
                        "generated": meta.get("generated", _state.get("generated", "")),
                        "rows": meta.get("rows", 0), "source": meta.get("source", _state.get("source", "")),
                        "builds": _state["builds"], "last_error": _state["last_error"],
                        "token_required": bool(TOKEN), "folder": DATA_DIR,
                        "source_path": INBOX,
                        "source_kind": ("file" if os.path.isfile(INBOX) else
                                        ("folder" if os.path.abspath(INBOX) != os.path.abspath(DATA_DIR) else "unggahan manual")),
                        "last_activity": (st["activity"][-1] if st["activity"] else None),
                        "views": len(st["views"])})
            return
        if u.path == "/api/views":
            self._json({"ok": True, "views": load_state()["views"]})
            return
        if u.path == "/api/health":
            self._json({"ok": True})
            return
        return super().do_GET()

    def do_POST(self):
        global _client_ip
        _client_ip = self.client_address[0]
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if u.path == "/api/refresh":
            if self._need_token(q):
                return
            if os.path.abspath(INBOX) != os.path.abspath(DATA_DIR) and inbox_files():
                info = pull_from_inbox()                  # ambil Excel terbaru dari folder sumber (OneDrive dll)
                ok = info is not None
                info = info or {"error": "file sumber belum stabil / tidak terbaca"}
            else:
                ok, info = regenerate(None, "refresh manual")
                if ok: add_activity("refresh", info.get("rows", 0), info.get("source", ""))
            self._json({"ok": ok, "info": info})
            return
        if u.path == "/api/views":
            if self._need_token(q):
                return
            return self._save_view()
        if u.path == "/api/upload":
            if self._need_token(q):
                return
            return self._upload(q)
        self._json({"ok": False, "error": "endpoint tidak dikenal"}, 404)

    def _body_json(self):
        try:
            n = int(self.headers.get("Content-Length") or 0)
            if n <= 0 or n > 100000:
                return {}
            return json.loads(self.rfile.read(n).decode("utf8"))
        except Exception:
            return {}

    def _save_view(self):
        """Simpan / hapus 'tampilan' (filter) supaya bisa dipakai bersama rekan."""
        body = self._body_json()
        s = load_state()
        if body.get("delete"):
            s["views"] = [v for v in s["views"] if v.get("id") != body["delete"]]
            save_state(s)
            return self._json({"ok": True, "views": s["views"]})
        name = str(body.get("name") or "").strip()[:40]
        h = str(body.get("hash") or "")[:400]
        if not name:
            return self._json({"ok": False, "error": "nama tampilan kosong"}, 400)
        if not h:
            return self._json({"ok": False, "error": "filter kosong"}, 400)
        s["views"] = [v for v in s["views"] if v.get("hash") != h]
        s["views"].append({"id": dt.datetime.now().strftime("%Y%m%d%H%M%S%f"), "name": name,
                           "hash": h, "by": _client_ip,
                           "created": dt.datetime.now().strftime("%Y-%m-%d %H:%M")})
        s["views"] = s["views"][-40:]
        save_state(s)
        log("tampilan tersimpan: %s" % name)
        self._json({"ok": True, "views": s["views"]})

    def _upload(self, q):
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        if length <= 0:
            return self._json({"ok": False, "error": "file kosong"}, 400)
        if length > MAX_UPLOAD:
            return self._json({"ok": False, "error": "file terlalu besar (maks 60 MB)"}, 413)
        name = self.headers.get("X-Filename") or (q.get("filename", [""])[0]) or "upload.xlsx"
        name = os.path.basename(name.replace("\\", "/")) or "upload.xlsx"
        if not name.lower().endswith(".xlsx"):
            return self._json({"ok": False, "error": "hanya file .xlsx yang diterima"}, 400)
        tmp = os.path.join(DATA_DIR, ".upload_tmp.xlsx")
        try:
            with open(tmp, "wb") as f:
                left = length
                while left > 0:
                    chunk = self.rfile.read(min(1 << 20, left))
                    if not chunk:
                        break
                    f.write(chunk)
                    left -= len(chunk)
        except Exception as e:
            return self._json({"ok": False, "error": "gagal menyimpan file: %s" % e}, 500)
        with open(tmp, "rb") as f:
            magic = f.read(4)
        if magic[:2] != b"PK":                      # file .xlsx itu zip
            os.remove(tmp)
            return self._json({"ok": False, "error": "file bukan Excel yang valid (.xlsx)"}, 400)
        dest = os.path.join(DATA_DIR, CANON)
        try:
            os.replace(tmp, dest)
        except OSError as e:
            os.remove(tmp)
            return self._json({"ok": False, "error": "gagal menyimpan: %s" % e}, 500)
        log("Excel baru diunggah: %s (%d byte)" % (name, length))
        ok, info = regenerate(dest, "upload")
        if ok:
            add_activity("upload Excel", info.get("rows", 0), name)
        return self._json({"ok": ok, "info": info, "file": CANON})


def lan_ips():
    ips = set()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ips.add(s.getsockname()[0])
        s.close()
    except OSError:
        pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127."):
                ips.add(ip)
    except OSError:
        pass
    return sorted(ips) or ["127.0.0.1"]


def bootstrap():
    """Pertama kali jalan (mis. Docker dengan folder data baru): belum ada data.js tapi sudah ada Excel -> bangun sekarang."""
    try:
        if os.path.exists(os.path.join(DATA_DIR, "data.js")):
            return
        if os.path.abspath(INBOX) != os.path.abspath(DATA_DIR) and inbox_files():
            pull_from_inbox()
        elif xlsx_signature(DATA_DIR):
            regenerate(None, "start awal")
    except Exception as e:
        log("bootstrap error: %s" % e)


if __name__ == "__main__":
    threading.Thread(target=bootstrap, daemon=True).start()
    threading.Thread(target=watcher, daemon=True).start()
    try:
        import update_data  # noqa: F401  (cek openpyxl ada/tidak)
        import openpyxl     # noqa: F401
        log("openpyxl siap")
    except Exception as e:
        log("PERINGATAN: openpyxl belum ada -> %s (jalankan: pip install openpyxl)" % e)
    print("=" * 62)
    print("  DASHBOARD HPS - SERVER LOKAL (jaringan kantor)")
    print("  Di komputer ini : http://127.0.0.1:%d/" % PORT)
    for ip in lan_ips():
        print("  Di jaringan      : http://%s:%d/" % (ip, PORT))
    print("  Unggah Excel    : http://%s:%d/upload.html" % (lan_ips()[0], PORT))
    if os.path.abspath(INBOX) != os.path.abspath(DATA_DIR):
        print("  Sumber realtime : %s" % INBOX)
        print("                   (file .xlsx yg berubah di folder ini otomatis rebuild)")
    if TOKEN:
        print("  Token upload     : aktif (wajib diisi di halaman unggah)")
    print("  (Tutup jendela / stop service untuk menghentikan server)")
    print("=" * 62, flush=True)
    log("server berjalan di port %d" % PORT)
    ThreadingHTTPServer.allow_reuse_address = True
    httpd = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        log("server dihentikan")

