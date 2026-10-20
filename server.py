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
import sys, os, json, time, socket, threading, datetime as dt
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


def xlsx_signature():
    """Tanda perubahan file Excel di folder data (nama + ukuran + waktu modif)."""
    sig = []
    try:
        files = os.listdir(DATA_DIR)
    except OSError:
        return ()
    for f in sorted(files):
        if f.lower().endswith(".xlsx") and not f.startswith("~$") and not f.startswith("."):
            p = os.path.join(DATA_DIR, f)
            try:
                st = os.stat(p)
                sig.append((f, st.st_size, int(st.st_mtime)))
            except OSError:
                pass
    return tuple(sig)


def watcher():
    """Kalau ada Excel yang diganti (mis. disalin lewat share jaringan), rebuild otomatis."""
    last = xlsx_signature()
    while True:
        time.sleep(20)
        try:
            cur = xlsx_signature()
            if cur != last:
                last = cur
                newest = max((os.path.join(DATA_DIR, f) for f, _, _ in cur), key=os.path.getmtime) if cur else None
                regenerate(newest, "perubahan file Excel")
        except Exception as e:
            log("watcher error: %s" % e)


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
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if u.path == "/api/status":
            meta = read_meta()
            self._json({"ok": True, "server": "HPS LAN", "started": _state["started"],
                        "generated": meta.get("generated", _state.get("generated", "")),
                        "rows": meta.get("rows", 0), "source": meta.get("source", _state.get("source", "")),
                        "builds": _state["builds"], "last_error": _state["last_error"],
                        "token_required": bool(TOKEN), "folder": DATA_DIR})
            return
        if u.path == "/api/health":
            self._json({"ok": True})
            return
        return super().do_GET()

    def do_POST(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if u.path == "/api/refresh":
            if self._need_token(q):
                return
            ok, info = regenerate(None, "refresh manual")
            self._json({"ok": ok, "info": info})
            return
        if u.path == "/api/upload":
            if self._need_token(q):
                return
            return self._upload(q)
        self._json({"ok": False, "error": "endpoint tidak dikenal"}, 404)

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


if __name__ == "__main__":
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
