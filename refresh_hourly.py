"""Dijalankan penjadwal Windows tiap 1 jam: ambil data HPS terbaru lalu minta server membangun ulang dashboard.

Urutan:
  1. Jika graph_config.json ada -> tarik Excel dari SharePoint (graph_sync.py --once)
  2. Minta server refresh (POST /api/refresh): server menyalin Excel terbaru dari folder sumber
     (hps_source.txt, mis. folder OneDrive) lalu membangun ulang data.js
Hasil dicatat di hps_refresh.log (maks ~200 KB, otomatis dipangkas).
"""
import os, sys, json, subprocess, urllib.request, urllib.error, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(HERE, "hps_refresh.log")
PORT = os.environ.get("HPS_PORT", "8080")
TOKEN = os.environ.get("HPS_TOKEN", "")


def log(msg):
    line = "%s  %s\n" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    try:
        old = open(LOG, encoding="utf8").read() if os.path.exists(LOG) else ""
        open(LOG, "w", encoding="utf8").write((old + line)[-200000:])
    except OSError:
        pass


def main():
    if os.path.exists(os.path.join(HERE, "graph_config.json")):
        r = subprocess.run([sys.executable, os.path.join(HERE, "graph_sync.py"), "--once"],
                           cwd=HERE, capture_output=True, text=True, timeout=300)
        log("SharePoint: %s" % ("OK" if r.returncode == 0 else "GAGAL (lihat graph_sync.log)"))
    req = urllib.request.Request("http://127.0.0.1:%s/api/refresh" % PORT, data=b"", method="POST",
                                 headers={"X-Token": TOKEN} if TOKEN else {})
    try:
        with urllib.request.urlopen(req, timeout=180) as r:
            j = json.loads(r.read().decode("utf8"))
        info = j.get("info") or {}
        log("refresh %s: %s baris (data per %s)" % ("OK" if j.get("ok") else "GAGAL: " + str(info.get("error")),
                                                   info.get("rows", "?"), info.get("generated", "?")))
    except (urllib.error.URLError, OSError) as e:
        log("server tidak menjawab di port %s: %s" % (PORT, e))
        raise SystemExit(1)


if __name__ == "__main__":
    main()
