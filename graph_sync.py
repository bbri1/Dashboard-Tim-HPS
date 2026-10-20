"""Tarik file Excel dari SharePoint / OneDrive secara realtime (Microsoft Graph).

File Excel-nya di-share lewat link OneDrive for Business, contoh:
  https://<tenant>.sharepoint.com/:x:/r/personal/<user>/_layouts/15/Doc.aspx?sourcedoc={ITEM-ID}

Yang dibutuhkan:
  1. Azure AD App Registration (cukup 1x) -> client id + secret + tenant id
     + permission "Files.Read.All" (Application) yang di-setujui admin
  2. Isi file graph_config.json di folder ini:
       {
         "tenant_id": "xxxx-xxxx",
         "client_id": "xxxx-xxxx",
         "client_secret": "xxxx",
         "user": "00396104@hq.bri.co.id",
         "item_id": "62fc6116-5f6d-4fa3-af62-6f29ebe6efc5",
         "target_dir": "C:/HPS-Dashboard/inbox",
         "poll_seconds": 60
       }
  3. Jalankan: python graph_sync.py   (bisa juga dijalankan sebagai service)

Skrip ini hanya BAWA file-nya ke folder target. Server HPS yang memakai
folder target itu sebagai "sumber realtime" (lihat hps_source.txt).
"""
import os, sys, json, time, shutil, urllib.request, urllib.error, urllib.parse, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
CFG = os.path.join(HERE, "graph_config.json")
LOG = os.path.join(HERE, "graph_sync.log")
GRAPH = "https://graph.microsoft.com/v1.0"


def log(msg):
    line = "%s  %s" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line, flush=True)
    try:
        with open(LOG, "a", encoding="utf8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def load_cfg():
    if not os.path.exists(CFG):
        raise SystemExit("graph_config.json belum ada - buat dulu (lihat komentar di atas file ini).")
    with open(CFG, encoding="utf8") as f:
        c = json.load(f)
    for k in ("tenant_id", "client_id", "client_secret", "user", "item_id"):
        if not c.get(k):
            raise SystemExit("graph_config.json belum lengkap: field '%s' kosong." % k)
    return c


def token(cfg):
    """Ambil access token (client credentials / app-only)."""
    url = "https://login.microsoftonline.com/%s/oauth2/v2.0/token" % cfg["tenant_id"]
    body = urllib.parse.urlencode({
        "client_id": cfg["client_id"],
        "client_secret": cfg["client_secret"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode()
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf8"))["access_token"]


def get_meta(tok, cfg):
    """Ambil metadata file (ETag + tanggal ubah).
    owner_user: pemilik drive tempat file berada (bisa berbeda dari akun pemakai)."""
    owner = cfg.get("owner_user") or cfg["user"]
    url = "%s/users/%s/drive/items/%s" % (GRAPH, urllib.parse.quote(owner), cfg["item_id"])
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + tok})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf8"))


def download(tok, cfg, dest):
    owner = cfg.get("owner_user") or cfg["user"]
    url = "%s/users/%s/drive/items/%s/content" % (GRAPH, urllib.parse.quote(owner), cfg["item_id"])
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + tok})
    tmp = dest + ".part"
    with urllib.request.urlopen(req, timeout=180) as r, open(tmp, "wb") as f:
        shutil.copyfileobj(r, f)
    os.replace(tmp, dest)


def main():
    cfg = load_cfg()
    target = cfg.get("target_dir") or os.path.join(HERE, "inbox")
    os.makedirs(target, exist_ok=True)
    name = cfg.get("file_name") or "Progress_Report_HPS.xlsx"
    dest = os.path.join(target, name)
    interval = int(cfg.get("poll_seconds") or 60)
    log("sync SharePoint -> %s (setiap %d detik)" % (dest, interval))
    last_etag, last_time = None, None
    while True:
        try:
            tok = token(cfg)
            meta = get_meta(tok, cfg)
            etag = meta.get("eTag") or meta.get("cTag") or ""
            mtime = meta.get("lastModifiedDateTime") or ""
            if etag != last_etag or mtime != last_time:
                download(tok, cfg, dest)
                last_etag, last_time = etag, mtime
                log("ter-download: %s (diubah %s, %d byte)" % (name, mtime, os.path.getsize(dest)))
            else:
                log("tidak ada perubahan (cek terakhir %s)" % mtime)
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf8", "replace")[:300]
            log("GAGAL (%s): %s" % (e.code, detail))
        except Exception as e:
            log("GAGAL: %s" % e)
        time.sleep(interval)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
