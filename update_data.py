"""Bangun ulang data.js + data.meta.json dari file Excel Progress Report HPS.

Pemakaian:
    python update_data.py                 -> pakai file .xlsx terbaru di folder ini
    python update_data.py "nama.xlsx"     -> pakai file tertentu
    python update_data.py "C:\folder"     -> pakai .xlsx terbaru di folder tsb
Butuh: pip install openpyxl
"""
import sys, re, json, zipfile, tempfile, os, glob, datetime as dt
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.environ.get("HPS_DATA_DIR", HERE)   # folder tujuan data.js + file Excel
WIB = dt.timezone(dt.timedelta(hours=7))  # waktu Indonesia Barat, sama dengan di server CI
# sheet -> (tim, nama estimator aktif atau None untuk 'Lainnya')
SHEETS = {
 "NIT (Hendra Kusuma)": ("NIT", "Hendra Kusuma"), "NIT (Heru Setiawan)": ("NIT", "Heru Setiawan"),
 "NIT (Dian Nanda)": ("NIT", "Dian Nanda"), "NIT (Others)": ("NIT", None),
 "IT (Nugraeningtyas Viratri)": ("ITS", "Nugraeningtyas Viratri"), "IT (Jessicha Nurdiasari)": ("ITS", "Jessicha Nurdiasari"),
 "IT (Wahyuda Pujianto)": ("ITS", "Wahyuda Pujianto"), "IT (Others)": ("ITS", None),
 "CSE (Kamila Mahmudah)": ("CSE", "Kamila Mahmudah"), "CSE (Tiara Bella)": ("CSE", "Tiara Bella"),
 "CSE (Raina Arfa)": ("CSE", "Raina Arfa"), "CSE (Others)": ("CSE", None),
}


def newest_xlsx(folder):
    cand = [p for p in glob.glob(os.path.join(folder, "*.xlsx"))
            if not os.path.basename(p).startswith("~$") and not os.path.basename(p).startswith(".")]
    if not cand:
        return None
    # terbaru; kalau seri mtime, file baku (Progress_Report_HPS.xlsx) yang menang
    return max(cand, key=lambda p: (os.path.getmtime(p), os.path.basename(p) == "Progress_Report_HPS.xlsx"))


def find_source(arg):
    """Tentukan file Excel sumber: argumen boleh file, folder, atau kosong."""
    if arg and os.path.isdir(arg):
        src = newest_xlsx(arg)
        if not src:
            raise SystemExit("ERROR: tidak ada file .xlsx di folder: " + arg)
        return src
    if arg and os.path.isfile(arg):
        return arg
    src = newest_xlsx(OUTDIR)
    if not src:
        raise SystemExit("ERROR: tidak ada file .xlsx di folder script.")
    return src


def fix_styles(path):  # beberapa file Excel punya nilai font family tak valid -> openpyxl gagal
    tmp = tempfile.mktemp(suffix=".xlsx")
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            b = zi.read(it.filename)
            if it.filename == "xl/styles.xml":
                b = re.sub(rb'<family val="(\d+)"/>', lambda m: b'<family val="2"/>' if int(m.group(1)) > 14 else m.group(0), b)
            zo.writestr(it, b)
    return tmp


def d(v):
    if isinstance(v, dt.datetime): return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date): return v.strftime("%Y-%m-%d")
    return None
def n(v):
    return float(v) if isinstance(v, (int, float)) else 0


def networkdays(a, b, holidays):
    """Sama dengan rumus Excel NETWORKDAYS(start, end, hari_libur)."""
    if a is None or b is None: return None
    d0 = a.date() if isinstance(a, dt.datetime) else a
    d1 = b.date() if isinstance(b, dt.datetime) else b
    if d1 < d0: return None
    c, n_ = d0, 0
    while c <= d1:
        if c.weekday() < 5 and c not in holidays: n_ += 1
        c += dt.timedelta(days=1)
    return n_


def build(src):
    tmp = fix_styles(src)
    wb = load_workbook(tmp, data_only=True, read_only=True)
    H = {  # nama kolom -> pencarian prefix header
     "no": "No. HPS", "nama": "Nama Pengadaan", "user": "User", "tmin": "Tanggal Permintaan HPS", "pre": "Tanggal Pre Screening",
     "lkp": "Tanggal Dokumen Lengkap", "put": "Tanggal Putusan", "rilis": "Tanggal Rilis HPS", "hps": "Nilai HPS",
     "status": "Status", "sla": "SLA (Hari Kerja)", "kat": "Kategori SLA", "trfi": "Tanggal Permintaan RFI"}
    # hari libur (dipakai untuk menghitung SLA yang belum tersimpan di file)
    holidays = set()
    if "Daftar Libur" in wb.sheetnames:
        for hr in wb["Daftar Libur"].iter_rows(min_row=3, max_row=30, max_col=2, values_only=True):
            hv = hr[1] if len(hr) > 1 else None
            if isinstance(hv, dt.datetime): holidays.add(hv.date())
            elif isinstance(hv, dt.date): holidays.add(hv)
    rows, seen = [], set()
    dipakai = []
    for sh, (tim, est) in SHEETS.items():
        if sh not in wb.sheetnames:
            print("  lewati sheet tidak ada:", sh)
            continue
        dipakai.append(sh)
        R = [list(r) for r in wb[sh].iter_rows(values_only=True)]
        if not R: continue
        h = [str(x).strip() if x else "" for x in R[0]]
        ix = {k: h.index(v) for k, v in H.items() if v in h}
        if "nama" not in ix or "no" not in ix:
            print("  lewati sheet (header tidak cocok):", sh)
            continue
        ang = next((i for i, x in enumerate(h) if x in ("Nilai Anggaran/Usulan", "Nilai Anggaran")), None)
        for r in R[1:]:
            if not r or not isinstance(r[0], (int, float)) or not r[ix["nama"]]: continue  # baris data = ada nomor urut & nama
            key = (r[ix["no"]] or r[ix["nama"]], sh, r[0])
            if key in seen: continue
            seen.add(key)
            st = str(r[ix["status"]] or "").strip() if "status" in ix else ""
            if st.lower() == "done": st = "Done"
            g = lambda k: r[ix[k]] if k in ix and ix[k] < len(r) else None
            sla = g("sla") if isinstance(g("sla"), (int, float)) else None
            if sla is None and st.lower() == "done":
                # rumus Excel: NETWORKDAYS(RFI atau Tgl Permintaan, Tgl Rilis, hari libur)
                sl = networkdays(g("trfi") or g("tmin"), g("trilis"), holidays)
                if sl is not None: sla = sl
            rows.append([str(g("no") or "-").strip(), str(g("nama") or "").strip(), tim, est or "Others*", str(g("user") or "").strip(),
                         n(r[ang]) if ang is not None else 0, d(g("tmin")), d(g("pre")), d(g("tlkp")), d(g("tput")), d(g("trilis")),
                         n(g("hps")), st, sla, str(g("kat") or "").strip()])
    wb.close()
    try: os.remove(tmp)
    except OSError: pass
    return rows, dipakai


COLS = ["no", "nama", "tim", "est", "user", "anggaran", "tmin", "tpre", "tlkp", "tput", "trilis", "hps", "status", "sla", "kat"]


def generate(src=None):
    """Baca Excel -> tulis data.js + data.meta.json. Kembalikan dict meta."""
    src = src or find_source("")
    rows, dipakai = build(src)
    gen = dt.datetime.now(WIB).strftime("%Y-%m-%d %H:%M")
    out = {"generated": gen, "cols": COLS, "rows": rows}
    with open(os.path.join(OUTDIR, "data.js"), "w", encoding="utf8") as f:
        f.write("window.HPS=" + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";")
    meta = {"generated": gen, "rows": len(rows), "source": os.path.basename(src),
            "sheets": len(dipakai), "size": os.path.getsize(os.path.join(OUTDIR, "data.js"))}
    with open(os.path.join(OUTDIR, "data.meta.json"), "w", encoding="utf8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    return meta


if __name__ == "__main__":
    src = find_source(sys.argv[1] if len(sys.argv) > 1 else "")
    print("Sumber Excel :", src)
    meta = generate(src)
    print("%d baris ditulis ke data.js (%d sheet dipakai)" % (meta["rows"], meta["sheets"]))
    print("Timestamp   :", meta["generated"])
