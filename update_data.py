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


def find_source(arg):
    """Tentukan file Excel sumber: argumen boleh file, folder, atau kosong."""
    if arg and os.path.isdir(arg):
        cand = [p for p in glob.glob(os.path.join(arg, "*.xlsx"))
                if not os.path.basename(p).startswith("~$") and not os.path.basename(p).startswith(".")]
        if not cand:
            raise SystemExit("ERROR: tidak ada file .xlsx di folder: " + arg)
        return max(cand, key=os.path.getmtime)
    if arg and os.path.isfile(arg):
        return arg
    cand = [p for p in glob.glob(os.path.join(HERE, "*.xlsx"))
            if not os.path.basename(p).startswith("~$") and not os.path.basename(p).startswith(".")]
    if not cand:
        raise SystemExit("ERROR: tidak ada file .xlsx di folder script.")
    return max(cand, key=os.path.getmtime)


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
    return None
def n(v):
    return float(v) if isinstance(v, (int, float)) else 0


def build(src):
    tmp = fix_styles(src)
    wb = load_workbook(tmp, data_only=True, read_only=True)
    H = {  # nama kolom -> pencarian prefix header
     "no": "No. HPS", "nama": "Nama Pengadaan", "user": "User", "tmin": "Tanggal Permintaan HPS", "pre": "Tanggal Pre Screening",
     "lkp": "Tanggal Dokumen Lengkap", "put": "Tanggal Putusan", "rilis": "Tanggal Rilis HPS", "hps": "Nilai HPS",
     "status": "Status", "sla": "SLA (Hari Kerja)", "kat": "Kategori SLA"}
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
            rows.append([str(g("no") or "-").strip(), str(g("nama") or "").strip(), tim, est or "Lainnya*", str(g("user") or "").strip(),
                         n(r[ang]) if ang is not None else 0, d(g("tmin")), d(g("pre")), d(g("lkp")), d(g("put")), d(g("rilis")),
                         n(g("hps")), st, g("sla") if isinstance(g("sla"), (int, float)) else None,
                         str(g("kat") or "").strip()])
    wb.close()
    try: os.remove(tmp)
    except OSError: pass
    return rows, dipakai


if __name__ == "__main__":
    src = find_source(sys.argv[1] if len(sys.argv) > 1 else "")
    print("Sumber Excel :", src)
    rows, dipakai = build(src)
    cols = ["no", "nama", "tim", "est", "user", "anggaran", "tmin", "tpre", "tlkp", "tput", "trilis", "hps", "status", "sla", "kat"]
    gen = dt.datetime.now(WIB).strftime("%Y-%m-%d %H:%M")
    out = {"generated": gen, "cols": cols, "rows": rows}
    with open(os.path.join(HERE, "data.js"), "w", encoding="utf8") as f:
        f.write("window.HPS=" + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";")
    meta = {"generated": gen, "rows": len(rows), "source": os.path.basename(src),
            "sheets": len(dipakai), "size": os.path.getsize(os.path.join(HERE, "data.js"))}
    with open(os.path.join(HERE, "data.meta.json"), "w", encoding="utf8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print("%d baris ditulis ke data.js (%d sheet dipakai)" % (len(rows), len(dipakai)))
    print("Timestamp   :", gen)
