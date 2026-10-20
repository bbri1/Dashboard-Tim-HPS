"""Unggah data HPS terbaru ke GitHub Pages (dari laptop ini).

Yang dilakukan:
  1. Sinkronkan dengan GitHub (git pull) supaya data di server tidak tertinggal
  2. Cari file Excel (.xlsx) terbaru di folder ini
  3. Salin jadi Progress_Report_HPS.xlsx (nama baku di repository)
  4. Jalankan update_data.py  -> data.js + data.meta.json
  5. git commit + push ke https://github.com/bbri1/Dashboard-Tim-HPS
     -> GitHub Actions otomatis deploy, situs jadi baru dalam ~1-3 menit

Jalankan: python publish_data.py   (atau klik 2x publish_data.bat)
"""
import os, sys, glob, shutil, subprocess, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
CANON = "Progress_Report_HPS.xlsx"
REPO = "https://github.com/bbri1/Dashboard-Tim-HPS.git"
SITE = "https://bbri1.github.io/Dashboard-Tim-HPS/"
BRANCH = "main"


def run(args, cwd=HERE):
    p = subprocess.run(args, cwd=cwd, capture_output=True, text=True, encoding="utf8", errors="replace")
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def candidates():
    return [p for p in glob.glob(os.path.join(HERE, "*.xlsx"))
            if not os.path.basename(p).startswith("~$") and not os.path.basename(p).startswith(".")]


def pick_source():
    """File Excel terbaru (bakal dipakai). Kalah seri: file baku (Progress_Report_HPS.xlsx)."""
    c = candidates()
    if not c:
        return None
    return max(c, key=lambda p: (os.path.getmtime(p), os.path.basename(p) == CANON))


def main():
    canon = os.path.join(HERE, CANON)
    src = pick_source()
    if not src:
        print("GAGAL: tidak ada file .xlsx di folder ini.")
        print("       Salin file 'Progress Report HPS' (terbaru) ke folder ini lalu jalankan ulang.")
        return 1

    print("[1/4] Sinkron dengan GitHub ...")
    rc, _ = run(["git", "rev-parse", "--is-inside-work-tree"])
    if rc != 0:
        print("Folder ini belum jadi repository git.")
        print("Jalankan sekali: git init && git remote add origin %s" % REPO)
        return 1
    rc, out = run(["git", "pull", "--rebase", "--autostash", "origin", BRANCH])
    if rc != 0:
        print("  (!) pull gagal, lanjut dengan isi lokal:")
        print("      " + out.strip().splitlines()[-1] if out.strip() else "")

    print("\n[2/4] Menyiapkan Excel sumber ...")
    print("  dipakai :", os.path.basename(src))
    if os.path.abspath(src) != os.path.abspath(canon):
        if os.path.exists(canon) and os.path.getmtime(canon) > os.path.getmtime(src):
            print("  (!) File Excel di server GitHub lebih baru dari file lokal.")
            print("      Batal mengunggah. Cek https://github.com/bbri1/Dashboard-Tim-HPS")
            print("      atau hapus file Excel lama di folder ini lalu jalankan ulang.")
            return 1
        shutil.copy2(src, canon)
        print("  disalin :", CANON)

    print("\n[3/4] Membangun data.js ...")
    rc, out = run([sys.executable, os.path.join(HERE, "update_data.py")])
    print("  " + out.strip().replace("\n", "\n  "))
    if rc != 0:
        print("GAGAL saat membuat data.js (lihat pesan di atas).")
        return 1

    print("\n[4/4] Commit & push ke GitHub ...")
    run(["git", "add", "-A"])
    rc, _ = run(["git", "diff", "--cached", "--quiet"])
    if rc == 0:
        print("  tidak ada perubahan -> tidak ada yang diunggah.")
        print("\nSitus: " + SITE)
        return 0

    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    rc, out = run(["git", "commit", "-m", "update data HPS %s" % stamp])
    if rc != 0:
        print("  " + out.strip())
    rc, out = run(["git", "push", "origin", BRANCH])
    print("  " + out.strip().splitlines()[-1] if out.strip() else "")
    if rc != 0:
        print("\nGAGAL push. Penyebab umum: belum login GitHub.")
        print("  Opsi A: GitHub Desktop -> clone repo, drag file Excel ke folder repo, Commit, Push.")
        print("  Opsi B: setx GITHUB_TOKEN <token dari github.com/settings/tokens> lalu ulangi.")
        print("  Opsi C: upload manual di https://github.com/bbri1/Dashboard-Tim-HPS")
        print("          (hapus file lama, upload yang baru dengan nama sama).")
        return 1

    print("\nSelesai. Situs ter-update dalam 1-3 menit:")
    print("  " + SITE)
    return 0


if __name__ == "__main__":
    sys.exit(main())
