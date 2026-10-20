"""Unggah data HPS terbaru ke GitHub Pages (dari laptop ini).

Yang dilakukan:
  1. Cari file Excel (.xlsx) terbaru di folder ini
  2. Salin jadi Progress_Report_HPS.xlsx (nama baku di repository)
  3. Jalankan update_data.py  -> data.js + data.meta.json
  4. git commit + push ke https://github.com/bbri1/Dashboard-Tim-HPS
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


def newest_xlsx():
    cand = [p for p in glob.glob(os.path.join(HERE, "*.xlsx"))
            if not os.path.basename(p).startswith("~$") and not os.path.basename(p).startswith(".")
            and os.path.basename(p) != CANON]
    if cand:
        return max(cand, key=os.path.getmtime)
    canon = os.path.join(HERE, CANON)
    return canon if os.path.isfile(canon) else None


def main():
    src = newest_xlsx()
    if not src:
        print("GAGAL: tidak ada file .xlsx di folder ini.")
        print("       Salin file 'Progress Report HPS' (terbaru) ke folder ini lalu jalankan ulang.")
        return 1

    canon = os.path.join(HERE, CANON)
    if os.path.abspath(src) != os.path.abspath(canon):
        print("Excel sumber :", os.path.basename(src))
        shutil.copy2(src, canon)
        print("Disalin ke   :", CANON)
    else:
        print("Excel sumber :", CANON)

    print("\n[1/3] Membangun data.js ...")
    rc, out = run([sys.executable, os.path.join(HERE, "update_data.py")])
    print(out.strip())
    if rc != 0:
        print("GAGAL saat membuat data.js (lihat pesan di atas).")
        return 1

    print("\n[2/3] Menge-commit ke GitHub ...")
    rc, out = run(["git", "rev-parse", "--is-inside-work-tree"])
    if rc != 0:
        print("Folder ini belum jadi repository git.")
        print("Jalankan sekali: git init && git remote add origin %s" % REPO)
        return 1

    run(["git", "add", "-A"])
    rc, out = run(["git", "diff", "--cached", "--quiet"])
    if rc == 0:
        print("Tidak ada perubahan -> tidak ada yang diunggah.")
        print("Situs: " + SITE)
        return 0

    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    rc, out = run(["git", "commit", "-m", "update data HPS %s" % stamp])
    print(out.strip())

    print("\n[3/3] Mengirim ke GitHub ...")
    run(["git", "pull", "--rebase", "--autostash", "origin", BRANCH])
    rc, out = run(["git", "push", "origin", BRANCH])
    print(out.strip())
    if rc != 0:
        print("\nGAGAL push. Penyebab umum: belum login GitHub.")
        print("  Opsi A: pakai GitHub Desktop (kloning repo, lalu drag file Excel ke folder repo).")
        print("  Opsi B: cache token di terminal  ->  setx GITHUB_TOKEN <token dari github.com/settings/tokens>")
        print("  Opsi C: unggah manual ke https://github.com/bbri1/Dashboard-Tim-HPS (hapus file lama, upload yang baru).")
        return 1

    print("\nSelesai. Situs akan ter-update dalam 1-3 menit:")
    print("  " + SITE)
    return 0


if __name__ == "__main__":
    sys.exit(main())
