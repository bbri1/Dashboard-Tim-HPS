# Dashboard Progres HPS

**Situs: https://bbri1.github.io/Dashboard-Tim-HPS/**

Dashboard statis (HTML + JS) yang di-host di GitHub Pages, jadi:

- **tetap hidup 24 jam**, tidak bergantung pada laptop siapa pun (laptop boleh mati);
- bisa dibuka siapa saja di jaringan kantor (dan dari rumah), cukup pakai browser;
- **data ter-update otomatis** begitu file Excel baru diunggah — tidak perlu login ke server.

---

## 1. Alur kerja (ringkas)

```
File Excel (Progress_Report_HPS.xlsx)
        │  diunggah ke repository GitHub
        ▼
GitHub Actions  ──►  update_data.py  ──►  data.js + data.meta.json
        │
        ▼
GitHub Pages  ──►  https://bbri1.github.io/Dashboard-Tim-HPS/
        │
        ▼
Browser tiap user: cek data.meta.json tiap 60 detik → kalau berubah, data tampilannya
                  otomatis ter-update (tanpa tekan F5)
```

Waktu dari "Excel diunggah" sampai "situs jadi baru": **1–3 menit**.

---

## 2. Cara memperbarui data

### Cara A — Dari laptop (paling mudah, disarankan)

1. Salin file Excel terbaru ke folder proyek ini (folder yang berisi `publish_data.bat`).
2. Klik 2x **`publish_data.bat`**.
3. Tunggu. Script akan: memilih Excel terbaru → membuat `data.js` → commit → push ke GitHub.
4. Buka https://bbri1.github.io/Dashboard-Tim-HPS/ → data sudah baru dalam 1–3 menit.

> Kalau muncul "GAGAL push. belum login GitHub", buka **GitHub Desktop**
> (https://desktop.github.com) → clone repo `bbri1/Dashboard-Tim-HPS` → drag file Excel
> ke folder repo → Commit → Push. Situs tetap ter-update otomatis.

### Cara B — Langsung dari browser (tanpa perlu Python di laptop)

1. Buka https://github.com/bbri1/Dashboard-Tim-HPS → folder utama → **Add file → Upload files**.
2. Hapus file `Progress_Report_HPS.xlsx` yang lama, lalu upload file Excel baru dengan
   **nama yang sama persis** (`Progress_Report_HPS.xlsx`).
3. Commit → GitHub Actions otomatis membangun ulang `data.js` dan deploy situs.

### Cara C — Hanya memperbarui tampilan lokal

- Klik 2x `update_data.bat` → `data.js` diperbarui di folder ini (tidak ada yang diunggah).
- Klik 2x `start_server.bat` → buka `http://127.0.0.1:8080/` di browser.

---

## 3. Kalau internet mati (cadangan lokal)

Klik 2x `start_server.bat`, lalu rekan bisa buka `http://<IP-KANTOR>:8080/`.
Alamat IP ditampilkan di jendela server. Versi ini memakai data terakhir yang ada
di folder proyek.

---

## 4. Kebutuhan program (hanya untuk cara A dan C di laptop)

- Python 3 (cek: `python --version`)
- `pip install openpyxl` (sekali saja)
- Git + sudah login (untuk cara A). Alternatif: GitHub Desktop.

Server cloudnya (GitHub Pages) tidak butuh apa pun dari laptop.

---

## 5. Isi folder

| File | Fungsi |
|---|---|
| `index.html` | Tampilan dashboard (tidak butuh internet) |
| `data.js` | Data hasil baca Excel (dibuat otomatis oleh `update_data.py`) |
| `data.meta.json` | Timestamp + jumlah baris, dipakai browser untuk cek perubahan |
| `update_data.py` | Pembaca Excel → `data.js` |
| `publish_data.py` / `.bat` | Pilih Excel terbaru + build + commit + push |
| `update_data.bat` | Bangun ulang `data.js` di laptop saja |
| `server.py` / `start_server.bat` | Server lokal (cadangan saat internet mati) |
| `Progress_Report_HPS.xlsx` | Sumber data (sumber kebenaran) |
| `.github/workflows/deploy.yml` | Otomatisasi: Excel → `data.js` → deploy |

---

## 6. Catatan privasi

Repository `bbri1/Dashboard-Tim-HPS` bersifat **publik**, jadi siapa pun yang mengetahui
alamatnya bisa membuka dashboard (isi `data.js` juga bisa diunduh). Data di dalamnya
memuat nama pengadaan, nama user, dan nilai anggaran.

Kalau datanya tidak boleh publik, syaratkan akses — contoh: taruh situs di
Cloudflare Pages dengan login basic auth, atau minta admin mengubah repository menjadi
private (GitHub Pages private butuh GitHub Pro).

---

## 7. Kalau ada masalah

| Gejala | Penyebab & solusi |
|---|---|
| Situs masih data lama | Tunggu 1–3 menit / klik refresh. Lihat tab **Actions** di repository: workflow harus hijau. |
| Workflow gagal | Buka **Actions** → klik run yang merah → lihat log.通常是 `pip install openpyxl` atau file Excel rusak. |
| `GAGAL push` di `publish_data.bat` | Belum login GitHub → pakai GitHub Desktop atau cara B. |
| Halaman "data.js tidak ditemukan" | `data.js` belum ada → jalankan `update_data.bat`. |
| Ganti sumber data | Jangan ubah nama sheet / urutan header di Excel, agar pembacaan tetap berhasil. |
