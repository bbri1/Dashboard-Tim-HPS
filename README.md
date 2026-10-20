# Dashboard Progres HPS — Jaringan Kantor (lokal)

Berjalan **di jaringan lokal kantor**, tanpa internet, tanpa cloud.
Satu PC/server di kantor menjalankan `server.py`; semua rekan membuka alamat IP
tersebut di browser. Data ter-update otomatis begitu Excel baru diunggah —
laptop siapa pun boleh mati.

```
[PC SERVER: C:\HPS-Dashboard]  --(port 8080)-->  [PC/HP rekan di kantor]
        |  Excel baru (upload / salin ke folder)
        v
   data.js + data.meta.json  -->  semua layar ter-update dalam 60 detik
```

## Mulai cepat

1. Di **PC server**: salin folder ini → `C:\HPS-Dashboard`, pasang Python 3,
   `pip install openpyxl`, lalu klik 2x **`install_service.bat`**.
2. Cek IP PC server (`ipconfig`), misal `192.168.1.20`.
3. Sampaikan ke rekan: **`http://192.168.1.20:8080/`**
4. Update data kapan saja: **`http://192.168.1.20:8080/upload.html`** → drag Excel.

Detail & troubleshooting: **[LAN_SETUP.md](LAN_SETUP.md)** · Panduan singkat: `BACA_SAYA.txt`

## Fitur tampilan

- **6 KPI card** — Total HPS, Nilai HPS Selesai, % Selesai, Sedang Diproses, Rata-rata SLA, Perlu Perhatian (>30 hari)
- **Pipeline status** (klik = filter) & **Tren permintaan 12 bulan** (area chart SVG, klik titik = filter)
- **SLA per estimator** (dengan sebaran 1–10 / 11–15 / >15 hari), **Ringkasan per tim** (donat), **Top 5 nilai HPS**, **Sebaran umur HPS on progress**
- **Tabel HPS** — sortable, status pill, umur/SLA berwarna, pagination, export CSV
- **Mode terang/gelap** (tombol kanan atas, tersimpan di browser), **kerapatan tabel**, filter aktif sebagai chip yang bisa dihapus
- Shortcut `/` untuk fokus ke kolom pencarian · auto-refresh tiap 60 detik (tanpa F5)
- Untuk layar monitor kantor: buka `http://<IP>:8080/?theme=light`

## Isi folder

| File | Fungsi |
|---|---|
| `index.html` | Tampilan dashboard |
| `upload.html` | Halaman unggah Excel (drag & drop) |
| `server.py` | Server lokal: menyajikan dashboard, API unggah/refresh, watcher Excel |
| `update_data.py` | Excel → `data.js` (+ `data.meta.json`) |
| `data.js` / `data.meta.json` | Data hasil baca Excel & penanda waktu |
| `install_service.ps1` / `.bat` | Pasang service Windows: auto-start, auto-restart, firewall |
| `uninstall_service.bat` | Copot service |
| `start_server.bat` | Jalankan server manual (testing) |
| `update_data.bat` | Build `data.js` di PC server |
| `publish_data.bat` / `.py` | Opsional: salin data ke repo cadangan GitHub |
| `Dockerfile`, `docker-compose.yml`, `hps-dashboard.service` | Paket untuk Raspberry Pi / NAS / Linux |
| `Progress_Report_HPS.xlsx` | Sumber data (benar) |

## API server (untuk otomasi / integrasi)

| Method | Endpoint | Fungsi |
|---|---|---|
| GET | `/api/status` | Status server, jumlah baris, error terakhir |
| GET | `/api/health` | Health check |
| POST | `/api/upload` | Unggah Excel (body = file .xlsx, header `X-Filename`) |
| POST | `/api/refresh` | Build ulang dari Excel yang ada di folder |

Setelan opsional (environment variable): `HPS_PORT` (default 8080),
`HPS_TOKEN` (token wajib untuk unggah), `HPS_DATA_DIR` (folder data).

## Catatan

- Butuh internet **hanya sekali** di PC server saat `pip install openpyxl`.
- PC server harus menyala; kalau dimatikan malam/weekend, dashboard ikut mati.
  Untuk 24 jam penuh, jalankan di Raspberry Pi / mini PC dengan Docker.
- GitHub hanya dipakai sebagai cadangan kode/data; situs publiknya sudah dimatikan.
- Jangan ubah nama sheet / urutan header di Excel agar pembacaan tetap berhasil.
