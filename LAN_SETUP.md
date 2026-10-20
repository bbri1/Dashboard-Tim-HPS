# Dashboard HPS — Setup Jaringan Kantor (Lokal)

Dashboard ini **tidak perlu internet**. Satu PC/server di kantor menjalankan
`server.py`, lalu semua rekan cukup membuka alamat IP PC itu di browser.

```
                 JARINGAN KANTOR (192.168.x.x)
                            |
   [PC SERVER: C:\HPS-Dashboard]  <-- server.py (port 8080)
        |                             index.html + data.js
        |  Excel baru diunggah / disalin ke folder
        v
   [PC/HP mana pun] --> http://192.168.1.20:8080/
```

---

## 1. Siapkan PC server (sekali saja)

| # | Langkah |
|---|---|
| 1 | Salin seluruh folder proyek ini ke PC server, contoh `C:\HPS-Dashboard` |
| 2 | Pasang **Python 3** di PC itu (python.org, centang *Add to PATH*) |
| 3 | Buka CMD di folder `C:\HPS-Dashboard`, jalankan: `pip install openpyxl` |
| 4 | Klik 2x **`install_service.bat`** → akan minta hak admin, lalu:<br>• membuat service auto-start + auto-restart<br>• membuka port 8080 di firewall<br>• langsung menyalakan server |
| 5 | Catat IP PC server: buka CMD, ketik `ipconfig` → misal `192.168.1.20` |

> Pakai `uninstall_service.bat` kalau ingin mencopot service-nya.

### Alamat untuk rekan-rekan
```
http://<IP-PC-SERVER>:8080/            -> dashboard
http://<IP-PC-SERVER>:8080/upload.html -> unggah Excel
```

**Penting:** address itu hanya berlaku di dalam jaringan kantor (tidak ada di internet).
Kalau alamat IP PC server berubah tiap hari, minta admin jaringan membuat
**reservasi DHCP** (IP statis) supaya link-nya tidak berubah-ubah.

---

## 2. Cara memperbarui data (laptop boleh mati)

### Cara A — Unggah dari browser (paling mudah)
1. Dari PC/HP mana pun di kantor, buka `http://<IP-PC-SERVER>:8080/upload.html`
2. Drag file Excel ke kotak yang tersedia (atau klik untuk memilih file)
3. Tunggu beberapa detik → muncul "Berhasil! 1459 baris"
4. Dashboard semua user ter-update otomatis dalam **maks 60 detik**
   (tidak perlu tekan F5; muncul notifikasi "Data diperbarui")

### Cara B — Salin ke folder server
Salin file `.xlsx` ke `C:\HPS-Dashboard` (lewat File Explorer / share jaringan).
Server mendeteksi perubahan dalam 20 detik lalu membangun ulang data.

### Cara C — Di PC server langsung
Klik 2x `update_data.bat`.

---

## 3. Yang perlu diketahui

| Hal | Keterangan |
|---|---|
| PC server harus menyala | Kalau PC dimatikan (malam/weekend), dashboard ikut mati. Untuk 24 jam, lihat bagian 4. |
| Update tanpa internet | Semua berjalan lokal. GitHub hanya dipakai sebagai cadangan, bukan untuk serving. |
| File Excel | `Progress_Report_HPS.xlsx` = sumber data. `data.js` = hasil bacanya (otomatis). |
| Refresh browser | Otomatis 60 detik. Kalau perlu paksa: tekan F5. |
| Log server | `hps_server.log` di folder server. Buka juga `http://<IP>:8080/api/status` untuk status. |
| Keamanan | Hanya orang di jaringan kantor yang bisa mengakses. Kalau mau batasi, set token: `setx HPS_TOKEN "rahasia"` lalu jalankan ulang service — token diisi di halaman upload. |

---

## 4. (Opsional) Supaya jalan 24 jam di Raspberry Pi / mini PC / NAS

### Docker (Raspberry Pi, mini PC, NAS, Linux)
```bash
# di folder ini
docker compose up -d
```
- Dashboard: `http://<IP>:8080/`
- Taruh file Excel di folder `./data` → rebuild otomatis 20 detik
- `docker-compose.yml` sudah memakai `restart: unless-stopped` (nyala lagi otomatis setelah restart/reboot)

### systemd (Linux tanpa Docker)
```bash
sudo cp hps-dashboard.service /etc/systemd/system/
sudo mkdir -p /var/lib/hps-dashboard
sudo systemctl daemon-reload
sudo systemctl enable --now hps-dashboard
journalctl -u hps-dashboard -f
```

### Windows tanpa service (server manual)
```bat
start_server.bat
```
Berguna untuk testing; kalau jendela ditutup, server ikut berhenti.

---

## 5. Kalau ada masalah

| Gejala | Solusi |
|---|---|
| Rekan tidak bisa buka `http://IP:8080` | 1) cek server hidup: di PC server buka `http://127.0.0.1:8080`<br>2) cek firewall: `install_service.bat` ulang<br>3) cek PC server & rekan satu jaringan (WiFi kantor yang sama) |
| Halaman upload error "token salah" | Isi kolom token di halaman upload (dikunci di browser setelah diisi) |
| Dashboard masih data lama | Tunggu 60 detik; cek `hps_server.log`; klik tombol "Perbarui" di halaman upload |
| `data.js tidak ditemukan` | Jalankan `update_data.bat` di folder server (butuh `pip install openpyxl`) |
| Gagal upload: "file bukan Excel valid" | File bukan .xlsx asli (atau .xls) — export ulang dari Excel sebagai .xlsx |
| Error Python | `pip install openpyxl` (butuh internet **sekali saja** saat installing) |
