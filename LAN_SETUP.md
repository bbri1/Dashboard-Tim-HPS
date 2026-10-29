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

## 3. Data realtime dari SharePoint / OneDrive

File Excel HPS berada di OneDrive for Business (SharePoint), link-nya seperti:

```
https://<tenant>.sharepoint.com/:x:/r/personal/00396104_hq_bri_co_id/_layouts/15/Doc.aspx?sourcedoc={62fc6116-5f6d-4fa3-af62-6f29ebe6efc5}
```

Supaya dashboard ikut ter-update otomatis begitu Excel di SharePoint berubah, ada **2 jalur**.
Pilih salah satu (Jalur A paling cepat, tidak perlu izin admin).

### Jalur A — OneDrive sync di PC server (disarankan, ±10 menit)

> **Catatan penting:** file Excel berada di OneDrive **pribadi** pemilik (00396104@hq.bri.co.id).
> OneDrive client **tidak** otomatis menyinkronkan file yang hanya di-*share* ke akun lain
> (masuk ke folder "Shared with me" di browser, tidak di-sync). Jadi 2 kemungkinan:
> - **File-nya dipindahkan ke folder SharePoint tim** (folder bersama) -> di OneDrive akun
>   90194837 pilih **"Add shortcut from SharePoint"** -> folder tim itu yang ter-sync
> - atau file tetap di OneDrive pribadi pemilik -> pakai **Jalur B (Graph API)** di bawah.

1. Di **PC server**, buka OneDrive (ikon awan di taskbar; kalau belum ada, pasang dari
   `https://www.microsoft365.com/apps/onedrive`) lalu **login dengan akun `90194837@hq.bri.co.id`**.
2. Pastikan file-nya bisa diakses akun tersebut:
   - kalau file dipindah ke **folder SharePoint tim** -> OneDrive 90194837: "Add shortcut from SharePoint"
   - atau file di-*share* dari OneDrive pribadi pemilik ke `90194837@hq.bri.co.id`
3. OneDrive akan men-download foldernya ke
   `C:\Users\90194837\OneDrive - PT Bank Rakyat Indonesia (BRI)`.
4. Buat file **`hps_source.txt`** di folder aplikasi HPS (`C:\HPS-Dashboard`), isi **satu baris**
   path folder tempat file Excel-nya:
   ```
   C:\Users\90194837\OneDrive - PT Bank Rakyat Indonesia (BRI)\Documents\HPS
   ```
   (boleh juga path file `.xlsx` langsung)
5. Restart PC server (atau jalankan ulang service HPS) supaya path dibaca.
6. **Selesai.** Setiap kali Excel di OneDrive di-save, dashboard ter-update dalam **±30 detik**.

Cek hasilnya: di bagian bawah dashboard ada tulisan *"Sumber realtime: ..."*, dan di
`hps_server.log` ada baris `Excel diambil dari sumber: ...` + `data.js diperbarui (sumber realtime)`.

### Jalur B — Server menarik sendiri dari SharePoint (Microsoft Graph)

Tidak butuh OneDrive di PC server, tapi perlu **Azure AD App Registration** (sekali saja, biasanya
memerlukan persetujuan admin):

1. Azure Portal -> **Entra ID -> App registrations -> New registration**
   - Name: `HPS Dashboard Sync`, Supported account types: *Single tenant* -> Create
2. **Certificates & secrets -> New client secret** -> salin nilainya (sekali saja, tidak bisa dilihat lagi)
3. **API permissions -> Add a permission -> Microsoft Graph -> Application permissions**
   - pilih `Files.Read.All` (dan `Sites.Read.All` bila filenya di SharePoint library) -> **Grant admin consent**
4. Salin `graph_config.json.example` -> `graph_config.json`, lalu isi:
   ```json
   {
     "tenant_id": "<tenant-id>",
     "client_id": "<client-id>",
     "client_secret": "<secret>",
     "user": "90194837@hq.bri.co.id",
     "owner_user": "00396104@hq.bri.co.id",
     "item_id": "62fc6116-5f6d-4fa3-af62-6f29ebe6efc5",
     "target_dir": "C:/HPS-Dashboard/inbox",
     "poll_seconds": 60
   }
   ```
   - `user` = akun yang dipakai server / PC server
   - `owner_user` = pemilik drive tempat file Excel berada (boleh dikosongkan kalau filenya
     ada di drive `user`)
   (`item_id` diambil dari link SharePoint: bagian `sourcedoc={...}`)
5. Isi `hps_source.txt` dengan folder tujuan: `C:\HPS-Dashboard\inbox`
6. Jalankan `python graph_sync.py` (bisa juga dijadwalkan lewat Task Scheduler agar jalan terus).
   File ditarik tiap 60 detik bila berubah → server langsung rebuild.

### Catatan
- JANGAN menaruh `graph_config.json` di repository GitHub (isi secret).
- Bila Excel sedang dibuka & disimpan berkali-kali, dashboard menunggu file "stabil" ±5 detik
  supaya tidak membaca file yang belum selesai ditulis.
- Jalur A tetap bisa dipakai bersamaan dengan unggah manual lewat `upload.html`.

---

## 4. Yang perlu diketahui

| Hal | Keterangan |
|---|---|
| PC server harus menyala | Kalau PC dimatikan (malam/weekend), dashboard ikut mati. Untuk 24 jam, lihat bagian 4. |
| Update tanpa internet | Semua berjalan lokal. GitHub hanya dipakai sebagai cadangan, bukan untuk serving. |
| File Excel | `Progress_Report_HPS.xlsx` = sumber data. `data.js` = hasil bacanya (otomatis). |
| Refresh browser | Otomatis 60 detik. Kalau perlu paksa: tekan F5. |
| Log server | `hps_server.log` di folder server. Buka juga `http://<IP>:8080/api/status` untuk status. |
| Keamanan | Hanya orang di jaringan kantor yang bisa mengakses. Kalau mau batasi, set token: `setx HPS_TOKEN "rahasia"` lalu jalankan ulang service — token diisi di halaman upload. |

---

## 5. (Opsional) Supaya jalan 24 jam di Raspberry Pi / mini PC / NAS

### Docker (Raspberry Pi, mini PC, NAS, Linux)
```bash
# di folder ini
docker compose up -d
```
- Dashboard: `http://<IP>:8080/`
- Taruh file Excel di folder `./data` → rebuild otomatis 20 detik
- `docker-compose.yml` sudah memakai `restart: unless-stopped` (nyala lagi otomatis setelah restart/reboot)
- Data (`data.js`, Excel, `hps_state.json`) tersimpan di folder `./data` sehingga aman saat container dibuat ulang
- Setelah mengubah kode: `docker compose up -d --build`
- **Realtime dari SharePoint** (OneDrive client tidak ada di dalam container, jadi pakai Graph API):
  1. salin `graph_config.json.example` menjadi `graph_config.json`, isi tenant/client/secret, dan ubah `"target_dir"` menjadi `"/data"`
  2. `docker compose --profile sharepoint up -d`
  Layanan `graph-sync` menarik file ke `./data`, lalu dashboard membangun ulang datanya otomatis.
- Di Windows butuh Docker Desktop. Kalau server dashboard-nya PC Windows biasa, cara Scheduled Task yang sekarang tetap lebih sederhana.

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

## 6. Kalau ada masalah

| Gejala | Solusi |
|---|---|
| Rekan tidak bisa buka `http://IP:8080` | 1) cek server hidup: di PC server buka `http://127.0.0.1:8080`<br>2) cek firewall: `install_service.bat` ulang<br>3) cek PC server & rekan satu jaringan (WiFi kantor yang sama) |
| Halaman upload error "token salah" | Isi kolom token di halaman upload (dikunci di browser setelah diisi) |
| Dashboard masih data lama | Tunggu 60 detik; cek `hps_server.log`; klik tombol "Perbarui" di halaman upload |
| `data.js tidak ditemukan` | Jalankan `update_data.bat` di folder server (butuh `pip install openpyxl`) |
| Gagal upload: "file bukan Excel valid" | File bukan .xlsx asli (atau .xls) — export ulang dari Excel sebagai .xlsx |
| Error Python | `pip install openpyxl` (butuh internet **sekali saja** saat installing) |
