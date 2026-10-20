# ============================================================
#  Pasang Dashboard HPS sebagai service di PC kantor (Windows)
#  - auto start saat PC dinyalakan / user login
#  - auto restart kalau server mati
#  - buka port firewall agar bisa diakses dari jaringan kantor
#  Jalankan: klik 2x install_service.bat  (atau: powershell -ExecutionPolicy Bypass -File install_service.ps1)
# ============================================================
$ErrorActionPreference = "Continue"
$TaskName = "HPS Dashboard Server"
$Port = 8080
$Dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

Write-Host ""
Write-Host " ============================================================"
Write-Host "  PASANG DASHBOARD HPS - SERVER JARINGAN KANTOR"
Write-Host " ============================================================"
Write-Host "  Folder   : $Dir"
Write-Host "  Port     : $Port"
Write-Host "  Admin    : $isAdmin"
Write-Host ""

# ---- 1. cari python ----
$py = (Get-Command python.exe -ErrorAction SilentlyContinue).Source
if (-not $py) { $py = (Get-Command python -ErrorAction SilentlyContinue).Source }
if (-not $py) {
  Write-Host "  [X] Python tidak ditemukan. Pasang Python 3 dulu (python.org), lalu jalankan ulang." -ForegroundColor Red
  exit 1
}
$pyw = Join-Path (Split-Path -Parent $py) "pythonw.exe"
if (-not (Test-Path $pyw)) { $pyw = $py }
Write-Host "  [1/5] Python  : $py"

# ---- 2. cek openpyxl ----
& $py -c "import openpyxl" 2>$null
if ($LASTEXITCODE -ne 0) {
  Write-Host "  [2/5] openpyxl belum ada, menginstall..." -ForegroundColor Yellow
  & $py -m pip install --quiet --upgrade openpyxl
  if ($LASTEXITCODE -ne 0) { Write-Host "  [!] openpyxl gagal diinstall (butuh internet sekali saja)." -ForegroundColor Yellow }
} else { Write-Host "  [2/5] openpyxl : sudah ada" }

# ---- 3. firewall ----
Write-Host "  [3/5] Membuka port $Port di firewall..."
try {
  if (Get-NetFirewallRule -DisplayName "Dashboard HPS TCP $Port" -ErrorAction SilentlyContinue) {
    Write-Host "        (aturan sudah ada)"
  } else {
    New-NetFirewallRule -DisplayName "Dashboard HPS TCP $Port" -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow -Profile Domain,Private -ErrorAction Stop | Out-Null
    Write-Host "        (rule baru dibuat)"
  }
} catch {
  Write-Host "        (!) Gagal membuat aturan firewall. Jalankan sebagai admin, atau tambahkan manual:" -ForegroundColor Yellow
  Write-Host "        New-NetFirewallRule -DisplayName 'Dashboard HPS TCP $Port' -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow -Profile Domain,Private" -ForegroundColor Yellow
}

# ---- 4. jadwalkan service ----
Write-Host "  [4/5] Mendaftarkan service '$TaskName'..."
$action = New-ScheduledTaskAction -Execute $pyw -Argument "`"$Dir\server.py`" $Port" -WorkingDirectory $Dir
$settings = New-ScheduledTaskSettingsSet -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero) -StartWhenAvailable -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId ("$env:USERDOMAIN\$env:USERNAME") -LogonType Interactive -RunLevel Limited
$triggers = @()
if ($isAdmin) { $triggers += New-ScheduledTaskTrigger -AtStartup }
$triggers += New-ScheduledTaskTrigger -AtLogOn -User ("$env:USERDOMAIN\$env:USERNAME")
try {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Description "Dashboard HPS - server lokal jaringan kantor" -Force | Out-Null
  Write-Host "        (service terdaftar: auto-start + auto-restart)"
} catch {
  Write-Host "        (!) Gagal mendaftarkan service: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ---- 5. jalankan sekarang ----
Write-Host "  [5/5] Menjalankan server sekarang..."
Get-Process python, pythonw -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $py -or $_.Path -eq $pyw } | ForEach-Object {
  try { $cl = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)").CommandLine } catch { $cl = "" }
  if ($cl -like "*server.py*") { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
}
Start-Process -FilePath $pyw -ArgumentList "`"$Dir\server.py`" $Port" -WorkingDirectory $Dir
Start-Sleep -Seconds 3

# ---- info akses ----
$ips = @()
try { $ips = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254*" } | Select-Object -ExpandProperty IPAddress) } catch {}
Write-Host ""
Write-Host " ============================================================"
Write-Host "  SELESAI. Dashboard HPS aktif di:"
Write-Host "    Komputer ini : http://127.0.0.1:$Port/"
foreach ($ip in $ips) { Write-Host "  Jaringan      : http://$ip`:$Port/" }
Write-Host ""
Write-Host "  Cara update data (tanpa laptop):"
foreach ($ip in $ips) { Write-Host "    - buka http://$ip`:$Port/upload.html di browser, lalu drag file Excel" }
Write-Host "    - atau salin file .xlsx ke folder ini (share jaringan), rebuild otomatis 20 detik"
Write-Host ""
Write-Host "  Uninstall: klik 2x uninstall_service.bat"
Write-Host " ============================================================"
Write-Host ""
