# Copot service Dashboard HPS (task terjadwal + firewall + proses server)
$TaskName = "HPS Dashboard Server"
$ErrorActionPreference = "SilentlyContinue"

Write-Host ""
Write-Host " Menghapus service '$TaskName' ..."
if (Get-ScheduledTask -TaskName $TaskName) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "  service dicopot"
} else {
  Write-Host "  service tidak ada (sudah bersih)"
}

Get-NetFirewallRule -DisplayName "Dashboard HPS TCP 8080" | Remove-NetFirewallRule
Write-Host "  firewall dibersihkan"

$procs = Get-CimInstance Win32_Process | Where-Object { $_.Name -like "python*" -and $_.CommandLine -like "*server.py*" }
foreach ($p in $procs) {
  Stop-Process -Id $p.ProcessId -Force
  Write-Host "  server dihentikan (PID $($p.ProcessId))"
}
Write-Host ""
Write-Host " Selesai."
Write-Host ""
