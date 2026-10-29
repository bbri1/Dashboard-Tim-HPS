# Pasang penjadwal: ambil data HPS & refresh dashboard setiap 1 jam (untuk user yang sedang login).
$ErrorActionPreference = "Stop"
$name = "HPS Refresh Hourly"
$dir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$py   = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $py) { $py = "C:\Python314\pythonw.exe" }
$action  = New-ScheduledTaskAction -Execute $py -Argument ('"{0}\refresh_hourly.py"' -f $dir) -WorkingDirectory $dir
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Hours 1)
$set     = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger -Settings $set -Description "Dashboard HPS - ambil data & refresh tiap 1 jam" | Out-Null
Write-Host "Terpasang: $name (tiap 1 jam). Log: $dir\hps_refresh.log"
