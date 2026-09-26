@echo off
chcp 65001 >nul
powershell -NoProfile -Command "$c = Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue; if ($c) { $c | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }; Write-Host 'Da tat engine soat HSTT' } else { Write-Host 'Engine dang khong chay' }"
timeout /t 3 >nul
