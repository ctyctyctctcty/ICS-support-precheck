@echo off
setlocal

set "APP_DIR=%~dp0"
set "APP_URL=http://127.0.0.1:8010"

cd /d "%APP_DIR%"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$url = '%APP_URL%';" ^
  "$ready = $false;" ^
  "try { $r = Invoke-WebRequest -Uri ($url + '/api/health') -UseBasicParsing -TimeoutSec 2; $ready = ($r.StatusCode -eq 200) } catch { }" ^
  "if (-not $ready) {" ^
  "  Start-Process -FilePath py -ArgumentList 'start_web.py' -WorkingDirectory '%APP_DIR%' -WindowStyle Minimized;" ^
  "  for ($i = 0; $i -lt 25; $i++) {" ^
  "    Start-Sleep -Milliseconds 400;" ^
  "    try { $r = Invoke-WebRequest -Uri ($url + '/api/health') -UseBasicParsing -TimeoutSec 2; if ($r.StatusCode -eq 200) { $ready = $true; break } } catch { }" ^
  "  }" ^
  "}" ^
  "Start-Process $url;" ^
  "if (-not $ready) { Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Web service is still starting, or failed to start. If the page does not open, run: py start_web.py', 'ICS Support Precheck') }"

endlocal
