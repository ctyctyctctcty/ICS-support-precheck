@echo off
setlocal

set "SCRIPT=%~dp0server\uninstall_startup_task.ps1"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent());" ^
  "$admin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);" ^
  "if ($admin) { & '%SCRIPT%' } else { Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT%""' }"

endlocal
