@echo off
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server\status_startup_task.ps1"
pause

endlocal
