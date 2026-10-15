@echo off
setlocal
chcp 65001 >nul
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_vocab.ps1"
set "RESULT=%ERRORLEVEL%"
endlocal & exit /b %RESULT%
