@echo off
REM ============================================================
REM  Sistem Absensi - public hosting launcher (Cloudflare Tunnel)
REM  Starts: Laravel on 127.0.0.1:8000  +  public HTTPS tunnel
REM ============================================================
setlocal
set APP=C:\xampp\htdocs\Sistem-Absensi-main\Sistem-Absensi-main
set CF=C:\Users\zohar\AppData\Local\hermes\cache\cloudflared.exe
set LOGDIR=%APP%\storage\logs

echo [1/2] Starting Laravel on 127.0.0.1:8000 ...
start "absensi-laravel" /MIN cmd /c "cd /d %APP% && php artisan serve --host=127.0.0.1 --port=8000 >> %LOGDIR%\serve.log 2>&1"

timeout /t 6 /nobreak >nul

echo [2/2] Starting Cloudflare Tunnel (public HTTPS URL) ...
start "absensi-tunnel" /MIN cmd /c ""%CF%" tunnel --url http://127.0.0.1:8000 --no-autoupdate >> %LOGDIR%\tunnel.log 2>&1"

timeout /t 12 /nobreak >nul

echo.
echo ============================================================
echo  PUBLIC URL (share this):
findstr /C:"trycloudflare.com" "%LOGDIR%\tunnel.log"
echo ============================================================
echo.
echo Keep this PC running for the site to stay online.
echo.
endlocal
