@echo off
set HTTP_PROXY=http://cache2:3128
set HTTPS_PROXY=http://cache2:3128
cd /d C:\Claude\gitea-summary

rem ── Launch Brave with CDP if not already running on port 9222 ──────────────
powershell -command "try { (New-Object Net.Sockets.TcpClient('localhost', 9222)).Close(); exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Starting Brave with remote debugging port 9222 ...
    start "" "C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe" --remote-debugging-port=9222 --no-first-run --no-default-browser-check
    rem Wait for Brave to be ready
    :wait_loop
    timeout /t 2 /nobreak >nul
    powershell -command "try { (New-Object Net.Sockets.TcpClient('localhost', 9222)).Close(); exit 0 } catch { exit 1 }" >nul 2>&1
    if %ERRORLEVEL% NEQ 0 goto wait_loop
    echo Brave is ready.
) else (
    echo Brave CDP already running on port 9222.
)

python gitea_summary.py
exit /b %ERRORLEVEL%
