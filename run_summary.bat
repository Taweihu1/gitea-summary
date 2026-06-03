@echo off
setlocal

chcp 65001 >nul

set HTTP_PROXY=http://cache2:3128
set HTTPS_PROXY=http://cache2:3128
set http_proxy=http://cache2:3128
set https_proxy=http://cache2:3128
set NO_PROXY=localhost,127.0.0.1
set no_proxy=localhost,127.0.0.1

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

rem -- Always restart Chrome (kills existing, clears locks, relaunches fresh)
python "%SCRIPT_DIR%launch_chrome_cdp.py"

rem -- Wait for CDP (up to ~30s)
echo   Waiting for Chrome CDP to come up (up to ~30s)...
set /a WAIT_TRIES=0
:wait_chrome
timeout /t 2 /nobreak >nul
curl -s --noproxy localhost --max-time 2 http://localhost:9222/json/version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo   Chrome CDP ready.
    goto run_script
)
set /a WAIT_TRIES+=1
if %WAIT_TRIES% LSS 15 goto wait_chrome

echo.
echo   [ERROR] Chrome did not expose CDP on port 9222 within 30s.
exit /b 1

:run_script
echo.
python "%SCRIPT_DIR%gitea_summary.py" %*

endlocal
exit /b %ERRORLEVEL%