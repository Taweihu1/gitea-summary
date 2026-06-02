@echo off
set HTTP_PROXY=http://cache2:3128
set HTTPS_PROXY=http://cache2:3128
cd /d C:\Claude\gitea-summary
python gitea_summary.py
exit /b %ERRORLEVEL%
