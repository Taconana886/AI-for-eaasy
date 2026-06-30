@echo off
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
  python app.py
) else (
  "C:\Users\HUAWEI\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" app.py
)
pause
