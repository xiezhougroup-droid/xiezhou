@echo off
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "web_app.py" --lan
  pause
  exit /b
)

where python >nul 2>nul
if %errorlevel%==0 (
  python "web_app.py" --lan
  pause
  exit /b
)

echo 未检测到 Python。
echo 请先安装 Python 3.11 或以上版本，并勾选 Add Python to PATH。
pause
