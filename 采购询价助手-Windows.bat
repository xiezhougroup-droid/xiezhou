@echo off
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "采购询价助手.py"
  pause
  exit /b
)

where python >nul 2>nul
if %errorlevel%==0 (
  python "采购询价助手.py"
  pause
  exit /b
)

echo 未检测到 Python。
echo 请先安装 Python 3.11 或以上版本，并勾选 Add Python to PATH。
echo 安装完成后重新双击本文件。
pause
