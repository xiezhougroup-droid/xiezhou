@echo off
chcp 65001 >nul
cd /d "%~dp0"

set PORT=8876

where python >nul 2>nul
if %errorlevel%==0 (
  set PYTHON=python
) else (
  where py >nul 2>nul
  if %errorlevel%==0 (
    set PYTHON=py -3
  ) else (
    echo 未找到 Python。请先安装 Python 3。
    pause
    exit /b 1
  )
)

where ssh >nul 2>nul
if not %errorlevel%==0 (
  echo 未找到 ssh，无法启动公网临时演示。
  echo Windows 10/11 通常可以在“可选功能”里安装 OpenSSH Client。
  pause
  exit /b 1
)

echo 正在启动采购询价智能助手本地服务：http://127.0.0.1:%PORT%
start "采购询价本地服务" /B %PYTHON% web_app.py --no-open
timeout /t 2 /nobreak >nul

echo.
echo 正在创建公网临时演示链接...
echo 如果下方出现 https:// 开头的网址，把它发给同事即可访问。
echo 关闭本窗口后，公网链接会失效。
echo.

ssh -o ServerAliveInterval=60 -o StrictHostKeyChecking=no -R 80:localhost:%PORT% nokey@localhost.run

pause
