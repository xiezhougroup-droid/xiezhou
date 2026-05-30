#!/bin/zsh
set -e
cd "$(dirname "$0")"

CODEX_PYTHON="/Users/houzhou/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

if [ -x "$CODEX_PYTHON" ]; then
  "$CODEX_PYTHON" web_app.py
elif command -v python3 >/dev/null 2>&1; then
  python3 web_app.py
elif command -v python >/dev/null 2>&1; then
  python web_app.py
else
  echo "未检测到 Python。请先安装 Python 3.11 或以上版本。"
  read -p "按回车键退出..."
fi
