#!/bin/zsh
set -e

cd "$(dirname "$0")"

PORT=8876
CODEX_PYTHON="/Users/houzhou/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

if [ -x "$CODEX_PYTHON" ]; then
  PYTHON="$CODEX_PYTHON"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON=python3
elif command -v python >/dev/null 2>&1; then
  PYTHON=python
else
  echo "未找到 Python。请先安装 Python 3。"
  exit 1
fi

if ! command -v ssh >/dev/null 2>&1; then
  echo "未找到 ssh，无法启动公网临时演示。"
  exit 1
fi

echo "正在启动采购询价智能助手本地服务：http://127.0.0.1:${PORT}"
"$PYTHON" web_app.py --no-open &
WEB_PID=$!

cleanup() {
  kill "$WEB_PID" >/dev/null 2>&1 || true
}
trap cleanup EXIT

sleep 2

echo ""
echo "正在创建公网临时演示链接..."
echo "如果下方出现 https:// 开头的网址，把它发给同事即可访问。"
echo "关闭本窗口后，公网链接会失效。"
echo ""

ssh -o ServerAliveInterval=60 -o StrictHostKeyChecking=no -R 80:localhost:${PORT} nokey@localhost.run
