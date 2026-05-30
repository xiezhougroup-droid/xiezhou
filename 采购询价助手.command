#!/bin/zsh
set -e
cd "$(dirname "$0")"

if command -v python3 >/dev/null 2>&1; then
  python3 采购询价助手.py
elif command -v python >/dev/null 2>&1; then
  python 采购询价助手.py
else
  echo "未检测到 Python。请先安装 Python 3.11 或以上版本。"
  read -p "按回车键退出..."
fi
