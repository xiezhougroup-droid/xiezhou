#!/bin/zsh
set -e
cd "$(dirname "$0")/.."
/Users/houzhou/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 采购询价本地工作流/src/analyze_supplier_quotes.py "$@"
