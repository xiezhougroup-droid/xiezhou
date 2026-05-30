#!/bin/zsh
set -e
cd "$(dirname "$0")/.."

PY="/Users/houzhou/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

echo "步骤1/4：按样板标准生成采购包拆分总表"
"$PY" 采购询价本地工作流/src/build_split_from_standard.py "$@"

LATEST_SPLIT=$(ls -td 采购询价本地工作流/output/*_按昨日样板标准拆分 2>/dev/null | head -1)
SPLIT_FILE="$LATEST_SPLIT/采购包拆分总表-按昨日样板标准.xlsx"

echo "步骤2/4：根据拆分总表批量生成全部采购包邀请清单"
"$PY" 采购询价本地工作流/src/generate_invitations_from_split.py --standard "$SPLIT_FILE" --packages all

echo "步骤3/4：分析已回收的供应商报价"
"$PY" 采购询价本地工作流/src/analyze_supplier_quotes.py

echo "步骤4/4：根据采购员谈判反馈生成最终推荐结果"
"$PY" 采购询价本地工作流/src/finalize_recommendation.py
