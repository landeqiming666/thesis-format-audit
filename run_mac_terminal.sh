#!/bin/bash
set -e

cd "$(dirname "$0")"

if [ "$#" -lt 1 ]; then
  echo "用法：./run_mac_terminal.sh 论文.docx [报告.html]"
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3，请先安装 Python 3。"
  exit 1
fi

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source ".venv/bin/activate"
python -m pip install --upgrade pip >/dev/null
python -m pip install -r requirements.txt

DOCX_PATH="$1"
if [ "$#" -ge 2 ]; then
  OUT_FILE="$2"
else
  BASE_NAME="$(basename "$DOCX_PATH" .docx)"
  OUT_DIR="$HOME/Desktop/thesis_format_audit_reports"
  mkdir -p "$OUT_DIR"
  OUT_FILE="$OUT_DIR/${BASE_NAME}_format_audit_report.html"
fi

python thesis_format_audit.py "$DOCX_PATH" --out "$OUT_FILE" || true
echo "报告已生成：$OUT_FILE"
