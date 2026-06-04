#!/bin/bash
set -e

cd "$(dirname "$0")/.."

echo "======================================"
echo "本科毕业设计论文格式检测工具"
echo "======================================"
echo

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3。请先安装 Python 3："
  echo "https://www.python.org/downloads/macos/"
  echo
  read -r -p "按回车退出..."
  exit 1
fi

if [ ! -d ".venv" ]; then
  echo "首次运行：正在创建本地 Python 环境..."
  python3 -m venv .venv
fi

source ".venv/bin/activate"
python -m pip install --upgrade pip >/dev/null
python -m pip install -r requirements.txt

echo
echo "请选择要检测的 .docx 文件..."
DOCX_PATH="$(osascript -e 'POSIX path of (choose file with prompt "请选择要检测的 Word 论文（.docx）")' 2>/dev/null || true)"

if [ -z "$DOCX_PATH" ]; then
  echo
  echo "如果没有弹出选择框，请把要检测的 .docx 文件拖到这个窗口，然后按回车："
  read -r DOCX_PATH
  DOCX_PATH="${DOCX_PATH%\"}"
  DOCX_PATH="${DOCX_PATH#\"}"
  DOCX_PATH="${DOCX_PATH%\'}"
  DOCX_PATH="${DOCX_PATH#\'}"
  DOCX_PATH="${DOCX_PATH//\\ / }"
fi

if [ ! -f "$DOCX_PATH" ]; then
  echo
  echo "没有找到文件：$DOCX_PATH"
  echo "请确认拖入的是 .docx 文件。"
  read -r -p "按回车退出..."
  exit 1
fi

BASE_NAME="$(basename "$DOCX_PATH" .docx)"
OUT_DIR="$HOME/Desktop/thesis_format_audit_reports"
mkdir -p "$OUT_DIR"
OUT_FILE="$OUT_DIR/${BASE_NAME}_format_audit_report.html"

echo
echo "正在检测：$DOCX_PATH"
python thesis_format_audit.py "$DOCX_PATH" --out "$OUT_FILE" || true

echo
echo "检测报告已生成："
echo "$OUT_FILE"
open "$OUT_FILE"
echo
read -r -p "检测完成，按回车关闭窗口..."
