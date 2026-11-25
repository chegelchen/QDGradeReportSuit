#!/bin/bash
cd "$(dirname "$0")"

PYTHON_BIN=""
for py in \
  /usr/bin/python3 \
  "$(command -v python3 2>/dev/null)" \
  /opt/homebrew/bin/python3 \
  /usr/local/bin/python3 \
  /Library/Frameworks/Python.framework/Versions/*/bin/python3; do
  if [ -n "$py" ] && [ -x "$py" ]; then
    if "$py" -c "import openpyxl, pymupdf, PIL, tkinter" >/dev/null 2>&1; then
      PYTHON_BIN="$py"
      break
    fi
  fi
done

if [ -z "$PYTHON_BIN" ]; then
  echo "❌ 未找到包含 openpyxl / pymupdf / Pillow / tkinter 的 Python 3 环境。"
  echo "请先运行: python3 -m pip install openpyxl pymupdf Pillow"
  exit 1
fi

"$PYTHON_BIN" main.py
