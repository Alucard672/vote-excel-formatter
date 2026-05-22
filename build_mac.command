#!/bin/bash
set -e
cd "$(dirname "$0")"

python3 -m pip install -r requirements.txt
python3 -m PyInstaller --noconfirm --onefile --windowed --name "Excel整理" vote_excel_formatter.py

echo
echo "打包完成：$(pwd)/dist/Excel整理"
