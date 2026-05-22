#!/bin/bash
set -e
cd "$(dirname "$0")"

python3 -m pip install -r requirements.txt
python3 -m PyInstaller --noconfirm --onefile --windowed --name "投票Excel整理工具" vote_excel_formatter.py

echo
echo "打包完成：$(pwd)/dist/投票Excel整理工具"
