#!/bin/bash
# Nhấp đúp trên macOS: tính thuế cho thư mục dữ liệu (mặc định thư mục 2026_Thue HKD), mở file kết quả.
cd "$(dirname "$0")"
DATA="${1:-$HOME/SynologyDrive/KE TOAN TRUONG/ANNUAL DATA/03. KẾ TOÁN - CHỨNG TỪ/Năm 2026/2026_Thue HKD}"
if [ ! -d .venv ]; then
  echo "Lần đầu: đang cài thư viện..."
  python3 -m venv .venv && .venv/bin/pip install -q -r requirements.txt pytest || { echo "Cần cài Python 3 (https://www.python.org/downloads/)"; read; exit 1; }
fi
OUT="$DATA/KetQua_ThueHKD_$(date +%Y%m%d_%H%M).xlsx"
.venv/bin/python cli.py "$DATA" -o "$OUT" && open "$OUT"
read -p "Nhấn Enter để đóng..."
