"""Chạy bằng dòng lệnh.

Ví dụ:
    python cli.py --thang 9 --nam 2026 "D:/Thue HKD/T9" -o "D:/Thue HKD/KetQua_T9.xlsx"
    python cli.py --tu 2026-09-01 --den 2026-09-30 file1.xlsx file2.xlsb ...
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

from thue_hkd import Config, chay, ghi_excel, ky_thang, liet_ke_file


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Tập hợp & tính thuế Hộ kinh doanh từ nhiều file xuất")
    ap.add_argument("duong_dan", nargs="+", help="File hoặc thư mục chứa các file xuất từ phần mềm")
    ap.add_argument("--thang", type=int)
    ap.add_argument("--nam", type=int)
    ap.add_argument("--tu", type=date.fromisoformat, help="Từ ngày (yyyy-mm-dd)")
    ap.add_argument("--den", type=date.fromisoformat, help="Đến ngày (yyyy-mm-dd)")
    ap.add_argument("--cau-hinh", help="File cấu hình .json (tùy chọn)")
    ap.add_argument("-o", "--output", help="File kết quả .xlsx")
    a = ap.parse_args(argv)

    if a.tu and a.den:
        tu, den = a.tu, a.den
    elif a.thang and a.nam:
        tu, den = ky_thang(a.thang, a.nam)
    else:
        ap.error("Cần --thang và --nam, hoặc --tu và --den")

    cfg = Config.from_json(a.cau_hinh)
    files = liet_ke_file(a.duong_dan)
    if not files:
        print("Không tìm thấy file Excel/CSV nào.", file=sys.stderr)
        return 1
    print(f"Đọc {len(files)} file ...")
    kq, loi = chay(files, tu, den, cfg)
    for l in loi:
        print("  ! " + l)
    print(kq.nguon.to_string(index=False))
    print()
    for nhan, df in ((cfg.nguon_dt, kq.tong_hop_dt), (cfg.nguon_cp, kq.to_cp), (cfg.nguon_kho, kq.kho),
                     ("TỔNG CỘNG", kq.tat_ca)):
        t = kq.tong(df)
        print(f"{nhan:<16} {len(df):>7} dòng | Kê khai {t['Số tiền Kê khai']:>18,.0f} | "
              f"GTGT {t['Tiền thuế GTGT']:>14,.0f} | TNCN {t['Tiền thuế TNCN']:>14,.0f}")
    print("\nKIỂM TRA DỮ LIỆU:")
    for r in kq.chan_doan.itertuples(index=False):
        so_dong = "" if r[2] is None or r[2] != r[2] else f"{int(r[2]):>8,}"
        print(f"  {r[0]:<28} {r[1]:<70} {so_dong}")
    if kq.canh_bao:
        print("\nLƯU Ý:")
        for cb in kq.canh_bao:
            print(f"  [{cb.muc}] {cb.noi_dung} ({cb.so_dong} dòng) → {cb.hau_qua}")
    out = Path(a.output or f"KetQua_ThueHKD_{tu:%Y%m%d}_{den:%Y%m%d}.xlsx")
    ghi_excel(kq, out, cfg)
    print(f"\nĐã ghi: {out.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
