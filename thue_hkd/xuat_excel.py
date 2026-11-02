"""Ghi kết quả ra file Excel (.xlsx) có định dạng sẵn: tổng ở đầu, lọc, cố định tiêu đề."""
from __future__ import annotations

import io
from datetime import date
from pathlib import Path

import pandas as pd

from .config import Config
from .tinh_toan import COT_KHO, COT_TIEN, KetQua

COT_SO_KHO = ["VAT %", "TT giá gốc", "T.Tiền thuế", "TT tính thuế", "Số tiền chi",
              "Thuế 1%", "Thuế 0.5%", "Tổng thuế"]


class _Ghi:
    def __init__(self, wb):
        self.wb = wb
        self.f_tieu_de = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                                        "text_wrap": True, "valign": "vcenter", "border": 1})
        self.f_so = wb.add_format({"num_format": "#,##0;-#,##0;0"})
        self.f_ngay = wb.add_format({"num_format": "dd/mm/yyyy"})
        self.f_tong = wb.add_format({"bold": True, "num_format": "#,##0;-#,##0;0", "bg_color": "#FFF2CC"})
        self.f_nhan = wb.add_format({"bold": True, "bg_color": "#FFF2CC"})
        self.f_h1 = wb.add_format({"bold": True, "font_size": 14})
        self.f_dam = wb.add_format({"bold": True})
        self.f_wrap = wb.add_format({"text_wrap": True, "valign": "top"})

    def bang(self, ten: str, df: pd.DataFrame, cot_so=(), cot_ngay=(), dong_dau: int = 0,
             tong: bool = True, do_rong: dict | None = None):
        ws = self.wb.add_worksheet(ten[:31])
        r0 = dong_dau
        if tong:
            ws.write(r0, 0, "TỔNG (dòng đang lọc)", self.f_nhan)
            r0 += 1
        cols = list(df.columns)
        for j, c in enumerate(cols):
            ws.write(r0, j, c, self.f_tieu_de)
        n = len(df)
        for j, c in enumerate(cols):
            s = df[c].tolist()
            if c in cot_so:
                for i, v in enumerate(s):
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        continue
                    ws.write_number(r0 + 1 + i, j, float(v), self.f_so)
            elif c in cot_ngay:
                for i, v in enumerate(s):
                    if isinstance(v, date):
                        ws.write_datetime(r0 + 1 + i, j, pd.Timestamp(v).to_pydatetime(), self.f_ngay)
            else:
                for i, v in enumerate(s):
                    if v is None or v == "" or (isinstance(v, float) and pd.isna(v)):
                        continue
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        ws.write_number(r0 + 1 + i, j, v)
                    else:
                        ws.write_string(r0 + 1 + i, j, str(v))
        if tong:
            for j, c in enumerate(cols):
                if c in cot_so:
                    col = _ten_cot(j)
                    ws.write_formula(dong_dau, j, f"=SUBTOTAL(109,{col}{r0 + 2}:{col}{r0 + 1 + max(n, 1)})",
                                     self.f_tong, float(pd.to_numeric(df[c], errors="coerce").fillna(0).sum()))
        ws.autofilter(r0, 0, r0 + max(n, 1), max(len(cols) - 1, 0))
        ws.freeze_panes(r0 + 1, 2)
        ws.set_row(r0, 32)
        for j, c in enumerate(cols):
            w = (do_rong or {}).get(c)
            if w is None:
                w = 14 if c in cot_so or c in cot_ngay else min(max(len(str(c)), 10), 28)
            ws.set_column(j, j, w)
        return ws


def _ten_cot(j: int) -> str:
    s = ""
    j += 1
    while j:
        j, r = divmod(j - 1, 26)
        s = chr(65 + r) + s
    return s


RONG = {"Họ và tên KH": 24, "Tên Quầy Thuốc/Nhà Thuốc/Tổ chức": 36, "Tên CTKM": 36,
        "Ghi chú (Mã HH - Tên HH)": 36, "Số chứng từ": 20, "Mã CTKM": 20}


def ghi_excel(kq: KetQua, dich: str | Path | io.BytesIO | None = None, cfg: Config | None = None) -> bytes:
    cfg = cfg or Config()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        wb = xw.book
        g = _Ghi(wb)

        # --- Tổng hợp tất cả (tương đương sheet ALL) ---
        ws = g.bang("TỔNG HỢP", kq.tat_ca, COT_TIEN, ["Ngày ghi sổ"], dong_dau=5, do_rong=RONG)
        ws.write(0, 0, f"BẢNG TỔNG HỢP THUẾ HỘ KINH DOANH – từ {kq.tu_ngay:%d/%m/%Y} đến {kq.den_ngay:%d/%m/%Y}",
                 g.f_h1)
        ws.write(1, 0, "Nguồn", g.f_dam)
        ws.write(1, 1, "Số dòng", g.f_dam)
        for j, c in enumerate(COT_TIEN):
            ws.write(1, 6 + j, c, g.f_tieu_de)
        for i, (nhan, df) in enumerate(((cfg.nguon_dt, kq.tong_hop_dt), (cfg.nguon_cp, kq.to_cp),
                                        (cfg.nguon_kho, kq.kho))):
            ws.write(2 + i, 0, nhan)
            ws.write_number(2 + i, 1, len(df))
            for j, c in enumerate(COT_TIEN):
                ws.write_number(2 + i, 6 + j, float(df[c].sum()) if len(df) else 0, g.f_so)

        g.bang("THEO KHÁCH HÀNG", kq.theo_kh, COT_TIEN + ["Số dòng"], do_rong=RONG)
        g.bang("TONGHOP (Tổ DT)", kq.tong_hop_dt, COT_TIEN, ["Ngày ghi sổ"], do_rong=RONG)
        g.bang("CẤN TRỪ", kq.can_tru, COT_TIEN, ["Ngày ghi sổ"], do_rong=RONG)
        g.bang("XUẤT KM", kq.xuat_km, COT_TIEN, ["Ngày ghi sổ"], do_rong=RONG)
        g.bang("XUẤT VC", kq.xuat_vc, COT_TIEN, ["Ngày ghi sổ"], do_rong=RONG)
        g.bang("TỔ CP", kq.to_cp, COT_TIEN, ["Ngày ghi sổ"], do_rong=RONG)
        g.bang("KHO (Cường)", kq.kho_chi_tiet[COT_KHO] if len(kq.kho_chi_tiet) else kq.kho_chi_tiet,
               COT_SO_KHO, ["Ngày HĐ", "Ngày giao"], do_rong={"Tên KH": 22, "Tên đơn vị": 32, "Tên HH": 32})

        # --- Lưu ý / bất thường ---
        ws = wb.add_worksheet("LƯU Ý")
        ws.write(0, 0, "DANH SÁCH BẤT THƯỜNG CẦN XỬ LÝ (TỰ ĐỘNG)", g.f_h1)
        ws.set_column(0, 0, 14)
        ws.set_column(1, 1, 70)
        ws.set_column(2, 2, 10)
        ws.set_column(3, 3, 60)
        r = 2
        if not kq.canh_bao:
            ws.write(r, 0, "Không có bất thường")
        for cb in kq.canh_bao:
            ws.write(r, 0, cb.muc, g.f_tieu_de)
            ws.write(r, 1, cb.noi_dung, g.f_dam)
            ws.write(r, 2, cb.so_dong, g.f_so)
            ws.write(r, 3, cb.hau_qua, g.f_wrap)
            r += 1
            if cb.chi_tiet is not None and len(cb.chi_tiet):
                ct = cb.chi_tiet.head(500)
                for j, c in enumerate(ct.columns):
                    ws.write(r, 1 + j, c, g.f_dam)
                r += 1
                for row in ct.itertuples(index=False):
                    for j, v in enumerate(row):
                        if isinstance(v, (int, float)) and not isinstance(v, bool):
                            ws.write_number(r, 1 + j, v)
                        else:
                            ws.write(r, 1 + j, "" if v is None else str(v))
                    r += 1
            r += 1

        if kq.nguon is not None:
            g.bang("NGUỒN FILE", kq.nguon, ["Số dòng"], tong=False,
                   do_rong={"File": 50, "Sheet": 20, "Loại dữ liệu": 28})
        ws = wb.add_worksheet("THAM SỐ")
        ws.set_column(0, 0, 100)
        ws.write(0, 0, "Cấu hình đã dùng để tính (có thể lưu thành file .json để dùng lại):", g.f_dam)
        for i, line in enumerate(cfg.to_json().splitlines()):
            ws.write_string(i + 1, 0, line)
    data = buf.getvalue()
    if dich is not None:
        if isinstance(dich, io.BytesIO):
            dich.write(data)
        else:
            Path(dich).write_bytes(data)
    return data
