"""Quy trình đầy đủ: nhiều file -> nhận diện -> gộp -> tính -> kết quả."""
from __future__ import annotations

import calendar
from datetime import date
from pathlib import Path
from typing import Iterable

import pandas as pd

from . import doc_file as L
from .config import Config
from .tinh_toan import CanhBao, KetQua, ky_tu_du_lieu, tinh

DUOI_HO_TRO = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv", ".ods")


def ky_thang(thang: int, nam: int) -> tuple[date, date]:
    return date(nam, thang, 1), date(nam, thang, calendar.monthrange(nam, thang)[1])


def liet_ke_file(duong_dan: Iterable[str | Path]) -> list[Path]:
    """Nhận file hoặc thư mục (quét đệ quy). Bỏ file tạm của Excel (~$...)."""
    kq: list[Path] = []
    for p in map(Path, duong_dan):
        if p.is_dir():
            kq += sorted(x for x in p.rglob("*") if x.suffix.lower() in DUOI_HO_TRO and not x.name.startswith("~$"))
        elif p.suffix.lower() in DUOI_HO_TRO:
            kq.append(p)
    return kq


def doc_nhieu_file(nguon: Iterable, ghi_de_loai: dict[str, str] | None = None
                   ) -> tuple[list[L.BangDuLieu], list[str]]:
    """``nguon``: đường dẫn hoặc (tên, file-like). Trả về (bảng, lỗi)."""
    bangs: list[L.BangDuLieu] = []
    loi: list[str] = []
    for x in nguon:
        ten, f = (x if isinstance(x, tuple) else (Path(x).name, x))
        try:
            ds = L.doc_file(f, ten)
        except Exception as e:  # noqa: BLE001 - báo lỗi file cho người dùng
            loi.append(f"{ten}: không đọc được ({e})")
            continue
        if not ds:
            loi.append(f"{ten}: không nhận diện được loại dữ liệu (không thấy dòng tiêu đề quen thuộc)")
        bangs += ds
    if ghi_de_loai:
        for b in bangs:
            b.loai = ghi_de_loai.get(b.mo_ta, b.loai)
    return bangs, loi


def bang_nguon(bangs: list[L.BangDuLieu]) -> pd.DataFrame:
    return pd.DataFrame([{"File": b.ten_file, "Sheet": b.sheet, "Loại dữ liệu": L.TEN_LOAI.get(b.loai, b.loai),
                          "Số dòng": len(b.df), "Dòng tiêu đề": b.dong_tieu_de + 1} for b in bangs],
                        columns=["File", "Sheet", "Loại dữ liệu", "Số dòng", "Dòng tiêu đề"])


def chay(nguon: Iterable, tu: date | None = None, den: date | None = None, cfg: Config | None = None,
         ghi_de_loai: dict[str, str] | None = None) -> tuple[KetQua, list[str]]:
    """Không truyền ``tu``/``den`` thì kỳ = tháng có nhiều chứng từ nhất trong dữ liệu."""
    cfg = cfg or Config()
    bangs, loi = doc_nhieu_file(nguon, ghi_de_loai)
    bangs = [b for b in bangs if b.loai != L.KHONG_RO]
    trung: list = []
    du_lieu = L.gop_theo_loai(bangs, trung)
    if tu is None or den is None:
        ky = ky_tu_du_lieu(du_lieu, cfg)
        if ky is None:
            raise ValueError("Không xác định được tháng/năm từ dữ liệu – hãy chọn kỳ")
        tu, den = ky_thang(*ky)
    kq = tinh(du_lieu, tu, den, cfg)
    for loai, f, n, voi in trung:
        kq.canh_bao.insert(0, CanhBao("TRÙNG", f"{L.TEN_LOAI.get(loai, loai)}: {f} có {n:,} dòng giống hệt dòng ở {voi}",
                                      n, "Đã bỏ các dòng trùng để không cộng hai lần"))
    kq.nguon = bang_nguon(bangs)
    return kq, loi
