"""Đọc file xuất từ phần mềm (xlsx/xlsm/xls/xlsb/csv), tự nhận diện loại dữ liệu
và gộp nhiều file cùng loại thành một bảng.
"""
from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd

# ----------------------------------------------------------------------------
# Loại dữ liệu
# ----------------------------------------------------------------------------
CAN_TRU = "CAN_TRU"      # Sổ quỹ - Tổ Doanh thu (DATA CẤN TRỪ)
CHI = "CHI"              # Sổ quỹ - Tổ CP (DATA_CHI)
XUAT_KM = "XUAT_KM"      # DATA XUẤT KM
VOUCHER = "VOUCHER"      # VOUCHER
SO_CAI = "SO_CAI"        # SỔ CÁI
HSKH = "HSKH"            # Hồ sơ KH đã chuẩn hóa (sheet HSKH)
HSKH_IMPORT = "HSKH_IMPORT"  # Hồ sơ KH xuất thô từ CRM
THUE = "THUE"            # Bảng thuế suất theo mã hàng
KHO = "KHO"              # Dữ liệu bán hàng theo kho (phần Cường)
KHONG_RO = "KHONG_RO"

TEN_LOAI = {
    CAN_TRU: "Sổ quỹ - Cấn trừ (Tổ DT)",
    CHI: "Sổ quỹ - Chi (Tổ CP)",
    XUAT_KM: "Xuất khuyến mãi",
    VOUCHER: "Voucher",
    SO_CAI: "Sổ cái",
    HSKH: "Hồ sơ KH (chuẩn hóa)",
    HSKH_IMPORT: "Hồ sơ KH (xuất CRM)",
    THUE: "Bảng thuế suất",
    KHO: "Dữ liệu kho (Cường)",
    KHONG_RO: "Không nhận diện được",
}


def chuan_hoa(s) -> str:
    """Chuẩn hóa tên cột: NFC, bỏ xuống dòng, gộp khoảng trắng, chữ thường."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFC", str(s))
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


# Chữ ký nhận diện: tất cả các cột phải có mặt trong dòng tiêu đề.
CHU_KY = [
    (KHO, ["mã chi nhánh", "mã khách hàng", "ngày giao hàng", "loại", "tổng tiền giá gốc", "vat %"]),
    (XUAT_KM, ["mshh", "sohd", "mskh dso", "ttien co vat", "ngaygs"]),
    (VOUCHER, ["phiếu xuất", "mshh", "mstk", "số tiền", "mã số"]),
    (CAN_TRU, ["phiếu chi hạch toán", "mstk", "số tiền", "mã số"]),
    (SO_CAI, ["số chứng từ", "tài khoản", "tkđư"]),
    (HSKH_IMPORT, ["tên người đại diện", "tên đơn vị", "mã số thuế gtgt"]),
    (HSKH, ["số", "tên khách hàng", "tên đơn vị", "mst", "cccd"]),
    (THUE, ["mshh", "tenhh"]),
]
SO_DONG_DO_TIEU_DE = 40


@dataclass
class BangDuLieu:
    """Một bảng dữ liệu đọc được từ 1 sheet của 1 file."""
    ten_file: str
    sheet: str
    loai: str
    df: pd.DataFrame
    dong_tieu_de: int = 0

    @property
    def mo_ta(self) -> str:
        return f"{self.ten_file} [{self.sheet}]"


def _engine(ten: str) -> str | None:
    ext = Path(ten).suffix.lower()
    if ext in (".xlsb", ".xls", ".xlsx", ".xlsm", ".ods"):
        return "calamine"
    return None


def _doc_raw(nguon, ten: str) -> dict[str, pd.DataFrame]:
    ext = Path(ten).suffix.lower()
    if ext in (".csv", ".txt"):
        data = nguon.read() if hasattr(nguon, "read") else Path(nguon).read_bytes()
        for enc in ("utf-8-sig", "cp1258", "latin-1"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        sep = ";" if text.count(";") > text.count(",") else ","
        df = pd.read_csv(io.StringIO(text), header=None, dtype=object, sep=sep, keep_default_na=False)
        return {"csv": df}
    return pd.read_excel(nguon, sheet_name=None, header=None, dtype=object, engine=_engine(ten))


def _nhan_dien(raw: pd.DataFrame) -> tuple[str, int]:
    n = min(len(raw), SO_DONG_DO_TIEU_DE)
    for i in range(n):
        cot = {chuan_hoa(v) for v in raw.iloc[i].tolist() if v is not None and str(v).strip() != ""}
        if len(cot) < 2:
            continue
        for loai, ky in CHU_KY:
            if all(k in cot for k in ky):
                return loai, i
    return KHONG_RO, -1


def _lay_bang(raw: pd.DataFrame, dong: int) -> pd.DataFrame:
    tieu_de = [chuan_hoa(v) for v in raw.iloc[dong].tolist()]
    # Cột trùng tên: giữ cột đầu, các cột sau thêm hậu tố __2, __3...
    dem: dict[str, int] = {}
    cot = []
    for i, t in enumerate(tieu_de):
        if not t:
            t = f"__cot{i}"
        if t in dem:
            dem[t] += 1
            t = f"{t}__{dem[t]}"
        else:
            dem[t] = 1
        cot.append(t)
    df = raw.iloc[dong + 1:].copy()
    df.columns = cot
    # bỏ dòng trống hoàn toàn
    df = df.replace({"": None})
    df = df.dropna(how="all")
    return df.reset_index(drop=True)


def doc_file(nguon: str | Path | BinaryIO, ten: str | None = None) -> list[BangDuLieu]:
    """Đọc 1 file, trả về danh sách bảng (mỗi sheet nhận diện được là 1 bảng)."""
    ten = ten or (Path(nguon).name if isinstance(nguon, (str, Path)) else getattr(nguon, "name", "file"))
    sheets = _doc_raw(nguon, ten)
    ket_qua = []
    for sheet, raw in sheets.items():
        if raw.empty:
            continue
        loai, dong = _nhan_dien(raw)
        if loai == KHONG_RO:
            continue
        df = _lay_bang(raw, dong)
        if loai == CAN_TRU and _la_file_chi(ten, sheet, df):
            loai = CHI
        ket_qua.append(BangDuLieu(ten, str(sheet), loai, df, dong))
    return ket_qua


def _la_file_chi(ten: str, sheet: str, df: pd.DataFrame) -> bool:
    """Sổ quỹ của Tổ CP: tên file/sheet chứa 'CHI' (không phải 'CẤN TRỪ'),
    hoặc phần lớn phiếu là phiếu chi (PC...)."""
    nhan = chuan_hoa(re.sub(r"[_\-.]+", " ", f"{ten} {sheet}"))
    if re.search(r"cấn ?trừ|can ?tru", nhan):
        return False
    if re.search(r"(^|[^a-zà-ỹ])(data ?chi|to ?cp|tổ ?cp)([^a-zà-ỹ]|$)", nhan):
        return True
    cot = "phiếu chi hạch toán"
    if cot in df.columns:
        tt = df[cot].dropna().astype(str).str[:2].str.upper()
        if len(tt) and (tt == "PC").mean() > 0.5:
            return True
    return False


# Cột nhận dạng một dòng (để phát hiện cùng dữ liệu xuất ra ở nhiều file)
COT_KHOA = {
    CAN_TRU: ["phiếu chi hạch toán", "ngày chứng từ", "mstk", "số tiền", "mã số", "diễn giải"],
    XUAT_KM: ["mshh", "solo", "sohd", "so ct", "mskh dso", "so luong", "ttien co vat", "promotion no_"],
    VOUCHER: ["phiếu xuất", "mshh", "ngày chứng từ", "mstk", "số tiền", "mã số", "diễn giải"],
    KHO: ["số chứng từ", "số hóa đơn", "thứ tự", "mã hàng hóa", "số lô", "số lượng", "tổng tiền giá gốc", "loại"],
}
COT_KHOA[CHI] = COT_KHOA[CAN_TRU]


def _khoa_dong(df: pd.DataFrame, loai: str) -> pd.Series:
    from .tien_ich import chuoi
    cot = [c for c in COT_KHOA.get(loai, []) if c in df.columns] or sorted(c for c in df.columns if not c.startswith("__"))
    if not cot:
        return pd.Series([""] * len(df), index=df.index)
    return df[cot].apply(lambda col: col.map(chuoi)).agg("\x1f".join, axis=1)


def gop_theo_loai(bangs: Iterable[BangDuLieu], trung: list | None = None) -> dict[str, pd.DataFrame]:
    """Gộp các bảng cùng loại (giữ thứ tự file), thêm cột __nguon_file.

    Dòng giống hệt một dòng đã có ở FILE KHÁC cùng loại (file xuất trùng/chồng ngày) bị bỏ để không
    cộng hai lần; ghi vào ``trung`` dạng (loại, file, số dòng bỏ, file trùng với).
    """
    nhom: dict[str, list[pd.DataFrame]] = {}
    da_co: dict[str, dict[str, str]] = {}
    for b in bangs:
        d = b.df.copy()
        if b.loai not in (HSKH, HSKH_IMPORT, THUE, SO_CAI):
            khoa = _khoa_dong(d, b.loai)
            truoc = da_co.setdefault(b.loai, {})
            bo = khoa.isin(truoc.keys())
            if bo.any() and trung is not None:
                voi = sorted({truoc[k] for k in khoa[bo]})
                trung.append((b.loai, b.mo_ta, int(bo.sum()), ", ".join(voi)))
            d = d[~bo.values]
            for k in khoa[~bo].unique():
                truoc.setdefault(k, b.mo_ta)
        d["__nguon_file"] = b.mo_ta
        nhom.setdefault(b.loai, []).append(d)
    return {k: pd.concat(v, ignore_index=True, sort=False) for k, v in nhom.items()}
