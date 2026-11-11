"""Hàm tiện ích: làm sạch chuỗi, số, ngày và làm tròn kiểu Excel."""
from __future__ import annotations

import math
import re
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd

_EXCEL_EPOCH = date(1899, 12, 30)


def la_rong(x) -> bool:
    if x is None:
        return True
    if isinstance(x, float) and math.isnan(x):
        return True
    try:
        if x is pd.NaT or (not isinstance(x, (str, bytes)) and pd.isna(x)):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(x, str) and x.strip() == ""


def chuoi(x) -> str:
    """Chuỗi sạch: 1312.0 -> '1312', None -> ''."""
    if la_rong(x):
        return ""
    if isinstance(x, bool):
        return str(x)
    if isinstance(x, float):
        if x.is_integer():
            return str(int(x))
        return repr(x)
    if isinstance(x, int):
        return str(x)
    return str(x).strip()


def so(x) -> float:
    """Chuyển về số; chuỗi '1.234.567' / '1,234,567' đều đọc được; rỗng -> 0."""
    if la_rong(x):
        return 0.0
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return float(x)
    s = str(x).strip().replace(" ", "").replace(" ", "")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):      # 1.234.567,89
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", s):    # 1,234,567.89
        s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if neg else v


def lam_tron(x: float, chu_so: int = 0) -> float:
    """ROUND của Excel: làm tròn nửa xa số 0 (0.5 -> 1, -0.5 -> -1)."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return 0.0
    d = Decimal(repr(float(x)))
    q = Decimal(1).scaleb(-chu_so)
    return float(d.quantize(q, rounding=ROUND_HALF_UP))


_THANG_CHU = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8,
              "sep": 9, "oct": 10, "nov": 11, "dec": 12}


def _tao_ngay(y: int, mo: int, d: int) -> date | None:
    if y < 100:
        y += 2000
    if mo > 12 and d <= 12:          # mm/dd kiểu Mỹ (vd 09/13/2026)
        mo, d = d, mo
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def ngay(x) -> date | None:
    """Đọc ngày từ: datetime / số serial Excel / số yyyymmdd / chuỗi.

    Chuỗi chấp nhận: dd/mm/yyyy, d/m/yy, dd-mm-yyyy, dd.mm.yyyy, yyyy/mm/dd, yyyy-mm-ddThh:mm,
    có giờ phía sau (kể cả SA/CH), dấu ' ở đầu, khoảng trắng ẩn, 9-Sep-2026, Sep 9 2026,
    20260909, "ngày 09 tháng 09 năm 2026", số serial dạng chữ (46274). Mặc định ngày trước tháng.
    """
    if la_rong(x):
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    if isinstance(x, pd.Timestamp):
        return x.date()
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        if float(x).is_integer() and 19000101 <= x <= 21001231:
            v = int(x)
            return _tao_ngay(v // 10000, v // 100 % 100, v % 100)
        if 1 < x < 2958466:
            return _EXCEL_EPOCH + timedelta(days=int(x))
        return None
    s = str(x).replace("\u00a0", " ").replace("\u200b", "").strip().lstrip("'`").strip().lower()
    if not s:
        return None
    m = re.search(r"ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})", s)
    if m:
        return _tao_ngay(int(m[3]), int(m[2]), int(m[1]))
    m = re.search(r"(?<!\d)(\d{4})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})(?!\d)", s)
    if m:
        return _tao_ngay(int(m[1]), int(m[2]), int(m[3]))
    m = re.search(r"(?<!\d)(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{4}|\d{2})(?!\d)", s)
    if m:
        return _tao_ngay(int(m[3]), int(m[2]), int(m[1]))
    m = re.search(r"(?<!\d)(\d{1,2})[\s\-/.]*([a-z]{3})[a-z]*[\s\-/.,]*(\d{4}|\d{2})(?!\d)", s)
    if m and m[2] in _THANG_CHU:
        return _tao_ngay(int(m[3]), _THANG_CHU[m[2]], int(m[1]))
    m = re.search(r"([a-z]{3})[a-z]*\.?\s*(\d{1,2}),?\s*(\d{4})", s)
    if m and m[1] in _THANG_CHU:
        return _tao_ngay(int(m[3]), _THANG_CHU[m[1]], int(m[2]))
    if re.fullmatch(r"\d{8}", s):
        return ngay(int(s))
    if re.fullmatch(r"\d+(\.\d+)?", s):
        return ngay(float(s))
    return None


def bo_lap_doi(s: str) -> str:
    """'Võ CaVõ Ca' -> 'Võ Ca' (lỗi nối chuỗi 2 lần thường gặp trong HSKH)."""
    if not s or len(s) < 4 or len(s) % 2:
        return s
    h = len(s) // 2
    return s[:h] if s[:h] == s[h:] else s
