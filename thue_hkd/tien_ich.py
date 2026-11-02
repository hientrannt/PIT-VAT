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


def ngay(x) -> date | None:
    """Đọc ngày từ datetime / số serial Excel / chuỗi dd/mm/yyyy, yyyy/mm/dd, yyyy-mm-dd."""
    if la_rong(x):
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    if isinstance(x, pd.Timestamp):
        return x.date()
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        if 1 < x < 2958466:
            return _EXCEL_EPOCH + timedelta(days=int(x))
        return None
    s = str(x).strip().split(" ")[0].split("T")[0]
    m = re.fullmatch(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.fullmatch(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", s)
        if not m:
            if re.fullmatch(r"\d+(\.\d+)?", s):
                return ngay(float(s))
            return None
        d, mo, y = map(int, m.groups())
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def bo_lap_doi(s: str) -> str:
    """'Võ CaVõ Ca' -> 'Võ Ca' (lỗi nối chuỗi 2 lần thường gặp trong HSKH)."""
    if not s or len(s) < 4 or len(s) % 2:
        return s
    h = len(s) // 2
    return s[:h] if s[:h] == s[h:] else s
