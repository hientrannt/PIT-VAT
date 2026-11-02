"""Phần mềm tập hợp & tính thuế Hộ kinh doanh (GTGT 1% + TNCN 0.5%)."""
from .config import Config
from .quy_trinh import chay, ky_thang, liet_ke_file
from .tinh_toan import KetQua, tinh
from .xuat_excel import ghi_excel

__all__ = ["Config", "KetQua", "chay", "ghi_excel", "ky_thang", "liet_ke_file", "tinh"]
__version__ = "1.0.0"
