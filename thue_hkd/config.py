"""Tham số & bảng quy tắc dùng chung.

Mọi quy tắc trước đây phải sửa trong công thức Excel (ánh xạ TK, phân loại,
điều kiện lọc...) đều nằm ở đây. Có thể ghi đè bằng file JSON (xem
``cau_hinh_mau.json``) mà không cần sửa code.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path


@dataclass
class Config:
    # --- Tỷ lệ thuế khoán HKD ---
    ty_le_gtgt: float = 0.01        # Thuế GTGT 1%
    ty_le_tncn: float = 0.005       # Thuế TNCN 0.5%

    # --- Người phụ trách theo nguồn ---
    nguoi_phu_trach_dt: str = "Dung"
    nguoi_phu_trach_cp: str = "MINH HIEN"
    nguoi_phu_trach_kho: str = "Cuong"

    # --- Tên nguồn (cột "Nguồn" ở sheet TỔNG HỢP) ---
    nguon_dt: str = "Tổ DT"
    nguon_cp: str = "Tổ CP"
    nguon_kho: str = "Cường (13887)"

    # --- Loại KH được tính thuế (so khớp theo tiền tố) ---
    loai_kh_chiu_thue: list = field(default_factory=lambda: ["7. HỘ KINH DOANH", "2. HIỆU THUỐC"])
    loai_kh_khong_chiu: str = "KHÔNG CHỊU VAT/PIT (không phải HKD)"

    # --- Ánh xạ TK Nợ -> TK Có (XUẤT KM, VOUCHER) ---
    tk_no_sang_tk_co: dict = field(default_factory=lambda: {
        "1312": "51185",
        "33871": "51185",
        "6418": "15614",
        "335": "15614",
        "33412": "33311",
    })

    # --- Bảng phân loại giao dịch: [TK Nợ, TK Có ("*" = bất kỳ), Phân loại] ---
    # Duyệt từ trên xuống, dòng khớp đầu tiên được dùng.
    bang_phan_loai: list = field(default_factory=lambda: [
        ["335", "15614", "Quà tặng KH (CP)"],
        ["335", "*", "CKTT"],
        ["635", "*", "CKTT"],
        ["1312", "1311", "Chi TM TB KM"],
        ["1312", "51185", "Quà TB KM"],
        ["33871", "51185", "Quà TB CLB"],
        ["6418", "15614", "Quà tặng KH (CP)"],
        ["33412", "33311", "Quà Xuất tặng NV"],
        ["52111", "*", "Chi TM TB CLB"],
        ["52112", "*", "Chi TM TB CLB"],
    ])

    # --- Phân loại có cộng thêm VAT (giá thị trường = giá chưa thuế + VAT) ---
    phan_loai_cong_vat: list = field(default_factory=lambda: ["Quà TB CLB", "Quà TB KM"])

    # --- Điều kiện lọc TK Nợ cho XUẤT KM / VOUCHER ---
    tk_no_loc_xuat: list = field(default_factory=lambda: ["1312", "33871", "6418", "335"])

    # --- Sổ quỹ (DATA CẤN TRỪ / DATA_CHI): tiền tố phiếu -> (TK Có, Ghi chú) ---
    tien_to_phieu: dict = field(default_factory=lambda: {
        "BTH": ["1311", "Cấn trừ"],
        "PCN": ["1121", "CHI NH"],
    })

    # --- Dữ liệu kho (phần Cường): mã "Loại" -> PL KM ---
    loai_kho: dict = field(default_factory=lambda: {
        "0": "Xuất Bán",
        "1": "Xuất Khuyến Mãi",
        "10": "Xuất Tặng",
    })
    pl_km_khong_tinh: list = field(default_factory=lambda: ["Xuất Bán"])
    ma_kho: str = "NTG01"           # để trống = lấy mọi kho

    # --- Lọc theo kỳ (tháng/năm lấy từ NGÀY CHỨNG TỪ / NGÀY HÓA ĐƠN) ---
    loc_ngay_can_tru: bool = True
    loc_ngay_to_cp: bool = True
    # Cột ngày dùng để xác định kỳ: thử lần lượt, cột đầu tiên có dữ liệu ngày được dùng.
    cot_ngay: dict = field(default_factory=lambda: {
        "CAN_TRU": ["Ngày chứng từ", "Ngày ghi sổ"],
        "CHI": ["Ngày chứng từ", "Ngày ghi sổ"],
        "XUAT_KM": ["Ngay Hoa Don", "Ngay CT", "NgayGS"],
        "VOUCHER": ["Ngày chứng từ", "Ngày ghi sổ"],
        "KHO": ["Ngày hóa đơn", "Ngày giao hàng"],
    })

    # --- Cặp [TK Nợ, TK Có] chắc chắn KHÔNG tính thuế (không báo ở LƯU Ý) ---
    # 6418 / 1121: chi mời cơm sinh nhật KH (Tổ CP) – không tính thuế.
    # 33882 / *  : chi lại tiền CKTT (đã đưa vào TK 33882) – không tính thuế.
    # TK Có "*" = bất kỳ.
    cap_tk_khong_tinh: list = field(default_factory=lambda: [["6418", "1121"], ["33882", "*"]])

    @property
    def ty_le_tong(self) -> float:
        return self.ty_le_gtgt + self.ty_le_tncn

    # ------------------------------------------------------------------
    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, path_or_text: str | Path | None) -> "Config":
        cfg = cls()
        if not path_or_text:
            return cfg
        text = str(path_or_text)
        if not text.lstrip().startswith("{"):
            text = Path(text).read_text(encoding="utf-8")
        data = json.loads(text)
        for k, v in data.items():
            if hasattr(cfg, k):
                setattr(cfg, k, v)
        return cfg
