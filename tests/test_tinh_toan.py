"""Kiểm thử bằng dữ liệu giả (không chứa dữ liệu khách hàng thật)."""
from datetime import date

import pandas as pd
import pytest

from thue_hkd import Config, chay, ghi_excel
from thue_hkd import doc_file as L
from thue_hkd.tien_ich import bo_lap_doi, chuoi, lam_tron, ngay, so

TU, DEN = date(2026, 9, 1), date(2026, 9, 30)
HKD, KHONG = "7. HỘ KINH DOANH", "KHÔNG CHỊU VAT/PIT (không phải HKD)"


# ----------------------------------------------------------------- tiện ích
def test_lam_tron_kieu_excel():
    assert lam_tron(2.5) == 3
    assert lam_tron(-2.5) == -3
    assert lam_tron(1439086.2944) == 1439086
    assert lam_tron(1472.08) == 1472


def test_doc_ngay_va_so():
    assert ngay("12/09/2026") == date(2026, 9, 12)
    assert ngay("2026/09/11") == date(2026, 9, 11)
    assert ngay(46277) == date(2026, 9, 12)
    assert so("1.234.567") == 1234567
    assert so("1,234,567.5") == 1234567.5
    assert chuoi(1312.0) == "1312"
    assert bo_lap_doi("Võ CaVõ Ca") == "Võ Ca"
    assert bo_lap_doi("Lan") == "Lan"


# ----------------------------------------------------------------- dữ liệu giả
def _ghi(tmp_path, ten, df, dong_tieu_de_thua=0):
    p = tmp_path / ten
    with pd.ExcelWriter(p) as xw:
        df.to_excel(xw, index=False, startrow=dong_tieu_de_thua)
    return p


@pytest.fixture
def bo_file(tmp_path):
    hskh = pd.DataFrame({
        "Số": ["KH1", "KH2", "KH3"],
        "Tên Khách Hàng": ["Nguyễn A", "Trần B", "Cty C"],
        "Tên Đơn Vị": ["NT An", "QT Bình", "CTY C"],
        "MST": ["111", "", "333"],
        "CCCD": ["0011", "0022", ""],
        "PLoai (theo VAT/TNCN)": [HKD, HKD, KHONG],
    })
    thue = pd.DataFrame({"Mshh": ["HH1", "HH2"], "Tenhh": ["Bánh", "Quà"], "Thuế": [0.08, 0.1]})
    so_cai = pd.DataFrame({"Người Nhập": ["x"] * 3, "Số Chứng Từ": ["a", "b", "c"],
                           "Ngày Hóa Đơn": ["HD1", "HD2", "HD3"], "Tài Khoản": ["33871", "6418", "1312"],
                           "TKĐƯ": ["51185", "15614", "51185"]})
    km_cols = dict(Mshh=["HH1", "HH2", "HH1", "HH1"], Sohd=["HD1", "HD2", "HD1", "HD9"],
                   **{"Mskh Dso": ["KH1", "KH2", "KH3", "KH1"], "TTT Thue": [100000, 200000, 50000, 10000],
                      "Ttien Co Vat": [100000, 216000, 50000, 10000],
                      "NgayGS": ["12/09/2026", "13/09/2026", "12/09/2026", "12/09/2026"],
                      "Madv": ["NTG"] * 4, "Maso Ctkm": ["CT1"] * 4, "Ten Ctkm": ["CLB"] * 4})
    km = pd.DataFrame(km_cols)
    voucher = pd.DataFrame({"Phiếu xuất": ["PXV1", "PXV2", "PXV3"], "Mshh": ["VC1", "VC1", "VC1"],
                            "Ngày ghi sổ": ["09/09/2026", "09/10/2026", "09/09/2026"], "MSDV": ["NTG"] * 3,
                            "mstk": ["6418", "6418", "33412"], "Số tiền": [500000, 500000, 500000],
                            "Mã số": ["KH2", "KH2", "NV9"], "Mã CTKM": ["SN"] * 3, "Tên CTKM": ["Sinh nhật"] * 3})
    so_quy = pd.DataFrame({"Phiếu chi\n hạch toán": ["BTH/1", "PTN/2", "BTH/3"],
                           "Ngày ghi sổ": ["21/09/2026"] * 3, "msdv": ["NTG"] * 3,
                           "mstk": ["1312", "1121", "1312"], "Số tiền": [1000000, 999, 2000000],
                           "Mã số": ["KH1", "KH1", "KH3"]})
    kho = pd.DataFrame({"Mã chi nhánh": ["NTG"] * 4, "Mã khách hàng": ["KH1", "KH1", "KH2", "KH3"],
                        "Tên khách hàng": ["Nguyễn A", "Nguyễn A", "Trần B", "Cty C"],
                        "Tên đơn vị": ["NT An", "NT An", "QT Bình", "CTY C"],
                        "Số hóa đơn": ["K1", "K1", "K2", "K3"],
                        "Ngày hóa đơn": ["2026/09/03"] * 4,
                        "Ngày giao hàng": ["05/09/2026", "05/09/2026", "30/09/2026", "05/09/2026"],
                        "Loại": [0, 1, 10, 1], "Mã hàng hóa": ["A", "B", "C", "D"],
                        "Tên hàng hóa": ["a", "b", "c", "d"], "Tổng tiền giá gốc": [600000, 48000, -1350000, 1000],
                        "VAT %": [5, 5, 5, 5], "Mã kho": ["NTG01"] * 4})
    return [
        _ghi(tmp_path, "HSKH.xlsx", hskh),
        _ghi(tmp_path, "Thue.xlsx", thue),
        _ghi(tmp_path, "SoCai.xlsx", so_cai, dong_tieu_de_thua=3),   # có dòng tiêu đề báo cáo phía trên
        _ghi(tmp_path, "KM_1.xlsx", km.iloc[:2]),                     # 1 loại chia nhiều file
        _ghi(tmp_path, "KM_2.xlsx", km.iloc[2:]),
        _ghi(tmp_path, "Voucher.xlsx", voucher),
        _ghi(tmp_path, "So quy can tru.xlsx", so_quy),
        _ghi(tmp_path, "Kho.xlsx", kho),
    ]


def test_nhan_dien_file(bo_file):
    loai = {p.name: L.doc_file(p)[0].loai for p in bo_file}
    assert loai == {"HSKH.xlsx": L.HSKH, "Thue.xlsx": L.THUE, "SoCai.xlsx": L.SO_CAI,
                    "KM_1.xlsx": L.XUAT_KM, "KM_2.xlsx": L.XUAT_KM, "Voucher.xlsx": L.VOUCHER,
                    "So quy can tru.xlsx": L.CAN_TRU, "Kho.xlsx": L.KHO}


def test_tinh_day_du(bo_file, tmp_path):
    kq, loi = chay(bo_file, TU, DEN)
    assert loi == []

    # Cấn trừ: chỉ BTH + HKD
    assert kq.can_tru["Số tiền Kê khai"].tolist() == [lam_tron(1000000 / 0.985)]
    assert kq.can_tru["Phân Loại"].tolist() == ["Chi TM TB KM"]

    # Xuất KM: KH3 không phải HKD bị loại, HD9 không có trong sổ cái bị loại
    km = kq.xuat_km
    assert km["Mã số KH"].tolist() == ["KH1", "KH2"]
    clb, cp = km.iloc[0], km.iloc[1]
    assert clb["Phân Loại"] == "Quà TB CLB" and clb["Số tiền VAT"] == 8000 and clb["Số tiền chi (giá thị trường đã có VAT)"] == 108000
    assert cp["Phân Loại"] == "Quà tặng KH (CP)" and cp["Số tiền VAT"] == 0 and cp["Số tiền chi (giá thị trường đã có VAT)"] == 200000
    assert cp["Số tiền sổ sách"] == 216000
    kk = lam_tron(108000 / 0.985)
    assert clb["Số tiền Kê khai"] == kk and clb["Tiền thuế GTGT"] == lam_tron(kk * 0.01)
    assert clb["Tiền thuế TNCN"] == lam_tron(kk * 0.005)
    assert clb["Mã số thuế (nếu có)"] == "111" and clb["Số CMND"] == "0011"

    # Voucher: tháng 10 bị loại theo kỳ; NV9 không có HSKH
    assert kq.xuat_vc["Số chứng từ"].tolist() == ["PXV1"]

    # Kho: chỉ PL=1 và không phải Xuất Bán
    assert kq.kho["Phân Loại"].tolist() == ["Xuất Khuyến Mãi", "Xuất Tặng"]
    assert kq.kho["Số tiền VAT"].tolist() == [2400, -67500]
    assert kq.kho["Số tiền Kê khai"].tolist() == [51168, -1439086]
    assert kq.kho["Tiền thuế GTGT"].tolist() == [512, -14391]
    assert kq.kho["Tiền thuế TNCN"].tolist() == [256, -7195]

    # Tổng hợp
    assert len(kq.tong_hop_dt) == 1 + 2 + 1
    assert len(kq.tat_ca) == len(kq.tong_hop_dt) + len(kq.kho)
    assert set(kq.tat_ca["Nguồn"]) == {"Tổ DT", "Cường (13887)"}
    assert kq.theo_kh["Số tiền Kê khai"].sum() == kq.tat_ca["Số tiền Kê khai"].sum()

    muc = {(c.muc, c.noi_dung.split(":")[0]) for c in kq.canh_bao}
    assert ("SỔ CÁI", "Xuất KM") in muc
    assert ("HSKH", "Voucher") in muc

    out = tmp_path / "kq.xlsx"
    ghi_excel(kq, out)
    sheets = pd.read_excel(out, sheet_name=None, engine="calamine")
    assert {"TỔNG HỢP", "THEO KHÁCH HÀNG", "LƯU Ý", "KHO (Cường)"} <= set(sheets)


def test_cau_hinh_them_tk(bo_file):
    cfg = Config()
    cfg.tk_no_loc_xuat = ["33871"]          # bỏ 6418 khỏi điều kiện lọc
    kq, _ = chay(bo_file, TU, DEN, cfg)
    assert kq.xuat_km["Mã số KH"].tolist() == ["KH1"]
    assert Config.from_json(cfg.to_json()).tk_no_loc_xuat == ["33871"]


@pytest.mark.parametrize("ten,ky_vong", [
    ("T09.2026_SoQuy_CAN_TRU.xlsx", L.CAN_TRU), ("So quy cấn trừ T9.xlsx", L.CAN_TRU),
    ("T10.2026_DATA_CHI.xlsx", L.CHI), ("So quy TO-CP thang 10.xlsx", L.CHI), ("Tổ CP T10.xlsx", L.CHI),
    ("SoQuy_T9.xlsx", L.CAN_TRU),
])
def test_ten_file_so_quy(tmp_path, ten, ky_vong):
    df = pd.DataFrame({"Phiếu chi hạch toán": ["BTH/1", "PTN/2"], "mstk": ["1312", "1121"],
                       "Số tiền": [1, 2], "Mã số": ["KH1", "KH1"]})
    p = _ghi(tmp_path, ten, df)
    assert L.doc_file(p)[0].loai == ky_vong


def test_tu_chon_ky_va_bo_trung(bo_file, tmp_path):
    import shutil
    ban_sao = tmp_path / "KM_1 ban sao.xlsx"
    shutil.copy(bo_file[3], ban_sao)
    kq, _ = chay(bo_file + [ban_sao])             # không truyền kỳ -> lấy từ ngày chứng từ / hóa đơn
    assert (kq.tu_ngay, kq.den_ngay) == (TU, DEN)
    assert kq.xuat_km["Mã số KH"].tolist() == ["KH1", "KH2"]   # dòng trùng không bị cộng 2 lần
    assert any(c.muc == "TRÙNG" for c in kq.canh_bao)
    assert kq.chan_doan is not None and len(kq.chan_doan)


@pytest.mark.parametrize("v,kq", [
    ("09/09/2026", date(2026, 9, 9)), ("9/9/2026", date(2026, 9, 9)), ("09/09/2026 00:00:00", date(2026, 9, 9)),
    ("09/09/2026 00:00", date(2026, 9, 9)), ("'09/09/2026", date(2026, 9, 9)), ("09-09-2026", date(2026, 9, 9)),
    ("09.09.2026", date(2026, 9, 9)), ("09/09/26", date(2026, 9, 9)), ("9-Sep-2026", date(2026, 9, 9)),
    ("Sep 9, 2026", date(2026, 9, 9)), ("20260909", date(2026, 9, 9)), ("46274", date(2026, 9, 9)),
    ("ngày 09 tháng 09 năm 2026", date(2026, 9, 9)), ("09/13/2026", date(2026, 9, 13)),
    ("12/09/2026 12:00:00 SA", date(2026, 9, 12)), ("01/10/2026", date(2026, 10, 1)),
    ("2026-09-11T00:00:00", date(2026, 9, 11)), ("abc", None), ("31/02/2026", None),
])
def test_ngay_dang_chu(v, kq):
    assert ngay(v) == kq
