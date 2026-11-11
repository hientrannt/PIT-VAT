"""Bộ máy tính thuế HKD: thay thế toàn bộ công thức của file Excel
"Tập hợp Thuế Hộ Kinh Doanh" (DATA CẤN TRỪ, DATA XUẤT KM, VOUCHER, DATA_CHI,
ThueHkdKho01 -> TONGHOP / ALL).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

import pandas as pd

from . import doc_file as L
from .config import Config
from .tien_ich import bo_lap_doi, chuoi, lam_tron, ngay, so

COT_KQ = [
    "Người phụ trách", "Mã số KH", "Họ và tên KH", "Tên Quầy Thuốc/Nhà Thuốc/Tổ chức",
    "Mã số thuế (nếu có)", "Số CMND", "Số tiền sổ sách", "Số tiền VAT",
    "Số tiền chi (giá thị trường đã có VAT)", "Số tiền Kê khai", "Tiền thuế GTGT",
    "Tiền thuế TNCN", "Đơn vị", "Số chứng từ", "Ngày ghi sổ", "TK Nợ", "TK Có",
    "Phân Loại", "Mã CTKM", "Tên CTKM",
]
COT_TIEN = COT_KQ[6:12]
COT_ALL = COT_KQ + ["Nguồn", "Ghi chú (Mã HH - Tên HH)"]

COT_KHO = [
    "Mã CN", "Mã KH", "Tên KH", "Tên đơn vị", "Số HĐ", "Ngày HĐ", "Ngày giao", "Loại",
    "Mã HH", "Tên HH", "PL KM", "Chịu thuế", "PL", "VAT %", "TT giá gốc", "T.Tiền thuế",
    "TT tính thuế", "Số tiền chi", "Thuế 1%", "Thuế 0.5%", "Tổng thuế",
]


class ThieuCot(Exception):
    pass


def _cot(df: pd.DataFrame, *ten: str, bat_buoc: bool = True) -> pd.Series:
    """Lấy cột theo tên (đã chuẩn hóa); thử lần lượt các tên thay thế."""
    for t in ten:
        k = L.chuan_hoa(t)
        if k in df.columns:
            return df[k]
    if bat_buoc:
        raise ThieuCot(f"Không tìm thấy cột {ten[0]!r} (đã thử: {', '.join(ten)})")
    return pd.Series([None] * len(df), index=df.index, dtype=object)


def _s(series: pd.Series) -> pd.Series:
    return series.map(chuoi)


def _n(series: pd.Series) -> pd.Series:
    return series.map(so).astype(float)


def _r(series: pd.Series) -> pd.Series:
    return series.map(lam_tron).astype(float)


# ----------------------------------------------------------------------------
# Kết quả & cảnh báo
# ----------------------------------------------------------------------------
@dataclass
class CanhBao:
    muc: str          # nhóm cảnh báo
    noi_dung: str
    so_dong: int = 0
    hau_qua: str = ""
    chi_tiet: pd.DataFrame | None = None


class NhatKy(list):
    """Danh sách cảnh báo + bảng kiểm tra (vì sao dòng bị loại) của từng nguồn."""
    def __init__(self):
        super().__init__()
        self.chan_doan: list[dict] = []


def _thang(ng: pd.Series) -> str:
    dem = ng.dropna().map(lambda d: f"{d.month:02d}/{d.year}").value_counts().sort_index()
    return ", ".join(f"{k} ({v:,} dòng)" for k, v in dem.items()) or "không đọc được ngày"


def _chan_doan(cb, nguon: str, ly_do: list, ng: pd.Series | None = None,
               tu: date | None = None, den: date | None = None, ten_cot: str = "ngày"):
    """``ly_do``: [(nhãn, mask_bị_loại), ...] theo thứ tự ưu tiên; dòng không bị loại = Được tính."""
    n = len(ly_do[0][1]) if ly_do else 0
    con = pd.Series(True, index=ly_do[0][1].index) if ly_do else None
    dong = []
    for nhan, m in ly_do:
        k = int((m & con).sum())
        if k:
            dong.append((nhan, k))
        con &= ~m
    tinh_duoc = int(con.sum()) if con is not None else 0
    ds = getattr(cb, "chan_doan", None)
    if ds is not None:
        ds.append({"Nguồn": nguon, "Kết quả": "Đọc được", "Số dòng": n})
        ds.append({"Nguồn": nguon, "Kết quả": "✔ Được tính thuế", "Số dòng": tinh_duoc})
        for nhan, k in dong:
            ds.append({"Nguồn": nguon, "Kết quả": "✘ " + nhan, "Số dòng": k})
        if ng is not None:
            ds.append({"Nguồn": nguon, "Kết quả": f"Các tháng (theo {ten_cot}): " + _thang(ng), "Số dòng": None})
    if ng is not None and n and tu is not None:
        trong = ng.map(lambda d: d is not None and tu <= d <= den)
        if not trong.any():
            cb.append(CanhBao("KỲ", f"{nguon}: không có dòng nào trong kỳ {tu:%d/%m/%Y}–{den:%d/%m/%Y}",
                              n, f"Theo {ten_cot}, dữ liệu thuộc: {_thang(ng)} → kiểm tra lại tháng/năm đang chọn"))


@dataclass
class KetQua:
    tu_ngay: date
    den_ngay: date
    can_tru: pd.DataFrame
    xuat_km: pd.DataFrame
    xuat_vc: pd.DataFrame
    to_cp: pd.DataFrame
    kho_chi_tiet: pd.DataFrame
    kho: pd.DataFrame
    tong_hop_dt: pd.DataFrame
    tat_ca: pd.DataFrame
    theo_kh: pd.DataFrame
    canh_bao: list[CanhBao] = field(default_factory=list)
    nguon: pd.DataFrame | None = None
    chan_doan: pd.DataFrame | None = None

    def tong(self, df: pd.DataFrame) -> dict:
        return {c: float(df[c].sum()) if len(df) else 0.0 for c in COT_TIEN}


# ----------------------------------------------------------------------------
# Hồ sơ khách hàng
# ----------------------------------------------------------------------------
def lap_ho_so_kh(hskh: pd.DataFrame | None, hskh_import: pd.DataFrame | None,
                 cfg: Config, canh_bao: list[CanhBao]) -> pd.DataFrame:
    """Bảng KH: index = Mã KH; cột ten_kh, ten_dv, mst, cccd, ploai.

    Ưu tiên HSKH đã chuẩn hóa (có chỉnh tay); HSKH xuất từ CRM dùng để bổ sung
    KH mới / ô còn trống.
    """
    phan = []
    if hskh is not None and len(hskh):
        d = pd.DataFrame({
            "ma": _s(_cot(hskh, "Số")),
            "ten_kh": _s(_cot(hskh, "Tên Khách Hàng")),
            "ten_dv": _s(_cot(hskh, "Tên Đơn Vị")),
            "mst": _s(_cot(hskh, "MST")),
            "cccd": _s(_cot(hskh, "CCCD")),
            "ploai": _s(_cot(hskh, "PLoai (theo VAT/TNCN)", "Phân loại chốt (gõ tay ưu tiên)",
                             "Phân Loại Hiền", bat_buoc=False)),
        })
        sua = 0
        for c in ("ten_kh", "ten_dv", "mst", "cccd"):
            moi = d[c].map(bo_lap_doi)
            sua += int((moi != d[c]).sum())
            d[c] = moi
        if sua:
            canh_bao.append(CanhBao("HSKH", "Giá trị bị lặp đôi trong HSKH (vd 'Võ CaVõ Ca') đã được tự sửa",
                                    sua, "Đã tự sửa – nên sửa lại trong file HSKH gốc"))
        phan.append(d)
    if hskh_import is not None and len(hskh_import):
        vat = _s(_cot(hskh_import, "Thu Thuế VAT", bat_buoc=False))
        pit = _s(_cot(hskh_import, "Thuế TNCN", bat_buoc=False))
        plh = _s(_cot(hskh_import, "Phân Loại Hiền", bat_buoc=False))
        chiu = vat.str.startswith("1") | pit.str.startswith("1")
        hkd = cfg.loai_kh_chiu_thue[0]
        ploai = [
            (p if any(p.startswith(x) for x in cfg.loai_kh_chiu_thue) else hkd) if c else cfg.loai_kh_khong_chiu
            for c, p in zip(chiu, plh)
        ]
        d = pd.DataFrame({
            "ma": _s(_cot(hskh_import, "Số")),
            "ten_kh": _s(_cot(hskh_import, "Tên người đại diện")),
            "ten_dv": _s(_cot(hskh_import, "Tên Đơn Vị")),
            "mst": _s(_cot(hskh_import, "Mã Số Thuế GTGT", bat_buoc=False)),
            "cccd": _s(_cot(hskh_import, "Số CCCD/Hộ Chiếu", bat_buoc=False)),
            "ploai": ploai,
        })
        phan.append(d)
    if not phan:
        return pd.DataFrame(columns=["ten_kh", "ten_dv", "mst", "cccd", "ploai"])
    kq = phan[0]
    kq = kq[kq["ma"] != ""].drop_duplicates("ma").set_index("ma")
    for d in phan[1:]:
        d = d[d["ma"] != ""].drop_duplicates("ma").set_index("ma")
        kq = kq.reindex(kq.index.union(d.index, sort=False))
        d = d.reindex(kq.index)
        for c in kq.columns:
            trong = kq[c].isna() | (kq[c] == "")
            kq.loc[trong, c] = d.loc[trong, c]
    return kq.fillna("")


def la_hkd(ploai: pd.Series, cfg: Config) -> pd.Series:
    return ploai.fillna("").map(lambda p: any(p.startswith(x) for x in cfg.loai_kh_chiu_thue))


def _tra_kh(ma: pd.Series, kh: pd.DataFrame) -> pd.DataFrame:
    return kh.reindex(ma.values).reset_index(drop=True).fillna("").set_index(ma.index)


# ----------------------------------------------------------------------------
# Phân loại giao dịch
# ----------------------------------------------------------------------------
def phan_loai(tk_no: str, tk_co: str, cfg: Config) -> str:
    for no, co, pl in cfg.bang_phan_loai:
        if str(no) == tk_no and (str(co) == "*" or str(co) == tk_co):
            return pl
    return ""


def _tinh_ke_khai(chi: pd.Series, cfg: Config) -> tuple[pd.Series, pd.Series, pd.Series]:
    kk = _r(chi / (1 - cfg.ty_le_tong))
    return kk, _r(kk * cfg.ty_le_gtgt), _r(kk * cfg.ty_le_tncn)


def ngay_ky(df: pd.DataFrame, cfg: Config, loai: str) -> tuple[pd.Series, str]:
    """Ngày dùng để xác định tháng/năm (ngày chứng từ / ngày nhận hàng) và tên cột đã dùng."""
    ds = (cfg.cot_ngay or {}).get(loai) or []
    for ten in ds:
        k = L.chuan_hoa(ten)
        if k in df.columns:
            v = df[k].map(ngay)
            if v.notna().any():
                return v, ten
    if ds:
        return pd.Series([None] * len(df), index=df.index, dtype=object), ds[0]
    return pd.Series([None] * len(df), index=df.index, dtype=object), "?"


def _trong_ky(ng: pd.Series, tu: date, den: date) -> pd.Series:
    return ng.map(lambda d: d is not None and tu <= d <= den)


def _khung(n_index) -> pd.DataFrame:
    return pd.DataFrame(index=n_index, columns=COT_KQ)


# ----------------------------------------------------------------------------
# Bảng thuế suất
# ----------------------------------------------------------------------------
def lap_bang_thue(thue: pd.DataFrame | None) -> dict[str, float]:
    """Mshh -> thuế suất (0.08). Nhận cả dạng 8 hoặc 0.08; file có 2 bảng song song
    (A:C và E:G) thì bảng bên phải được ưu tiên như công thức gốc."""
    bang: dict[str, float] = {}
    if thue is None or not len(thue):
        return bang
    cap = []
    for c in thue.columns:
        if c.split("__")[0] == "mshh":
            hau_to = c[len("mshh"):]
            for tc in ("thuế", "thue", "thuế suất", "thue suat"):
                if tc + hau_to in thue.columns:
                    cap.append((c, tc + hau_to))
                    break
    for cm, ct in cap:   # cặp sau ghi đè cặp trước
        for m, t in zip(thue[cm], thue[ct]):
            m = chuoi(m)
            if not m or chuoi(t) == "":
                continue
            v = so(t)
            bang[m] = v / 100 if v > 1 else v
    return bang


# ----------------------------------------------------------------------------
# 1) Sổ quỹ: DATA CẤN TRỪ (Tổ DT) & DATA_CHI (Tổ CP)
# ----------------------------------------------------------------------------
def xu_ly_so_quy(df: pd.DataFrame | None, kh: pd.DataFrame, cfg: Config, che_do: str,
                 tu: date, den: date, canh_bao: list[CanhBao]) -> pd.DataFrame:
    if df is None or not len(df):
        return pd.DataFrame(columns=COT_KQ)
    phieu = _s(_cot(df, "Phiếu chi hạch toán"))
    ma = _s(_cot(df, "Mã số"))
    tk_no = _s(_cot(df, "mstk"))
    tien = _n(_cot(df, "Số tiền"))
    ng = _cot(df, "Ngày ghi sổ").map(ngay)
    tk_co, ghi_chu = [], []
    for p in phieu:
        hit = next((v for k, v in cfg.tien_to_phieu.items() if p.upper().startswith(k.upper())), None)
        tk_co.append(str(hit[0]) if hit else "")
        ghi_chu.append(hit[1] if hit else "")
    tk_co = pd.Series(tk_co, index=df.index)
    ghi_chu = pd.Series(ghi_chu, index=df.index)
    pl = pd.Series([phan_loai(a, b, cfg) for a, b in zip(tk_no, tk_co)], index=df.index)
    info = _tra_kh(ma, kh)
    hkd = la_hkd(info["ploai"], cfg)

    loc = (pl != "") & hkd
    khong_hskh = info["ploai"] == ""
    ly_do = []
    if che_do == L.CAN_TRU:
        loc &= (ghi_chu == "Cấn trừ") & (tk_co == "1311")
        ly_do.append(("Không phải phiếu cấn trừ (BTH…)", ~((ghi_chu == "Cấn trừ") & (tk_co == "1311"))))
    ly_do += [("TK Nợ/TK Có không thuộc bảng phân loại", pl == ""),
              ("Mã KH không có trong HSKH", khong_hskh),
              ("KH không phải hộ kinh doanh", ~hkd)]
    if che_do == L.CAN_TRU:
        loc_ngay = cfg.loc_ngay_can_tru
        nguoi = cfg.nguoi_phu_trach_dt
    else:
        loc_ngay = cfg.loc_ngay_to_cp
        nguoi = cfg.nguoi_phu_trach_cp
    nk, tnk = ngay_ky(df, cfg, che_do)
    if loc_ngay:
        loc &= _trong_ky(nk, tu, den)
        ly_do.append((f"Ngoài kỳ (theo {tnk})", ~_trong_ky(nk, tu, den)))
    _chan_doan(canh_bao, L.TEN_LOAI[che_do], ly_do, nk if loc_ngay else None, tu, den, tnk)

    co_phieu = ghi_chu != ""
    _bao_kh_thieu(ma, info, df, L.TEN_LOAI[che_do], canh_bao, (pl != "") & (ma != tk_no), cot_ten=("Tên KH", "Người thanh toán"))
    cap = [(str(x), str(y)) for x, y in cfg.cap_tk_khong_tinh]
    bo_qua = pd.Series([any(a == x and (y == "*" or b == y) for x, y in cap)
                        for a, b in zip(tk_no, tk_co)], index=df.index, dtype=bool)
    _bao_tk_moi(tk_no, tk_co, pl, hkd & co_phieu & ~bo_qua, L.TEN_LOAI[che_do], canh_bao)

    chi = tien
    kk, gtgt, tncn = _tinh_ke_khai(chi, cfg)
    out = pd.DataFrame({
        "Người phụ trách": nguoi,
        "Mã số KH": ma,
        "Họ và tên KH": info["ten_kh"],
        "Tên Quầy Thuốc/Nhà Thuốc/Tổ chức": info["ten_dv"],
        "Mã số thuế (nếu có)": info["mst"],
        "Số CMND": info["cccd"],
        "Số tiền sổ sách": tien,
        "Số tiền VAT": 0.0,
        "Số tiền chi (giá thị trường đã có VAT)": chi,
        "Số tiền Kê khai": kk,
        "Tiền thuế GTGT": gtgt,
        "Tiền thuế TNCN": tncn,
        "Đơn vị": _s(_cot(df, "msdv", bat_buoc=False)),
        "Số chứng từ": phieu,
        "Ngày ghi sổ": ng,
        "TK Nợ": tk_no,
        "TK Có": tk_co,
        "Phân Loại": pl,
        "Mã CTKM": _s(_cot(df, "Mã CTKM", bat_buoc=False)),
        "Tên CTKM": _s(_cot(df, "Tên CTKM", bat_buoc=False)),
    }, index=df.index)
    return out[loc].reset_index(drop=True)


# ----------------------------------------------------------------------------
# 2) DATA XUẤT KM
# ----------------------------------------------------------------------------
def _tra_so_cai(so_hd: pd.Series, so_cai: pd.DataFrame | None) -> pd.Series:
    """VLOOKUP(Sohd, SỔ CÁI, Tài Khoản): tự tìm cột chứa số hóa đơn."""
    if so_cai is None or not len(so_cai):
        return pd.Series([""] * len(so_hd), index=so_hd.index)
    tk = _s(_cot(so_cai, "Tài Khoản"))
    can_tim = set(so_hd[so_hd != ""])
    tot_nhat, diem = None, 0
    for c in so_cai.columns:
        if c.startswith("__"):
            continue
        v = _s(so_cai[c])
        d = int(v.isin(can_tim).sum())
        if d > diem:
            tot_nhat, diem = v, d
    if tot_nhat is None:
        return pd.Series([""] * len(so_hd), index=so_hd.index)
    bang = pd.Series(tk.values, index=tot_nhat.values)
    bang = bang[~bang.index.duplicated(keep="first")]
    return so_hd.map(bang).fillna("")


def _thue_suat(mshh: pd.Series, bang_thue: dict, du_phong: pd.Series, ten: str,
               canh_bao: list[CanhBao], can_dung: pd.Series) -> pd.Series:
    """Tra thuế suất theo Mshh; thiếu thì dùng thuế suất trên chứng từ (``du_phong``).
    Chỉ cảnh báo những dòng thực sự cần thuế suất (``can_dung``)."""
    ts = mshh.map(bang_thue)
    thieu = ts.isna() & (mshh != "")
    if thieu.any():
        dp = du_phong.map(lambda v: so(v) / 100 if so(v) > 1 else so(v))
        ts = ts.where(~thieu, dp)
        thieu &= can_dung
    if thieu.any():
        ds = sorted(set(mshh[thieu]))
        canh_bao.append(CanhBao(
            "THUẾ SUẤT", f"{ten}: mã hàng chưa có trong bảng THUE: {', '.join(ds[:30])}"
            + (" ..." if len(ds) > 30 else ""), int(thieu.sum()),
            "Đã dùng thuế suất trên chứng từ – nên bổ sung vào bảng THUE"))
    return ts.fillna(0.0).astype(float)


def xu_ly_xuat_km(df: pd.DataFrame | None, so_cai: pd.DataFrame | None, bang_thue: dict,
                  kh: pd.DataFrame, cfg: Config, tu: date, den: date,
                  canh_bao: list[CanhBao]) -> pd.DataFrame:
    if df is None or not len(df):
        return pd.DataFrame(columns=COT_KQ)
    ma = _s(_cot(df, "Mskh Dso"))
    so_hd = _s(_cot(df, "Sohd"))
    tk_no = _tra_so_cai(so_hd, so_cai)
    tk_co = tk_no.map(lambda t: str(cfg.tk_no_sang_tk_co.get(t, "")))
    pl = pd.Series([phan_loai(a, b, cfg) for a, b in zip(tk_no, tk_co)], index=df.index)
    mshh = _s(_cot(df, "Mshh"))
    cong_vat = pl.isin(cfg.phan_loai_cong_vat)
    ts = _thue_suat(mshh, bang_thue, _cot(df, "Thuesuat_daura", "Thue", bat_buoc=False), "Xuất KM",
                    canh_bao, cong_vat)
    goc = _n(_cot(df, "TTT Thue", "Ttien Co Vat"))
    so_sach = _n(_cot(df, "Ttien Co Vat"))
    vat = _r(goc * ts).where(cong_vat, 0.0)
    chi = _r(goc + vat)
    kk, gtgt, tncn = _tinh_ke_khai(chi, cfg)
    ng = _cot(df, "NgayGS").map(ngay)
    info = _tra_kh(ma, kh)
    hkd = la_hkd(info["ploai"], cfg)

    _bao_kh_thieu(ma, info, df, "Xuất KM", canh_bao, tk_no != "", cot_ten=("Tenkh", "Ten Nguoi Mua"))
    khong_sc = (tk_no == "") & hkd
    if khong_sc.any():
        canh_bao.append(CanhBao(
            "SỔ CÁI", "Xuất KM: số hóa đơn không tìm thấy trong SỔ CÁI (không có TK Nợ)",
            int(khong_sc.sum()), "Bị loại khỏi XUẤT KM – kiểm tra/bổ sung file Sổ cái",
            pd.DataFrame({"Số HĐ": so_hd[khong_sc], "MSKH": ma[khong_sc], "Mshh": mshh[khong_sc]})))
    _bao_tk_moi(tk_no, tk_co, pl, hkd, "Xuất KM", canh_bao, tk_loc=cfg.tk_no_loc_xuat)

    def ghep(a, b):
        return _s(_cot(df, a, bat_buoc=False)) + _s(_cot(df, b, bat_buoc=False))

    out = pd.DataFrame({
        "Người phụ trách": cfg.nguoi_phu_trach_dt,
        "Mã số KH": ma,
        "Họ và tên KH": info["ten_kh"],
        "Tên Quầy Thuốc/Nhà Thuốc/Tổ chức": info["ten_dv"],
        "Mã số thuế (nếu có)": info["mst"],
        "Số CMND": info["cccd"],
        "Số tiền sổ sách": so_sach,
        "Số tiền VAT": vat,
        "Số tiền chi (giá thị trường đã có VAT)": chi,
        "Số tiền Kê khai": kk,
        "Tiền thuế GTGT": gtgt,
        "Tiền thuế TNCN": tncn,
        "Đơn vị": _s(_cot(df, "Madv", bat_buoc=False)),
        "Số chứng từ": so_hd,
        "Ngày ghi sổ": ng,
        "TK Nợ": tk_no,
        "TK Có": tk_co,
        "Phân Loại": pl,
        "Mã CTKM": ghep("Maso Ctkm", "Promotion No_"),
        "Tên CTKM": ghep("Ten Ctkm", "Ten Hm"),
    }, index=df.index)
    nk, tnk = ngay_ky(df, cfg, L.XUAT_KM)
    loc = tk_no.isin([str(x) for x in cfg.tk_no_loc_xuat]) & hkd & _trong_ky(nk, tu, den)
    _chan_doan(canh_bao, "Xuất KM", [
        ("Số HĐ không có trong Sổ cái (không tra được TK Nợ)", tk_no == ""),
        ("TK Nợ không thuộc danh sách tính thuế", ~tk_no.isin([str(x) for x in cfg.tk_no_loc_xuat])),
        ("Mã KH không có trong HSKH", info["ploai"] == ""),
        ("KH không phải hộ kinh doanh", ~hkd),
        (f"Ngoài kỳ (theo {tnk})", ~_trong_ky(nk, tu, den)),
    ], nk, tu, den, tnk)
    return out[loc].reset_index(drop=True)


# ----------------------------------------------------------------------------
# 3) VOUCHER
# ----------------------------------------------------------------------------
def xu_ly_voucher(df: pd.DataFrame | None, bang_thue: dict, kh: pd.DataFrame, cfg: Config,
                  tu: date, den: date, canh_bao: list[CanhBao]) -> pd.DataFrame:
    if df is None or not len(df):
        return pd.DataFrame(columns=COT_KQ)
    ma = _s(_cot(df, "Mã số"))
    tk_no = _s(_cot(df, "mstk", bat_buoc=False))
    for du_phong in ("tk Nợ", "mstk_goc"):          # mstk trống thì lấy TK ở cột khác
        tk_no = tk_no.where(tk_no != "", _s(_cot(df, du_phong, bat_buoc=False)))
    tk_co = tk_no.map(lambda t: str(cfg.tk_no_sang_tk_co.get(t, "")))
    pl = pd.Series([phan_loai(a, b, cfg) for a, b in zip(tk_no, tk_co)], index=df.index)
    mshh = _s(_cot(df, "Mshh", "MSHH voucher", bat_buoc=False))
    cong_vat = pl.isin(cfg.phan_loai_cong_vat)
    ts = _thue_suat(mshh, bang_thue, pd.Series([0] * len(df), index=df.index), "Voucher", canh_bao, cong_vat)
    tien = _n(_cot(df, "Số tiền"))
    vat = _r(tien * ts).where(cong_vat, 0.0)
    chi = _r(tien + vat)
    kk, gtgt, tncn = _tinh_ke_khai(chi, cfg)
    ng = _cot(df, "Ngày ghi sổ").map(ngay)
    info = _tra_kh(ma, kh)
    hkd = la_hkd(info["ploai"], cfg)

    _bao_kh_thieu(ma, info, df, "Voucher", canh_bao, pl != "", cot_ten=("Tên KH", "Người thanh toán"))
    _bao_tk_moi(tk_no, tk_co, pl, hkd, "Voucher", canh_bao, tk_loc=cfg.tk_no_loc_xuat)

    out = pd.DataFrame({
        "Người phụ trách": cfg.nguoi_phu_trach_dt,
        "Mã số KH": ma,
        "Họ và tên KH": info["ten_kh"],
        "Tên Quầy Thuốc/Nhà Thuốc/Tổ chức": info["ten_dv"],
        "Mã số thuế (nếu có)": info["mst"],
        "Số CMND": info["cccd"],
        "Số tiền sổ sách": tien,
        "Số tiền VAT": vat,
        "Số tiền chi (giá thị trường đã có VAT)": chi,
        "Số tiền Kê khai": kk,
        "Tiền thuế GTGT": gtgt,
        "Tiền thuế TNCN": tncn,
        "Đơn vị": _s(_cot(df, "MSDV", bat_buoc=False)),
        "Số chứng từ": _s(_cot(df, "Phiếu xuất")),
        "Ngày ghi sổ": ng,
        "TK Nợ": tk_no,
        "TK Có": tk_co,
        "Phân Loại": pl,
        "Mã CTKM": _s(_cot(df, "Mã CTKM", bat_buoc=False)),
        "Tên CTKM": _s(_cot(df, "Tên CTKM", bat_buoc=False)),
    }, index=df.index)
    nk, tnk = ngay_ky(df, cfg, L.VOUCHER)
    loc = tk_no.isin([str(x) for x in cfg.tk_no_loc_xuat]) & hkd & _trong_ky(nk, tu, den)
    _chan_doan(canh_bao, "Voucher", [
        ("TK Nợ (mstk) không thuộc danh sách tính thuế", ~tk_no.isin([str(x) for x in cfg.tk_no_loc_xuat])),
        ("Mã KH không có trong HSKH", info["ploai"] == ""),
        ("KH không phải hộ kinh doanh", ~hkd),
        (f"Ngoài kỳ (theo {tnk})", ~_trong_ky(nk, tu, den)),
    ], nk, tu, den, tnk)
    return out[loc].reset_index(drop=True)


# ----------------------------------------------------------------------------
# 4) Dữ liệu kho (ThueHkdKho01 - phần Cường)
# ----------------------------------------------------------------------------
def xu_ly_kho(df: pd.DataFrame | None, kh: pd.DataFrame, cfg: Config, tu: date, den: date,
              canh_bao: list[CanhBao]) -> tuple[pd.DataFrame, pd.DataFrame]:
    if df is None or not len(df):
        return pd.DataFrame(columns=COT_KHO), pd.DataFrame(columns=COT_ALL)
    ng_giao = _cot(df, "Ngày giao hàng").map(ngay)
    nk, tnk = ngay_ky(df, cfg, L.KHO)
    ngoai_ky = ~_trong_ky(nk, tu, den)
    ma_kho_all = _s(_cot(df, "Mã kho", bat_buoc=False))
    ds_kho = [x.strip() for x in str(cfg.ma_kho or "").split(",") if x.strip()]
    khac_kho = ~ma_kho_all.isin(ds_kho) if ds_kho else pd.Series(False, index=df.index)
    if (khac_kho & ~ngoai_ky).any():
        cac_kho = ma_kho_all[khac_kho & ~ngoai_ky].value_counts()
        canh_bao.append(CanhBao("KHO", "Dữ liệu kho: có dòng thuộc kho khác " + ", ".join(ds_kho) + ": "
                                + ", ".join(f"{k or '(trống)'} ({v:,} dòng)" for k, v in cac_kho.items()),
                                int(cac_kho.sum()), "Bị loại – nếu cần tính, thêm mã kho vào cấu hình (ma_kho)"))
    toan_bo_ng = nk
    loc = ~ngoai_ky & ~khac_kho
    df_goc = df
    df = df[loc]
    ng_giao = ng_giao[loc]
    ma = _s(_cot(df, "Mã khách hàng"))
    info = _tra_kh(ma, kh)
    chiu = info["ploai"].where(info["ploai"] != "", "NA")
    pl = la_hkd(info["ploai"], cfg).astype(int)
    loai = _cot(df, "Loại").map(lambda v: chuoi(so(v)) if chuoi(v) else "")
    pl_km = loai.map(lambda v: cfg.loai_kho.get(v, f"Loại {v}"))
    vat_pt = _n(_cot(df, "VAT %"))
    goc = _n(_cot(df, "Tổng tiền giá gốc"))
    khong_tinh = pl_km.isin(cfg.pl_km_khong_tinh)
    thue = _r(goc * vat_pt / 100).where(~khong_tinh, 0.0)
    tt = (goc + thue).where(~khong_tinh & (pl == 1), 0.0)
    chi = _r(tt / (1 - cfg.ty_le_tong))
    t1 = _r(chi * cfg.ty_le_gtgt)
    t5 = _r(chi * cfg.ty_le_tncn)
    ten_kh = _s(_cot(df, "Tên khách hàng", bat_buoc=False))
    ten_kh = ten_kh.where(ten_kh != "", info["ten_kh"])
    ten_dv = _s(_cot(df, "Tên đơn vị", bat_buoc=False))
    ten_dv = ten_dv.where(ten_dv != "", info["ten_dv"])
    ma_hh = _s(_cot(df, "Mã hàng hóa", bat_buoc=False))
    ten_hh = _s(_cot(df, "Tên hàng hóa", bat_buoc=False))
    so_hd = _s(_cot(df, "Số hóa đơn", bat_buoc=False))
    ma_cn = _s(_cot(df, "Mã chi nhánh", bat_buoc=False))

    _bao_kh_thieu(ma, info, df, "Dữ liệu kho", canh_bao, None, cot_ten=("Tên khách hàng",))

    ct = pd.DataFrame({
        "Mã CN": ma_cn, "Mã KH": ma, "Tên KH": ten_kh, "Tên đơn vị": ten_dv,
        "Số HĐ": so_hd, "Ngày HĐ": _cot(df, "Ngày hóa đơn", bat_buoc=False).map(ngay),
        "Ngày giao": ng_giao, "Loại": loai, "Mã HH": ma_hh, "Tên HH": ten_hh,
        "PL KM": pl_km, "Chịu thuế": chiu, "PL": pl, "VAT %": vat_pt, "TT giá gốc": goc,
        "T.Tiền thuế": thue, "TT tính thuế": tt, "Số tiền chi": chi, "Thuế 1%": t1,
        "Thuế 0.5%": t5, "Tổng thuế": t1 + t5,
    }).reset_index(drop=True)

    chon = (ct["PL"] == 1) & ~ct["PL KM"].isin(cfg.pl_km_khong_tinh)
    idx = df.index
    def _m(v):
        s = pd.Series(False, index=df_goc.index)
        s.loc[idx] = v
        return s
    _chan_doan(canh_bao, "Dữ liệu kho (Cường)", [
        (f"Ngoài kỳ (theo {tnk})", ngoai_ky),
        ("Khác mã kho đang chọn", khac_kho),
        ("Xuất Bán (không tính thuế)", _m(ct["PL KM"].isin(cfg.pl_km_khong_tinh).values)),
        ("Mã KH không có trong HSKH", _m((info["ploai"] == "").values)),
        ("KH không phải hộ kinh doanh", _m((ct["PL"] != 1).values)),
    ], toan_bo_ng, tu, den, tnk)
    c = ct[chon]
    inf = info.reset_index(drop=True)[chon.values]
    tong = pd.DataFrame({
        "Người phụ trách": cfg.nguoi_phu_trach_kho,
        "Mã số KH": c["Mã KH"],
        "Họ và tên KH": c["Tên KH"],
        "Tên Quầy Thuốc/Nhà Thuốc/Tổ chức": c["Tên đơn vị"],
        "Mã số thuế (nếu có)": inf["mst"].values,
        "Số CMND": inf["cccd"].values,
        "Số tiền sổ sách": c["TT giá gốc"],
        "Số tiền VAT": c["T.Tiền thuế"],
        "Số tiền chi (giá thị trường đã có VAT)": c["TT tính thuế"],
        "Số tiền Kê khai": c["Số tiền chi"],
        "Tiền thuế GTGT": c["Thuế 1%"],
        "Tiền thuế TNCN": c["Thuế 0.5%"],
        "Đơn vị": c["Mã CN"],
        "Số chứng từ": c["Số HĐ"],
        "Ngày ghi sổ": c["Ngày giao"],
        "TK Nợ": "", "TK Có": "",
        "Phân Loại": c["PL KM"],
        "Mã CTKM": "", "Tên CTKM": "",
        "Nguồn": cfg.nguon_kho,
        "Ghi chú (Mã HH - Tên HH)": c["Mã HH"] + " - " + c["Tên HH"],
    }).reset_index(drop=True)
    return ct, tong


# ----------------------------------------------------------------------------
# Cảnh báo
# ----------------------------------------------------------------------------
def _bao_kh_thieu(ma, info, df, ten_nguon, canh_bao, lien_quan=None, cot_ten=()):
    """MSKH không có trong HSKH (chỉ xét các dòng ``lien_quan`` – có ảnh hưởng tới kết quả)."""
    thieu = (info["ploai"] == "") & (info["ten_kh"] == "") & (ma != "")
    if lien_quan is not None:
        thieu &= lien_quan
    if not thieu.any():
        return
    ten = _cot(df, *cot_ten, bat_buoc=False) if cot_ten else pd.Series([""] * len(df), index=df.index)
    ct = (pd.DataFrame({"MSKH": ma[thieu], "Tên KH (trong nguồn)": _s(ten[thieu])})
          .groupby("MSKH", sort=False).agg(**{"Tên KH (trong nguồn)": ("Tên KH (trong nguồn)", "first"),
                                              "Số dòng": ("Tên KH (trong nguồn)", "size")}).reset_index())
    canh_bao.append(CanhBao(
        "HSKH", f"{ten_nguon}: {len(ct)} MSKH không có trong HSKH", int(thieu.sum()),
        "Thiếu Phân loại, MST, CMND → bị loại khỏi bảng kê", ct))


def _bao_tk_moi(tk_no, tk_co, pl, hkd, ten_nguon, canh_bao, tk_loc=None):
    """HKD nhưng cặp TK chưa có trong bảng phân loại / điều kiện lọc."""
    m = hkd & (tk_no != "") & ((pl == "") | (~tk_no.isin([str(x) for x in tk_loc]) if tk_loc else False))
    if not m.any():
        return
    ct = (pd.DataFrame({"TK Nợ": tk_no[m], "TK Có": tk_co[m], "Phân loại": pl[m]})
          .value_counts().rename("Số dòng").reset_index())
    canh_bao.append(CanhBao(
        "TÀI KHOẢN", f"{ten_nguon}: dòng của HKD có TK chưa nằm trong quy tắc tính thuế", int(m.sum()),
        "Không tính thuế – nếu cần tính, thêm TK vào cấu hình (bảng phân loại / TK lọc)", ct))


# ----------------------------------------------------------------------------
# Tổng hợp
# ----------------------------------------------------------------------------
def theo_khach_hang(tat_ca: pd.DataFrame) -> pd.DataFrame:
    if not len(tat_ca):
        return pd.DataFrame(columns=["Mã số KH", "Họ và tên KH", "Tên Quầy Thuốc/Nhà Thuốc/Tổ chức",
                                     "Mã số thuế (nếu có)", "Số CMND", *COT_TIEN, "Số dòng", "Nguồn"])
    g = tat_ca.groupby("Mã số KH", sort=False)
    kq = g.agg(**{
        "Họ và tên KH": ("Họ và tên KH", "first"),
        "Tên Quầy Thuốc/Nhà Thuốc/Tổ chức": ("Tên Quầy Thuốc/Nhà Thuốc/Tổ chức", "first"),
        "Mã số thuế (nếu có)": ("Mã số thuế (nếu có)", "first"),
        "Số CMND": ("Số CMND", "first"),
        **{c: (c, "sum") for c in COT_TIEN},
        "Số dòng": ("Số tiền Kê khai", "size"),
        "Nguồn": ("Nguồn", lambda s: ", ".join(dict.fromkeys(s))),
    }).reset_index()
    return kq.sort_values("Số tiền Kê khai", ascending=False, kind="stable").reset_index(drop=True)


def ky_tu_du_lieu(du_lieu: dict[str, pd.DataFrame], cfg: Config | None = None) -> tuple[int, int] | None:
    """Tháng/năm xuất hiện nhiều nhất trong các nguồn phát sinh (theo ngày chứng từ / ngày nhận hàng)."""
    cfg = cfg or Config()
    # Ưu tiên nguồn phát sinh của tổ (Xuất KM, Voucher, sổ quỹ); dữ liệu kho thường kéo dài nhiều tháng.
    for nhom in ((L.XUAT_KM, L.VOUCHER, L.CAN_TRU, L.CHI), (L.KHO,)):
        dem: dict[tuple[int, int], int] = {}
        for loai in nhom:
            df = du_lieu.get(loai)
            if df is None or not len(df):
                continue
            nk, _ = ngay_ky(df, cfg, loai)
            for d in nk.dropna():
                dem[(d.month, d.year)] = dem.get((d.month, d.year), 0) + 1
        if dem:
            return max(dem, key=dem.get)
    return None


def tinh(du_lieu: dict[str, pd.DataFrame], tu: date, den: date, cfg: Config | None = None) -> KetQua:
    """Chạy toàn bộ quy trình. ``du_lieu`` = kết quả ``doc_file.gop_theo_loai``."""
    cfg = cfg or Config()
    cb = NhatKy()
    kh = lap_ho_so_kh(du_lieu.get(L.HSKH), du_lieu.get(L.HSKH_IMPORT), cfg, cb)
    if not len(kh):
        cb.append(CanhBao("HSKH", "Chưa có file hồ sơ khách hàng (HSKH)", 0,
                          "Không xác định được KH nào là HKD → bảng kê sẽ rỗng"))
    bang_thue = lap_bang_thue(du_lieu.get(L.THUE))
    if L.XUAT_KM in du_lieu and L.SO_CAI not in du_lieu:
        cb.append(CanhBao("SỔ CÁI", "Có dữ liệu Xuất KM nhưng chưa có file Sổ cái", 0,
                          "Không tra được TK Nợ → toàn bộ Xuất KM bị loại"))

    ct = xu_ly_so_quy(du_lieu.get(L.CAN_TRU), kh, cfg, L.CAN_TRU, tu, den, cb)
    km = xu_ly_xuat_km(du_lieu.get(L.XUAT_KM), du_lieu.get(L.SO_CAI), bang_thue, kh, cfg, tu, den, cb)
    vc = xu_ly_voucher(du_lieu.get(L.VOUCHER), bang_thue, kh, cfg, tu, den, cb)
    cp = xu_ly_so_quy(du_lieu.get(L.CHI), kh, cfg, L.CHI, tu, den, cb)
    kho_ct, kho = xu_ly_kho(du_lieu.get(L.KHO), kh, cfg, tu, den, cb)

    dt = pd.concat([x for x in (ct, km, vc) if len(x)] or [pd.DataFrame(columns=COT_KQ)], ignore_index=True)
    phan = []
    for d, nguon in ((dt, cfg.nguon_dt), (cp, cfg.nguon_cp)):
        if len(d):
            d = d.copy()
            d["Nguồn"] = nguon
            d["Ghi chú (Mã HH - Tên HH)"] = ""
            phan.append(d)
    if len(kho):
        phan.append(kho)
    tat_ca = pd.concat(phan, ignore_index=True) if phan else pd.DataFrame(columns=COT_ALL)
    tat_ca = tat_ca[COT_ALL]

    thieu_id = (tat_ca["Mã số thuế (nếu có)"] == "") & (tat_ca["Số CMND"] == "")
    if thieu_id.any():
        cb.append(CanhBao("HSKH", "Dòng trong bảng kê thiếu cả MST và CMND/CCCD", int(thieu_id.sum()),
                          "Cần bổ sung để kê khai",
                          tat_ca.loc[thieu_id, ["Mã số KH", "Họ và tên KH", "Nguồn"]].drop_duplicates()))

    for loai, ten in ((L.CAN_TRU, "Sổ quỹ - Cấn trừ (Tổ DT)"), (L.XUAT_KM, "Xuất KM"), (L.VOUCHER, "Voucher"),
                      (L.CHI, "Sổ quỹ - Chi (Tổ CP)"), (L.KHO, "Dữ liệu kho (Cường)"), (L.SO_CAI, "Sổ cái"),
                      (L.HSKH, "Hồ sơ KH"), (L.THUE, "Bảng thuế suất")):
        if loai not in du_lieu and not (loai == L.HSKH and L.HSKH_IMPORT in du_lieu):
            cb.chan_doan.append({"Nguồn": ten, "Kết quả": "⚠ Chưa có file", "Số dòng": 0})
    kq = KetQua(tu, den, ct, km, vc, cp, kho_ct, kho, dt, tat_ca, theo_khach_hang(tat_ca), list(cb))
    kq.chan_doan = pd.DataFrame(cb.chan_doan, columns=["Nguồn", "Kết quả", "Số dòng"])
    return kq
