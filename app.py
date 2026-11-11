"""Giao diện web (chạy trên máy): kéo-thả nhiều file -> tính -> tải file Excel kết quả.

Chạy:  streamlit run app.py   (hoặc nhấp đúp chay_phan_mem.bat trên Windows)
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from thue_hkd import Config, ghi_excel, ky_thang
from thue_hkd import doc_file as L
from thue_hkd.quy_trinh import bang_nguon, doc_nhieu_file
from thue_hkd.tinh_toan import COT_TIEN, tinh

st.set_page_config(page_title="Thuế Hộ Kinh Doanh", page_icon="🧾", layout="wide")
st.title("🧾 Tập hợp & tính thuế Hộ kinh doanh")
st.caption("Kéo-thả tất cả file xuất từ phần mềm (nhiều file cùng loại cũng được). "
           "Phần mềm tự nhận diện loại dữ liệu, gộp lại và tính GTGT 1% + TNCN 0,5%.")

# ---------------- Thanh bên: kỳ & cấu hình ----------------
with st.sidebar:
    st.header("Kỳ tính thuế")
    hom_nay = date.today()
    thang_truoc = (hom_nay.month - 2) % 12 + 1
    nam_mac_dinh = hom_nay.year if hom_nay.month > 1 else hom_nay.year - 1
    kieu = st.radio("Chọn kỳ", ["Theo tháng", "Khoảng ngày"], horizontal=True)
    if kieu == "Theo tháng":
        c1, c2 = st.columns(2)
        thang = c1.number_input("Tháng", 1, 12, thang_truoc)
        nam = c2.number_input("Năm", 2000, 2100, nam_mac_dinh)
        tu, den = ky_thang(int(thang), int(nam))
    else:
        tu = st.date_input("Từ ngày", ky_thang(thang_truoc, nam_mac_dinh)[0], format="DD/MM/YYYY")
        den = st.date_input("Đến ngày", ky_thang(thang_truoc, nam_mac_dinh)[1], format="DD/MM/YYYY")
    st.write(f"Kỳ: **{tu:%d/%m/%Y} → {den:%d/%m/%Y}**")
    hd_bs = st.text_area("HĐ kê bổ sung – sót kỳ trước (phần Cường, cách nhau dấu ;)", "",
                         help="Số hóa đơn thuộc kỳ trước nhưng chưa kê (vd 1C26TNG|00037622) → tính vào kỳ này")

    st.header("Cấu hình")
    f_cfg = st.file_uploader("File cấu hình .json (không bắt buộc)", type=["json"])
    cfg = Config.from_json(f_cfg.getvalue().decode("utf-8")) if f_cfg else Config()
    with st.expander("Tham số chính"):
        cfg.ma_kho = st.text_input("Mã kho (phần Cường) – nhiều kho cách nhau dấu phẩy", cfg.ma_kho)
        cfg.nguoi_phu_trach_dt = st.text_input("Người phụ trách Tổ DT", cfg.nguoi_phu_trach_dt)
        cfg.nguoi_phu_trach_cp = st.text_input("Người phụ trách Tổ CP", cfg.nguoi_phu_trach_cp)
        cfg.nguoi_phu_trach_kho = st.text_input("Người phụ trách kho", cfg.nguoi_phu_trach_kho)
        cfg.loc_ngay_can_tru = st.checkbox("Lọc Cấn trừ theo kỳ", cfg.loc_ngay_can_tru)
        cfg.loc_ngay_to_cp = st.checkbox("Lọc Tổ CP theo kỳ", cfg.loc_ngay_to_cp)
    if hd_bs.strip():
        cfg.hd_ke_bo_sung = hd_bs
    st.download_button("Tải cấu hình hiện tại (.json)", cfg.to_json().encode("utf-8"),
                       "cau_hinh.json", "application/json")

# ---------------- Tải file ----------------
files = st.file_uploader("Chọn / kéo-thả các file xuất (.xlsx, .xls, .xlsb, .csv)",
                         type=["xlsx", "xlsm", "xls", "xlsb", "csv", "ods"], accept_multiple_files=True)
if not files:
    st.info("Cần các loại file: Sổ quỹ (cấn trừ / chi), Xuất KM, Voucher, Sổ cái, Hồ sơ KH (HSKH), "
            "Bảng thuế suất, Dữ liệu bán hàng theo kho. Thiếu loại nào thì phần đó sẽ trống.")
    st.stop()


@st.cache_data(show_spinner="Đang đọc file ...")
def _doc(ds: tuple[tuple[str, bytes], ...]):
    bangs, loi = doc_nhieu_file([(ten, io.BytesIO(b)) for ten, b in ds])
    return bangs, loi


bangs, loi = _doc(tuple((f.name, f.getvalue()) for f in files))
for l in loi:
    st.warning(l)

st.subheader("1. Kiểm tra loại dữ liệu đã nhận diện")
st.caption("Nếu nhận sai, chọn lại ở cột “Loại dữ liệu”.")
ten_sang_ma = {v: k for k, v in L.TEN_LOAI.items()}
bng = bang_nguon(bangs)
bng.insert(0, "Mã", [b.mo_ta for b in bangs])
sua = st.data_editor(
    bng, hide_index=True, use_container_width=True, disabled=["Mã", "File", "Sheet", "Số dòng", "Dòng tiêu đề"],
    column_config={"Mã": None, "Loại dữ liệu": st.column_config.SelectboxColumn(options=list(L.TEN_LOAI.values()))},
)
for b, moi in zip(bangs, sua["Loại dữ liệu"]):
    b.loai = ten_sang_ma.get(moi, b.loai)

co = {b.loai for b in bangs}
thieu = [L.TEN_LOAI[k] for k in (L.HSKH, L.SO_CAI, L.THUE) if k not in co and not (k == L.HSKH and L.HSKH_IMPORT in co)]
if thieu:
    st.warning("Chưa có: " + ", ".join(thieu))

# ---------------- Tính ----------------
st.subheader("2. Kết quả")
if st.button("▶ Tính thuế", type="primary") or st.session_state.get("da_tinh"):
    st.session_state["da_tinh"] = True
    with st.spinner("Đang tính ..."):
        kq = tinh(L.gop_theo_loai([b for b in bangs if b.loai != L.KHONG_RO]), tu, den, cfg)
        kq.nguon = bang_nguon(bangs)
        data = ghi_excel(kq, cfg=cfg)

    tong = pd.DataFrame([{"Nguồn": n, "Số dòng": len(d), **kq.tong(d)} for n, d in (
        (cfg.nguon_dt, kq.tong_hop_dt), (cfg.nguon_cp, kq.to_cp), (cfg.nguon_kho, kq.kho),
        ("TỔNG CỘNG", kq.tat_ca))])
    c = st.columns(3)
    t = kq.tong(kq.tat_ca)
    c[0].metric("Số tiền kê khai", f"{t['Số tiền Kê khai']:,.0f}")
    c[1].metric("Thuế GTGT", f"{t['Tiền thuế GTGT']:,.0f}")
    c[2].metric("Thuế TNCN", f"{t['Tiền thuế TNCN']:,.0f}")
    st.dataframe(tong.style.format({k: "{:,.0f}" for k in COT_TIEN}), hide_index=True, use_container_width=True)

    st.download_button("⬇ Tải file Excel kết quả", data, f"KetQua_ThueHKD_{tu:%Y%m%d}_{den:%Y%m%d}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

    if kq.canh_bao:
        st.subheader(f"⚠ Lưu ý ({len(kq.canh_bao)})")
        for cb in kq.canh_bao:
            with st.expander(f"[{cb.muc}] {cb.noi_dung} — {cb.so_dong} dòng"):
                st.write(cb.hau_qua)
                if cb.chi_tiet is not None and len(cb.chi_tiet):
                    st.dataframe(cb.chi_tiet, hide_index=True, use_container_width=True)
    else:
        st.success("Không có bất thường.")

    tabs = st.tabs(["Tổng hợp", "Theo khách hàng", "Cấn trừ", "Xuất KM", "Xuất VC", "Tổ CP", "Kho (Cường)"])
    for tab, df in zip(tabs, (kq.tat_ca, kq.theo_kh, kq.can_tru, kq.xuat_km, kq.xuat_vc, kq.to_cp,
                              kq.kho_chi_tiet)):
        with tab:
            st.caption(f"{len(df):,} dòng")
            st.dataframe(df.head(5000), hide_index=True, use_container_width=True)
