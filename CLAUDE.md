# Bối cảnh dự án PIT-VAT (đọc trước khi làm việc)

Phần mềm tập hợp & tính thuế Hộ kinh doanh (GTGT 1% + TNCN 0,5%) cho CN Nha Trang (NTG), thay cho file Excel
`2026.10_NTG_Tập hợp Thuế Hộ Kinh Doanh GỘP 3 TỔ (Doanh Thu_CP_13887).xlsb`.
Người dùng: kế toán, giao tiếp bằng tiếng Việt. Dữ liệu thật nằm ở
`/Users/hien/SynologyDrive/KE TOAN TRUONG/ANNUAL DATA/03. KẾ TOÁN - CHỨNG TỪ/Năm 2026/2026_Thue HKD`
— KHÔNG commit dữ liệu thật (đã chặn *.xls*/*.csv trong .gitignore).

## Hai bản, CÙNG một logic — sửa phải sửa cả hai
* Python: `thue_hkd/` (doc_file.py đọc & nhận diện, tinh_toan.py tính, xuat_excel.py, config.py), `cli.py`, `app.py` (Streamlit).
* Trình duyệt: `web/thue_hkd.html` (1 file, SheetJS). Sau khi sửa chạy `python scripts/tao_index.py` để tạo `web/index.html`.
* Test: `python -m pytest -q`. Khi đổi logic, đối chiếu kết quả Python và web trên cùng bộ file.

## Quy tắc đã chốt với người dùng
* Kỳ (tháng/năm): Voucher, Sổ quỹ chi (Tổ CP), Sổ quỹ cấn trừ → **Ngày chứng từ**;
  Kho 01 (dữ liệu bán hàng theo kho, phần Cường 13887) → **ngày nhận hàng** (cột "Ngày giao hàng");
  Kho 03 (Xuất KM, phiếu DKM/NTG03) → **ngày nhận hàng** (cột "Ngay NH", thiếu thì NgayGS). Mọi nguồn đều lọc theo kỳ.
* Ngày phần lớn là TEXT; mặc định ngày trước tháng (dd/mm/yyyy). Ngày không đọc được phải báo, không được âm thầm loại.
* Không tính thuế: 6418/1121 (mời cơm sinh nhật KH – Tổ CP), 33882 (chi lại CKTT).
* Cấn trừ: phiếu BTH → TK Có 1311; tính khi TK có phân loại và KH là hộ kinh doanh (HSKH).
* Gross-up: Kê khai = ROUND(chi/(1−1,5%)); GTGT = ROUND(KK×1%); TNCN = ROUND(KK×0,5%) (ROUND kiểu Excel).
* Đã đối chiếu tháng 09/2026 với file Excel gốc: khớp 5.849/5.849 dòng
  (Tổ DT 2.806 dòng, KK 624.526.712; Cường 3.043 dòng, KK 578.243.533; tổng KK 1.202.770.245, GTGT 12.027.688, TNCN 6.013.263).

## Tình trạng (11/11/2026, chạy trên dữ liệu thật – 9 file trong thư mục 2026_Thue HKD)
* Tháng 09: khớp Excel gốc (Tổ DT 2.806 = Xuất KM 2.736 + Voucher 70; Cấn trừ 0; Cường 3.043). Tổ CP 264 dòng, KK 57.614.359
  (CKTT T09 chi qua ngân hàng PCN, TK 635/335).
* Tháng 10: Tổ DT 1.256 (Cấn trừ 403, Xuất KM 776, Voucher 77) KK 242.083.989; Tổ CP 0 (chỉ có 33882 và 6418/1121 → không tính);
  Cường 2.980 KK 607.996.893; tổng KK 850.080.882, GTGT 8.500.928, TNCN 4.250.451.
* Cấn trừ tháng 09 = 0 là đúng dữ liệu: chỉ 4 dòng BTH, đều là công ty (0680/00004, 0583/00143).
* Voucher TK 33412 = tặng nhân viên (mã KH là mã NV) → không tính, không báo thiếu HSKH. Dòng chữ ký cuối file không báo lỗi ngày.

* 11/11/2026 người dùng chốt: Tổ CP **TK Nợ 1312 / TK Có 1121** ("Chi tiền khuyến mãi TB" chi qua NH) → TÍNH THUẾ,
  phân loại "Chi TM TB KM" (đã thêm vào config.py, web, cau_hinh_mau.json). Thêm 434 dòng ở T01/T02/T06/T07; T09, T10 không đổi.

* File Excel đối chiếu của người dùng: `.../03. KẾ TOÁN - CHỨNG TỪ/KPI/Năm 2026/2026.10_KPI -  Báo cáo/2026.10_Tiêu chí KPI_Ngày 1 hàng tháng/`
  (`2026.10_NTG_Tập hợp thuế Hộ Kinh Doanh _Tổ Doanh Thu.xlsx` = Tổ DT T10; `..._GỘP 3 TỔ ....xlsb` sheet ThueHkdKho01 = Cường T10).
  Tháng 10 khớp TỪNG DÒNG: Tổ DT 1.256 (KK 242.083.989), Cường 2.981 (KK 606.557.807, thuế 9.098.469), Tổ CP 0.
* **HĐ kê bổ sung (sót kỳ trước)** – như ô "HĐ kê bổ sung" trong sheet ThueHkdKho01: số HĐ kho ngoài kỳ được tính vào kỳ này,
  ghi chú "Kê bổ sung (sót kỳ trước)". Config `hd_ke_bo_sung`, CLI `--ke-bo-sung "1C26TNG|00037622"`, ô nhập trên Streamlit & web.
  T10 cần kê bổ sung 1C26TNG|00037622 (BDC ngày 30/09, −1.439.086) vì bảng kê T09 đã nộp không có dòng này.
* Lưu ý: file 2026.09 (bảng kê T09 đã nộp, Sheet2) tính "Quà tặng KH (CP)" theo giá ĐÃ CÓ VAT (KK Tổ DT 651.199.445);
  từ file 2026.10 trở đi (và phần mềm) tính theo giá TRƯỚC VAT (624.526.712). Người dùng chưa xác nhận cách nào đúng cho T09.

## Việc đang dở
1. 11/11: ĐÃ chạy bản web (Node + SheetJS 0.18.5) trên 9 file thật: khớp Python T01/02/06/07/09/10. Lỗi cũ: file Bảng cấn trừ/kê chi/voucher
   ghi sai vùng (!ref "A1:AF9") → SheetJS chỉ đọc 5 dòng; đã thêm `suaVung()` tính lại vùng theo ô thực có. Cấu hình cũ lưu trong
   localStorage được bổ sung các dòng `bang_phan_loai` mới.
2. Còn 2 dòng HKD TK 1331/1121 (thuế GTGT đầu vào) ở Tổ CP – chưa tính, có thể hỏi người dùng nếu cần.
