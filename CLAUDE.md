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

## Việc đang dở (ưu tiên)
1. Người dùng chạy trên dữ liệu thật thấy **Voucher = 0** (Tổ DT chỉ còn 2.736 dòng = riêng Xuất KM) và **cấn trừ sai/thiếu**.
   → Chạy `python cli.py "<thư mục dữ liệu thật>"`, xem phần KIỂM TRA DỮ LIỆU / LƯU Ý, mở file voucher thật xem tên cột,
   định dạng ngày, mã KH; sửa cho đúng ở CẢ Python và web; thêm test.
2. Hỏi người dùng số phiếu cụ thể nếu vẫn lệch.
