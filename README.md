# Phần mềm Tập hợp & Tính thuế Hộ kinh doanh (GTGT 1% + TNCN 0,5%)

Thay thế file Excel *“Tập hợp Thuế Hộ Kinh Doanh”* (các sheet DATA CẤN TRỪ, DATA XUẤT KM,
VOUCHER, DATA_CHI, ThueHkdKho01 → TONGHOP / ALL). Bạn **không cần dán dữ liệu vào một file Excel** nữa:
xuất từ phần mềm ra bao nhiêu file cũng được, kéo-thả tất cả vào, phần mềm sẽ tự:

1. **Nhận diện** từng file là loại dữ liệu gì (dựa vào dòng tiêu đề, dòng tiêu đề không cần nằm ở dòng 1).
2. **Gộp** các file cùng loại (ví dụ Xuất KM chia 3 file theo tuần).
3. **Tính** theo đúng logic của file Excel cũ và **xuất ra 1 file Excel kết quả**.

Đã đối chiếu với file `2026.10_NTG_Tập hợp Thuế HKD … 13887.xlsb` (kỳ 09/2026): khớp **từng dòng, từng đồng**
với sheet ALL (5.849 dòng; Tổ DT 2.806 dòng, Cường 3.043 dòng; tổng kê khai 1.202.770.245, GTGT 12.027.688,
TNCN 6.013.263).

## Cách chạy

### Windows (dễ nhất)
1. Cài Python 3.10 trở lên từ <https://www.python.org/downloads/> (nhớ tick **Add python.exe to PATH**).
2. Tải thư mục này về, nhấp đúp **`chay_phan_mem.bat`**. Lần đầu sẽ mất vài phút để cài thư viện.
3. Trình duyệt tự mở trang phần mềm (chạy trên máy bạn, dữ liệu không gửi đi đâu).
4. Chọn tháng/năm → kéo-thả các file → kiểm tra bảng “loại dữ liệu” → bấm **Tính thuế** → **Tải file Excel kết quả**.

### Dòng lệnh (chạy cả thư mục)
```bash
pip install -r requirements.txt
```
```bash
python cli.py --thang 9 --nam 2026 "D:/ThueHKD/T9" -o "D:/ThueHKD/KetQua_T9.xlsx"
```

## Các loại file cần xuất

| Loại | Tương ứng sheet cũ | Cột nhận diện (bắt buộc có) |
|---|---|---|
| Sổ quỹ – Cấn trừ (Tổ DT) | DATA CẤN TRỪ | Phiếu chi hạch toán, mstk, Số tiền, Mã số |
| Sổ quỹ – Chi (Tổ CP) | DATA_CHI | như trên; tên file/sheet chứa “CHI”/“TO CP” hoặc đa số phiếu là PC… |
| Xuất khuyến mãi | DATA XUẤT KM | Mshh, Sohd, Mskh Dso, Ttien Co Vat, NgayGS |
| Voucher | VOUCHER | Phiếu xuất, Mshh, mstk, Số tiền, Mã số |
| Sổ cái | SỔ CÁI | Số Chứng Từ, Tài Khoản, TKĐƯ |
| Hồ sơ KH (chuẩn hóa) | HSKH | Số, Tên Khách Hàng, Tên Đơn Vị, MST, CCCD (+ PLoai (theo VAT/TNCN)) |
| Hồ sơ KH (xuất CRM) | HSKH IMPORT | Số, Tên người đại diện, Tên Đơn Vị, Mã Số Thuế GTGT |
| Bảng thuế suất | THUE | Mshh, Tenhh, Thuế (0.08 hoặc 8 đều được) |
| Dữ liệu kho (Cường) | Data / ThueHkdKho01 | Mã chi nhánh, Mã khách hàng, Ngày giao hàng, Loại, Tổng tiền giá gốc, VAT % |

* Định dạng: `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`. Một file nhiều sheet cũng được — mỗi sheet được nhận diện riêng
  (có thể đưa luôn file Excel cũ vào).
* Nếu nhận diện sai, sửa lại ở cột “Loại dữ liệu” trên giao diện.
* Có cả HSKH chuẩn hóa và HSKH xuất CRM: ưu tiên HSKH chuẩn hóa (giữ các chỉnh sửa tay), file CRM bổ sung KH mới / ô trống.

## Logic tính (giữ nguyên như file Excel)

**Chung:** `Số tiền kê khai = ROUND(Số tiền chi / (1 − 1,5%))`, `GTGT = ROUND(kê khai × 1%)`,
`TNCN = ROUND(kê khai × 0,5%)` (làm tròn kiểu Excel). Chỉ lấy KH có phân loại “7. HỘ KINH DOANH” / “2. HIỆU THUỐC…”.

* **Cấn trừ:** phiếu `BTH…` → TK Có 1311, “Cấn trừ”; phân loại theo bảng TK; Số tiền chi = Số tiền. Không lọc ngày.
* **Xuất KM:** TK Nợ tra từ Sổ cái theo số hóa đơn; TK Có theo bảng ánh xạ; lọc TK Nợ 1312/33871/6418/335 và Ngày GS
  trong kỳ. Quà TB CLB / Quà TB KM: `VAT = ROUND(TTT Thuế × thuế suất)`, `Số tiền chi = TTT Thuế + VAT`;
  quà tặng KH (CP): VAT = 0, Số tiền chi = TTT Thuế. Số tiền sổ sách = Ttien Co Vat.
* **Voucher:** TK Nợ = mstk; lọc như Xuất KM; Số tiền chi = Số tiền + VAT.
* **Tổ CP:** sổ quỹ chi, lấy dòng có phân loại và KH là HKD.
* **Kho (Cường):** lọc Ngày giao trong kỳ + mã kho; Loại 0/1/10 = Xuất Bán / Xuất KM / Xuất Tặng;
  `T.Tiền thuế = ROUND(TT giá gốc × VAT%)` (trừ Xuất Bán); chỉ đưa vào tổng hợp dòng PL = 1 và không phải Xuất Bán.

Bảng phân loại, ánh xạ TK, TK lọc, người phụ trách, mã kho… nằm trong **`cau_hinh_mau.json`**. Khi phát sinh TK mới
chỉ cần sửa file này (rồi tải lên ở thanh bên hoặc dùng `--cau-hinh`), không phải sửa 4 chỗ công thức như trước.

## File Excel kết quả

| Sheet | Nội dung |
|---|---|
| TỔNG HỢP | Tương đương sheet ALL (22 cột, có cột Nguồn), tổng theo từng tổ ở đầu, tổng theo bộ lọc |
| THEO KHÁCH HÀNG | Cộng dồn theo từng KH – tiện để kê khai |
| TONGHOP (Tổ DT) | Cấn trừ + Xuất KM + Voucher |
| CẤN TRỪ / XUẤT KM / XUẤT VC / TỔ CP | Chi tiết từng nguồn |
| KHO (Cường) | Tương đương ThueHkdKho01 (21 cột) |
| LƯU Ý | Bất thường tự động: MSKH không có trong HSKH, số HĐ không có trong Sổ cái, mã hàng thiếu thuế suất, TK mới chưa có quy tắc, KH thiếu MST & CCCD, HSKH bị lặp chuỗi… |
| NGUỒN FILE / THAM SỐ | Danh sách file đã đọc và cấu hình đã dùng |

## Một số điểm cần xác nhận

* **Tổ CP:** file mẫu không có dòng kết quả nào cho Tổ CP nên chưa đối chiếu được. Hiện đang lấy các dòng sổ quỹ chi
  có phân loại theo bảng TK và KH là HKD. Ví dụ các khoản 6418 / TK Có 1121 (mời cơm sinh nhật KH) **không** được
  tính và sẽ hiện ở sheet LƯU Ý — nếu cần tính thì thêm dòng vào `bang_phan_loai` trong cấu hình.
* **Sheet HSKH** trong file mẫu có một số ô bị nối đôi (vd “Võ CaVõ Ca”, MST lặp 2 lần); phần mềm tự sửa và báo ở LƯU Ý.

## Phát triển

```bash
python -m pytest -q
```
Mã nguồn: `thue_hkd/doc_file.py` (đọc & nhận diện), `tinh_toan.py` (tính), `xuat_excel.py` (ghi kết quả),
`config.py` (quy tắc), `app.py` (giao diện), `cli.py` (dòng lệnh).
