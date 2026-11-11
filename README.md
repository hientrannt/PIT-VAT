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

### Bản chạy trên trình duyệt (không cần cài gì)
File `web/thue_hkd.html`: mở trực tiếp bằng Chrome/Edge (hoặc qua link Claude đã chia sẻ). Cùng logic với bản Python,
đã đối chiếu khớp 5.849/5.849 dòng. Dữ liệu chỉ xử lý trong trình duyệt.

### Windows (dễ nhất)
1. Cài Python 3.10 trở lên từ <https://www.python.org/downloads/> (nhớ tick **Add python.exe to PATH**).
2. Tải thư mục này về, nhấp đúp **`chay_phan_mem.bat`**. Lần đầu sẽ mất vài phút để cài thư viện.
3. Trình duyệt tự mở trang phần mềm (chạy trên máy bạn, dữ liệu không gửi đi đâu).
4. Chọn tháng/năm → kéo-thả các file → kiểm tra bảng “loại dữ liệu” → bấm **Tính thuế** → **Tải file Excel kết quả**.

### macOS
Nhấp đúp **`chay_tren_mac.command`**: tự cài thư viện lần đầu, tính toàn bộ thư mục `2026_Thue HKD` (tự lấy kỳ),
ghi file kết quả vào chính thư mục đó và mở ra. Giao diện web: `./chay_phan_mem.sh`.

### Dòng lệnh (chạy cả thư mục)
```bash
pip install -r requirements.txt
```
```bash
python cli.py "D:/ThueHKD/T9" -o "D:/ThueHKD/KetQua_T9.xlsx"   # tự lấy kỳ; thêm --thang 9 --nam 2026 để chỉ định
```

## Tháng/năm & dữ liệu trùng

* **Kỳ tính lấy theo ngày**: sổ quỹ & voucher theo *Ngày chứng từ*; **kho 03 (Xuất KM)
  theo ngày nhận hàng** (*Ngay NH*, thiếu thì *NgayGS*); **kho 01 (dữ liệu kho) theo ngày nhận hàng** (*Ngày giao hàng*). Đổi được ở `cot_ngay` trong cấu hình.
  Mọi nguồn (kể cả cấn trừ, Tổ CP) đều lọc theo kỳ.
* **Ngày dạng chữ** đọc được: dd/mm/yyyy, d/m/yy, dd-mm-yyyy, dd.mm.yyyy, yyyy-mm-dd, có giờ phía sau,
  dấu ' ở đầu, khoảng trắng ẩn, 9-Sep-2026, 20260909, "ngày 09 tháng 09 năm 2026". Mặc định **ngày trước tháng**.
  Ngày không đọc được / để trống được báo riêng ở LƯU Ý (mục NGÀY) kèm ví dụ, không âm thầm bị loại.
* Không chọn kỳ (CLI không truyền `--thang/--nam`; trang web khi thêm file) → **tự chọn tháng gần nhất có dữ liệu** (file thường chứa cả năm).
* Một thư mục để nhiều file (cả file tổng hợp cũ lẫn file xuất rời, hoặc 2 lần xuất chồng ngày): dòng giống hệt
  dòng đã có ở file khác cùng loại **bị bỏ, không cộng hai lần**, và báo ở LƯU Ý (mục TRÙNG).
* Sheet **KIỂM TRA DỮ LIỆU**: mỗi nguồn đọc được bao nhiêu dòng, được tính bao nhiêu, bị loại vì sao, dữ liệu thuộc tháng nào.

## Đặt tên file

Phần mềm nhận loại file theo **dòng tiêu đề cột**, không theo tên file. Tên file chỉ quan trọng với **sổ quỹ**:
sổ quỹ cấn trừ (Tổ DT) và sổ quỹ chi (Tổ CP) có cùng mẫu cột, nên phân biệt bằng tên:

* Tên chứa `CAN TRU` / `CAN_TRU` / `cấn trừ` → Sổ quỹ – Cấn trừ (Tổ DT)
* Tên chứa `DATA CHI` / `DATA_CHI` / `TO CP` / `TO_CP` / `Tổ CP` → Sổ quỹ – Chi (Tổ CP)
* Không có các chữ trên: đa số phiếu bắt đầu bằng `PC` → Tổ CP, còn lại → Cấn trừ.

Gợi ý đặt tên thống nhất: `T09.2026_<LOẠI>[_phần].xlsx`, ví dụ `T09.2026_SO_QUY_CAN_TRU.xlsx`,
`T10.2026_SO_QUY_TO_CP.xlsx`, `T09.2026_XUAT_KM_1.xlsx`, `T09.2026_XUAT_KM_2.xlsx`, `T09.2026_VOUCHER.xlsx`,
`T09.2026_SO_CAI.xlsx`, `HSKH.xlsx`, `BANG_THUE.xlsx`, `T09.2026_BAN_HANG_KHO_NTG01.xlsx`.

## Đưa lên mạng để có link riêng (Vercel / Cloudflare Pages)

Trang web là 1 file tĩnh `web/index.html` (tạo từ `web/thue_hkd.html` bằng `python scripts/tao_index.py`).
Không có máy chủ: dữ liệu vẫn chỉ xử lý trong trình duyệt người dùng.

**Cloudflare Pages – kéo-thả, không cần GitHub:** dash.cloudflare.com → Workers & Pages → Create → Pages →
*Upload assets* → đặt tên dự án → kéo-thả thư mục `web` (hoặc file zip) → Deploy → link `https://<tên>.pages.dev`.

**Vercel – nối với GitHub (tự cập nhật khi repo thay đổi):** vercel.com/new → Import `hientrannt/PIT-VAT` →
giữ nguyên cài đặt (đã có `vercel.json`, thư mục xuất là `web`) → Deploy → link `https://<tên>.vercel.app`.

**Cloudflare Pages – nối với GitHub:** Create → Pages → *Connect to Git* → chọn repo → Build command để trống,
Build output directory: `web` → Save and Deploy.

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

* **Tổ CP:** lấy các dòng sổ quỹ chi có phân loại theo bảng TK và KH là HKD. Các khoản 6418 / TK Có 1121
  (mời cơm sinh nhật KH) và 33882 (chi lại tiền CKTT) **không tính thuế** (đã xác nhận) và không báo ở LƯU Ý
  (`cap_tk_khong_tinh` trong cấu hình).
* **Sheet HSKH** trong file mẫu có một số ô bị nối đôi (vd “Võ CaVõ Ca”, MST lặp 2 lần); phần mềm tự sửa và báo ở LƯU Ý.

## Phát triển

```bash
python -m pytest -q
```
Mã nguồn: `thue_hkd/doc_file.py` (đọc & nhận diện), `tinh_toan.py` (tính), `xuat_excel.py` (ghi kết quả),
`config.py` (quy tắc), `app.py` (giao diện), `cli.py` (dòng lệnh).
