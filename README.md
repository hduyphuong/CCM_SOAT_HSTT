# Soát HSTT — webapp quản lý chi phí & hợp đồng (V1)

Trang web (GitHub Pages) + **engine chạy trên máy anh**. Hồ sơ thanh toán chỉ đi qua trang tới engine ở `127.0.0.1`; **không có dữ liệu nào lưu lên mạng**. Repo này chỉ chứa code.

## Chạy
1. Bấm đúp `chay_engine.bat` (hoặc `python engine\app.py`) → engine chạy tại `http://127.0.0.1:8765`.
2. Mở trang (`docs/index.html` trên GitHub Pages, hoặc chạy thử tại máy: `python -m http.server 5500 --directory docs` rồi mở `http://localhost:5500`).
3. Lần đầu Chrome hỏi **"cho phép truy cập thiết bị trong mạng nội bộ"** → bấm **Cho phép** (Chrome 142+ bắt buộc, chỉ hỏi 1 lần).

## Luồng
Nạp HSTT (đội / NTP / NCC) **hoặc hồ sơ thanh toán gửi CĐT (doanh thu)** → phân loại (HĐ, đối tác, đợt, loại hồ sơ) → tự kiểm 4 lớp (hồ sơ · theo HĐ · số học · đợt trước + ngân sách) → anh duyệt
(**Đồng ý** ghi sổ · **Yêu cầu chỉnh sửa** · **Trả đội**; còn cờ CHẶN thì khoá Đồng ý) → ghi sổ vào file khung Excel → báo cáo.

Engine chỉ nhận lệnh từ trang ghi trong `<DATA>\_CAU_HINH\trang.txt` (VD `https://ten-tai-khoan.github.io`) và từ `localhost`.

## Chạy engine
- **Chạy ẩn, tự bật khi đăng nhập Windows** (khuyên dùng): Task Scheduler tác vụ `Engine soat HSTT` → `pythonw engine\chay_an.pyw` (không cửa sổ, log ở `<DATA>\_CAU_HINH\log\engine.log`, đã chạy thì không bật bản thứ 2). Bật tay: `schtasks /run /tn "Engine soat HSTT"` · tắt: `tat_engine.bat`.
- Chạy có cửa sổ (gỡ lỗi): `chay_engine.bat` — đóng cửa sổ là engine tắt.

## Vị trí dữ liệu (`<DATA>`)
Ghi ở **một file trên máy chạy engine**, không lên git: `engine\cau_hinh_may.json`
```json
{"DATA": "<thư mục dữ liệu — ổ máy, hoặc thư mục Google Drive đồng bộ>"}
```
Đổi ổ / đổi máy / đổi công ty ⇒ chuyển nguyên thư mục `<DATA>` rồi **sửa đúng 1 dòng này** và khởi động lại engine.
Mọi đường dẫn bên trong `<DATA>` (file khung, hồ sơ đã nạp, hồ sơ đã xếp) đều lưu **tương đối** nên không phải sửa gì thêm.
Để trên Drive: chỉ **một máy** chạy engine ghi sổ; người khác được chia sẻ quyền **Người xem**.

## Cây thư mục `<DATA>` (xem `engine/cay.py`)
```
_CAU_HINH\        du_an.json · trang.txt · log\
_DUNG_CHUNG\      hồ sơ dùng cho nhiều dự án (HĐ nguyên tắc NCC, bảng giá chung)
<MÃ_DỰ_ÁN>\
  00_KHUNG\       file khung Excel CCM v3 (+ _backup\ trước mỗi lần ghi sổ)
  01_THIET_LAP\   01_BOQ_CDT · 02_NGAN_SACH · 03_GOI_THAU\CHON_THAU
  10_THU_CDT\     HOP_DONG · HSTT\Dxx_YYYYMMDD
  20_CHI_DOI_TAC\ <DTC|NTP|NCC|DVK>_<mã đối tác>\ HOP_DONG · BAO_GIA · HSTT\Dxx_YYYYMMDD · QUYET_TOAN
  30_BAO_CAO\     <YYYY-MM-DD>
  _HE_THONG\      nap\ (bản gốc đã nạp, chỉ đọc) · so_nap.json — engine quản lý
```
`_CAU_HINH\du_an.json` — mỗi dự án trỏ tới file khung (đường dẫn tương đối so với `<DATA>`):
```json
{"DU_AN_A": {"ten": "Tên dự án", "khung": "DU_AN_A\\00_KHUNG\\CCM_DU_AN_A.xlsx"}}
```
HSTT được duyệt và ghi sổ thành công ⇒ engine tự chép bản gốc vào `…\HSTT\Dxx_YYYYMMDD\` của đúng CĐT / đối tác.

## An toàn ghi sổ
Backup trước · Excel COM `DispatchEx` (không đụng Excel đang mở) · chép dòng mẫu giữ công thức · tính lại và **tự kiểm lũy kế = HSTT (≤ 10 đ)** · lệch ⇒ **không lưu**, file giữ nguyên · kiểm lại ngay trước khi ghi · hồ sơ đã ghi sổ không ghi lần 2.

## Mã nguồn
| File | Việc |
|---|---|
| `engine/doc_hstt.py` | đọc HSTT đội / NTP / NCC (tự dò cột theo tiêu đề) |
| `engine/kiem.py` | đọc khung · phân loại · 4 lớp kiểm |
| `engine/ghi_so.py` | kế hoạch ghi + ghi sổ bằng Excel COM |
| `engine/cdt.py` | phía CĐT: đọc hồ sơ gửi CĐT, kiểm với BOQ HĐ CĐT; khấu trừ/phạt/cấp vật tư ghi CHI bên đối tác |
| `engine/app.py` | API (stdlib, không cần cài thêm gói) |
| `docs/index.html` | giao diện |
