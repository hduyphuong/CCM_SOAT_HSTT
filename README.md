# Soát HSTT — webapp quản lý chi phí & hợp đồng (V1)

Trang web (GitHub Pages) + **engine chạy trên máy anh**. Hồ sơ thanh toán chỉ đi qua trang tới engine ở `127.0.0.1`; **không có dữ liệu nào lưu lên mạng**. Repo này chỉ chứa code.

## Chạy
1. Bấm đúp `chay_engine.bat` (hoặc `python engine\app.py`) → engine chạy tại `http://127.0.0.1:8765`.
2. Mở trang (`docs/index.html` trên GitHub Pages, hoặc chạy thử tại máy: `python -m http.server 5500 --directory docs` rồi mở `http://localhost:5500`).
3. Lần đầu Chrome hỏi **"cho phép truy cập thiết bị trong mạng nội bộ"** → bấm **Cho phép** (Chrome 142+ bắt buộc, chỉ hỏi 1 lần).

## Luồng
Nạp HSTT (đội / NTP / NCC) **hoặc hồ sơ thanh toán gửi CĐT (doanh thu)** → phân loại (HĐ, đối tác, đợt, loại hồ sơ) → tự kiểm 4 lớp (hồ sơ · theo HĐ · số học · đợt trước + ngân sách) → anh duyệt
(**Đồng ý** ghi sổ · **Yêu cầu chỉnh sửa** · **Trả đội**; còn cờ CHẶN thì khoá Đồng ý) → ghi sổ vào file khung Excel → báo cáo.

Engine chỉ nhận lệnh từ trang ghi trong `WEBAPP_SOAT_HSTT_DATA\trang.txt` (VD `https://ten-tai-khoan.github.io`) và từ `localhost`.

## Cấu hình dự án
`D:\QLCP_HD\WEBAPP_SOAT_HSTT_DATA\du_an.json` — mỗi dự án trỏ tới 1 file khung CCM v3:
```json
{"DU_AN_A": {"ten": "Tên dự án", "khung": "<đường dẫn file khung CCM v3 .xlsx>"}}
```
Dữ liệu app (file đã nạp — chỉ đọc, sổ nạp, backup trước mỗi lần ghi sổ): `D:\QLCP_HD\WEBAPP_SOAT_HSTT_DATA\<dự án>\`.

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
