"""KHAI BÁO TAY tạm ứng / hoàn ứng CĐT — anh Phương yêu cầu 01/10/2026 (trước đó phải nhờ viết script, giờ có nút trên web).
Ghi 1 dòng N8_TT_CDT loại TAM_UNG (số dương) hoặc HOAN_UNG (số âm) — không qua BoQ, không ảnh hưởng lũy kế doanh thu (chỉ lũy kế TAM_UNG/HOAN_UNG
dùng để tính 'dư tạm ứng còn lại' ở cdt.kiem_cdt/ke_hoach_cdt). Dùng khi CĐT chuyển/khấu trừ tạm ứng mà không qua hồ sơ thanh toán chuẩn."""
import datetime as dt, openpyxl, warnings
warnings.filterwarnings("ignore")

def ds_hop_dong_cdt(khung):
    """[(mã HĐ, số HĐ, dư tạm ứng còn lại hiện tại)] — chỉ HĐ GỐC phía CĐT (tạm ứng gắn với HĐ gốc, không gắn PL/KT/VT/PHAT)."""
    wb = openpyxl.load_workbook(khung, read_only=True, data_only=True)
    try:
        import kiem as K
        k = K.doc_khung(khung); b = k["cdt"]
        return [dict(ma_hd=m, so_hd=h["so_hd"], tu_treo=round(b["tu_treo"].get(m, 0.0), 2))
                for m, h in b["hd"].items() if not any(m.endswith(x) for x in ("-KT", "-VT", "-PHAT"))]
    finally: wb.close()

def ghi(khung, backup, ma_hd, loai, so_tien, ngay, ghi_chu="", nguoi="anh"):
    """Ghi 1 khoản tạm ứng (loai=TAM_UNG, so_tien > 0) hoặc hoàn ứng (loai=HOAN_UNG, so_tien > 0 — tự đổi dấu âm khi ghi). Trả dict kết quả."""
    import kiem as K, ghi_so as G
    so_tien = float(so_tien)
    if so_tien <= 0: raise ValueError("Số tiền phải lớn hơn 0 (chọn Hoàn ứng nếu là khoản trừ)")
    if loai not in ("TAM_UNG", "HOAN_UNG"): raise ValueError("Loại phải là TAM_UNG hoặc HOAN_UNG")
    try: ngay_d = dt.date.fromisoformat(str(ngay)[:10])
    except ValueError: raise ValueError("Ngày không hợp lệ (YYYY-MM-DD)")
    k = K.doc_khung(khung)
    if ma_hd not in k["cdt"]["hd"]: raise ValueError(f"'{ma_hd}' không có trong N4_HD_CDT của khung — chưa có HĐ gốc CĐT này")
    tu_truoc = k["cdt"]["tu_treo"].get(ma_hd, 0.0)
    if loai == "HOAN_UNG" and so_tien > tu_truoc + 1:
        raise ValueError(f"Hoàn ứng {so_tien:,.0f} VƯỢT dư tạm ứng còn lại {tu_truoc:,.0f} — anh kiểm lại số tiền hoặc số dư trước khi ghi")
    dau = so_tien if loai == "TAM_UNG" else -so_tien
    x = dict(ma_hd=ma_hd, dot=None, ngay=ngay_d, loai=loai, stt=None, kl=None, dg=None, so_tien=dau,
              ghi=f"{'Tạm ứng' if loai == 'TAM_UNG' else 'Hoàn ứng'} CĐT" + (f" · {ghi_chu.strip()[:100]}" if ghi_chu and ghi_chu.strip() else "")
                 + f" · khai báo tay bởi {nguoi} {dt.datetime.now():%d/%m/%Y %H:%M}")
    ph = dict(tt="N8_TT_CDT", ct="N5_BOQ_CDT", ma_hd=ma_hd, dong_moi=[], dong_tt=[x], lk_hstt=None)
    kq = G.ghi(khung, dict(phan=[ph]), f"{'tạm ứng' if loai == 'TAM_UNG' else 'hoàn ứng'} CĐT tay · {ma_hd}", backup)
    if not kq["ok"]: raise ValueError(kq["ly_do"])
    tu_sau = tu_truoc + dau
    return dict(ok=True, ma_hd=ma_hd, loai=loai, so_tien=so_tien, ngay=ngay_d.isoformat(), tu_truoc=round(tu_truoc, 2), tu_sau=round(tu_sau, 2), backup=kq["backup"],
                thong_bao=f"Đã ghi {'tạm ứng' if loai == 'TAM_UNG' else 'hoàn ứng'} {so_tien:,.0f}đ vào {ma_hd} · ngày {ngay_d:%d/%m/%Y} · dư tạm ứng còn lại {tu_sau:,.0f}đ (trước đó {tu_truoc:,.0f}đ) · đã backup")
