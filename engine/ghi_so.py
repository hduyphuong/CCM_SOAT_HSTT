"""GHI SỔ 1 hồ sơ đã duyệt vào FILE KHUNG (Excel COM).
Quy tắc: backup trước · DispatchEx (không đụng Excel anh đang mở, CẤM taskkill) · chép dòng mẫu để giữ công thức + định dạng ·
tính lại rồi TỰ KIỂM lũy kế = hồ sơ (≤10đ) cho từng phần · lệch ⇒ đóng KHÔNG LƯU (ghi đủ hoặc không ghi gì).
Kế hoạch gồm nhiều PHẦN: {tt: sheet thanh toán, ct: sheet dòng HĐ, ma_hd, dong_moi, dong_tt, lk_hstt}."""
import os, shutil, datetime as dt
import pythoncom, win32com.client as w32
from doc_hstt import na
import cong_thuc as T
NGUONG = 10
XL_UP = -4162

def ke_hoach(hs, kq, k):
    """Tính các dòng sẽ ghi (chưa đụng file) — để anh xem trước khi bấm Đồng ý."""
    if hs.get("loai") == "CDT":
        import cdt; return cdt.ke_hoach_cdt(hs, kq, k)
    kh = _ke_hoach_doi_tac(hs, kq, k)
    kh["phan"] = [dict(tt="N9_TT_DoiTac", ct="N7_HD_DoiTac_ChiTiet", ma_hd=kh["ma_hd"], dong_moi=kh["dong_hd_moi"], dong_tt=kh["dong_tt"], lk_hstt=kh["lk_hstt"])]
    return kh

def _ke_hoach_doi_tac(hs, kq, k):
    ma = kq["phan_loai"]["ma_hd"]; cv = hs["cover"]; dot = cv["dot"]; ngay = cv["ngay"]
    ps_co = [d["stt"] for d in k["dong"].get(ma, []) if str(d["stt"]).startswith("PS")]
    n_ps = max([int(s[2:]) for s in ps_co if s[2:].isdigit()] or [0])
    dong_moi, tt, gan, gop, thu_tu = [], [], {}, {}, []
    for l, st in kq["khop"]:
        if st is None and (l.get("nhom_moi") or hs.get("mau") == "HOAN_UNG_BCH"):                          # hoàn ứng BCH: dòng theo MÃ NS, HĐ tạm ⇒ thêm TRONG_HD, không phải phát sinh
            st = l["stt"]
            if st not in gan: gan[st] = st; dong_moi.append(dict(stt=st, noi_dung=l["ds"], dvt=l["dvt"], don_gia=l["dg"], pham_vi="TRONG_HD", nhom=l.get("nhom"),
                                                                 nhom_moi=l["nhom_moi"], nguon="webapp · hoàn ứng BCH (dòng theo mã NS)"))
        if st is None:
            key = na(l["ds"])
            if key not in gan:
                n_ps += 1; gan[key] = f"PS{n_ps}"
                dong_moi.append(dict(stt=gan[key], noi_dung=l["ds"], dvt=l["dvt"], don_gia=l["dg"]))
            st = gan[key]
        g = gop.setdefault(st, dict(kt=0.0, vt=0.0))                  # 1 dòng HĐ có thể nằm ở nhiều khung giá (NCC) ⇒ GỘP rồi mới so kỳ trước
        g["kt"] += l["kl_kt"]; g["vt"] += l["tt_lk"] - l["tt_kn"]; thu_tu.append((l, st))
    for st, g in gop.items():
        dkl = g["kt"] - k["lk_kl"].get((ma, st), 0.0); dv = g["vt"] - k["lk_tien_dong"].get((ma, st), 0.0)
        if abs(dkl) > 1e-6: tt.append(dict(loai="DIEU_CHINH", stt=st, kl=dkl, dg=round(dv / dkl, 4), so_tien=None, ghi="HSTT sửa kỳ trước"))
        elif abs(dv) > 1: tt.append(dict(loai="DIEU_CHINH", stt=st, kl=None, dg=None, so_tien=round(dv, 2), ghi="HSTT đổi giá kỳ trước"))
    for l, st in thu_tu:
        if abs(l["kl_kn"]) > 1e-9:
            tt.append(dict(loai="THUC_HIEN", stt=st, kl=l["kl_kn"], dg=l["dg"], so_tien=None, ghi="", **({"vat_rieng": l["vat_rieng"]} if "vat_rieng" in l else {})))
            if abs(l["kl_kn"] * l["dg"] - l["tt_kn"]) > 1: tt.append(dict(loai="DIEU_CHINH", stt=st, kl=None, dg=None, so_tien=round(l["tt_kn"] - l["kl_kn"] * l["dg"], 2), ghi="làm tròn"))
    O = k["tu_treo"][ma]
    if kq["phan_loai"]["loai_hs"] == "TAM_UNG":
        d = hs["du_tru"] - O
        if abs(d) > 1: tt.append(dict(loai="TAM_UNG" if d > 0 else "HOAN_UNG", stt=None, kl=None, dg=None, so_tien=d, ghi="Tạm ứng giữa kỳ" if d > 0 else "Hoàn ứng"))
    else:
        rec = -hs["hu_k"] if hs["hu_k"] else (-hs["tu_k"] if hs["tu_k"] < 0 else 0)
        moi = hs["du_tru"] - O + rec
        if moi > 1: tt.append(dict(loai="TAM_UNG", stt=None, kl=None, dg=None, so_tien=moi, ghi="Tạm ứng (thể hiện ở HSTT này)"))
        if rec > 1: tt.append(dict(loai="HOAN_UNG", stt=None, kl=None, dg=None, so_tien=-rec, ghi="Hoàn ứng"))
    for x in tt: x.update(ma_hd=ma, dot=dot, ngay=ngay)
    return dict(ma_hd=ma, dong_hd_moi=dong_moi, dong_tt=tt, lk_hstt=(hs["tong"] or (0, 0, 0))[2])

def _ghi_phan(wb, ph, ten_file):
    w7, w9, ma = wb.Worksheets(ph["ct"]), wb.Worksheets(ph["tt"]), ph["ma_hd"]
    w1 = wb.Worksheets("N1_DanhMuc")
    for d in ph["dong_moi"]:                                          # nhóm CV mới (hoàn ứng BCH → mã NS) ⇒ thêm vào N1 Q:T trước khi dòng HĐ tra
        if d.get("nhom_moi") and not any(str(w1.Range(f"Q{r}").Value or "") == d["nhom_moi"][0] for r in range(3, w1.Cells(w1.Rows.Count, 17).End(XL_UP).Row + 1)):
            r = w1.Cells(w1.Rows.Count, 17).End(XL_UP).Row + 1
            for col, v in zip("QRST", d["nhom_moi"]): w1.Range(f"{col}{r}").Value = v
    for d in ph["dong_moi"]:                                          # dòng HĐ phát sinh mới — chép dòng cuối để giữ công thức
        last = w7.Cells(w7.Rows.Count, 1).End(XL_UP).Row; r = last + 1
        if last >= 2: w7.Range(f"A{last}:O{last}").Copy(w7.Range(f"A{r}"))
        else:                                                         # sheet trống: không có dòng mẫu ⇒ dựng công thức
            for col, f in T.ct_cong_thuc(ph["ct"], r).items(): w7.Range(f"{col}{r}").Formula = f
        for col, v in (("A", ma), ("B", d["stt"]), ("D", d.get("pham_vi") or "NGOAI_HD"), ("E", d["noi_dung"]), ("F", d["dvt"]), ("G", None),
                       ("H", d["don_gia"]), ("J", ""), ("K", d.get("nhom") or ""), ("O", d.get("nguon") or f"webapp — phát sinh từ {ten_file[:40]}")):
            w7.Range(f"{col}{r}").Value = v
    for x in ph["dong_tt"]:                                           # dòng thanh toán
        last = w9.Cells(w9.Rows.Count, 1).End(XL_UP).Row; r = last + 1
        if last >= 2: w9.Range(f"A{last}:R{last}").Copy(w9.Range(f"A{r}"))
        else:
            for col, f in T.tt_cong_thuc(ph["tt"], r, co_noi_dung=True).items(): w9.Range(f"{col}{r}").Formula = f
            for col, f in (("C", "dd/mm/yyyy"), ("O", "dd/mm/yyyy"), ("K", '#,##0;[Red]-#,##0;"–"'), ("N", '#,##0;[Red]-#,##0;"–"')): w9.Range(f"{col}{r}").NumberFormat = f
        ngay = (dt.date(x["ngay"].year, x["ngay"].month, x["ngay"].day) - dt.date(1899, 12, 30)).days if x["ngay"] else None   # serial: COM đổi datetime theo múi giờ ⇒ lùi 1 ngày
        dong = x["stt"] is not None
        vals = {"A": ma, "B": x["dot"], "C": ngay, "D": x["loai"], "E": x["stt"] if dong else None,
                "F": f"=IFERROR(VLOOKUP(A{r}&\"|\"&E{r},'{ph['ct']}'!$C:$E,3,0),\"\")" if dong else x["ghi"],
                "G": f"=IFERROR(VLOOKUP(A{r}&\"|\"&E{r},'{ph['ct']}'!$C:$F,4,0),\"\")" if dong else None,
                "H": x["kl"], "I": x["dg"], "J": x["so_tien"], "P": None, "R": f"webapp · {ten_file[:50]}" + (f" · {x['ghi']}" if x["ghi"] and dong else "")}
        for col, v in vals.items():
            c = w9.Range(f"{col}{r}")
            if isinstance(v, str) and v.startswith("="): c.Formula = v
            else: c.Value = v
        if ph["tt"] == "N9_TT_DoiTac":                                # tiền chi: đặt lại công thức CHUẨN (dòng chép có thể mang thuế suất riêng của dòng trên)
            cn = T.cot(T.TT_COLS)["tien_thanh_toan"]; f = T.tt_cong_thuc(ph["tt"], r, co_noi_dung=True)[cn]
            if x.get("vat_rieng") is not None:                        # dòng CÓ hoá đơn trong HĐ không-VAT (hoàn ứng BCH): VAT của HĐ ⇒ thuế suất của dòng
                i = f.index("*(1+") + 4; j = f.index(")*", i); f = f[:i] + f"{x['vat_rieng']:.10f}" + f[j:]
            w9.Range(f"{cn}{r}").Formula = f

def _lk(app, w9, ma):
    f = app.WorksheetFunction
    return sum(f.SumIfs(w9.Range("K2:K5000"), w9.Range("A2:A5000"), ma, w9.Range("D2:D5000"), t) for t in ("THUC_HIEN", "DIEU_CHINH"))

def ghi(path, kh, ten_file, thu_muc_backup):
    os.makedirs(thu_muc_backup, exist_ok=True)
    bk = os.path.join(thu_muc_backup, f"{os.path.splitext(os.path.basename(path))[0]}_truoc_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx")
    shutil.copy2(path, bk)
    pythoncom.CoInitialize()
    app = w32.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False
    wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(path))
        for ph in kh["phan"]: _ghi_phan(wb, ph, ten_file)
        app.CalculateFullRebuild()
        lk_chinh = None
        for ph in kh["phan"]:                                         # tự kiểm từng phần có số đối chiếu
            if ph["lk_hstt"] is None: continue
            lk = _lk(app, wb.Worksheets(ph["tt"]), ph["ma_hd"]); lk_chinh = lk if lk_chinh is None else lk_chinh
            if abs(lk - ph["lk_hstt"]) > NGUONG:
                wb.Close(False); wb = None
                return dict(ok=False, ly_do=f"Sau khi ghi, lũy kế {ph['ma_hd']} trong khung {lk:,.0f} ≠ hồ sơ {ph['lk_hstt']:,.0f} — KHÔNG lưu, file giữ nguyên", backup=bk)
        wb.Save(); wb.Close(False); wb = None
        return dict(ok=True, luy_ke_khung=lk_chinh, so_dong_tt=sum(len(p["dong_tt"]) for p in kh["phan"]),
                    so_dong_hd_moi=sum(len(p["dong_moi"]) for p in kh["phan"]), backup=bk)
    finally:
        if wb is not None: wb.Close(False)
        app.Quit()
