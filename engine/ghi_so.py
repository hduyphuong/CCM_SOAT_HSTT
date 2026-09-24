"""GHI SỔ 1 HSTT đã duyệt vào FILE KHUNG (Excel COM).
Quy tắc: backup trước · DispatchEx (không đụng Excel anh đang mở, CẤM taskkill) · chép dòng mẫu để giữ công thức + định dạng ·
tính lại rồi TỰ KIỂM lũy kế = HSTT (≤10đ) · lệch ⇒ đóng KHÔNG LƯU (ghi đủ hoặc không ghi gì)."""
import os, shutil, datetime as dt
import pythoncom, win32com.client as w32
from doc_hstt import na
NGUONG = 10
XL_UP = -4162

def ke_hoach(hs, kq, k):
    """Tính các dòng sẽ ghi (chưa đụng file) — để anh xem trước khi bấm Đồng ý."""
    ma = kq["phan_loai"]["ma_hd"]; cv = hs["cover"]; dot = cv["dot"]; ngay = cv["ngay"]
    ps_co = [d["stt"] for d in k["dong"].get(ma, []) if str(d["stt"]).startswith("PS")]
    n_ps = max([int(s[2:]) for s in ps_co if s[2:].isdigit()] or [0])
    dong_moi, tt, gan, gop, thu_tu = [], [], {}, {}, []
    for l, st in kq["khop"]:
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
            tt.append(dict(loai="THUC_HIEN", stt=st, kl=l["kl_kn"], dg=l["dg"], so_tien=None, ghi=""))
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

def ghi(path, kh, ten_file, thu_muc_backup):
    os.makedirs(thu_muc_backup, exist_ok=True)
    bk = os.path.join(thu_muc_backup, f"{os.path.splitext(os.path.basename(path))[0]}_truoc_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx")
    shutil.copy2(path, bk)
    pythoncom.CoInitialize()
    app = w32.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False
    wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(path))
        w7, w9 = wb.Worksheets("N7_HD_DoiTac_ChiTiet"), wb.Worksheets("N9_TT_DoiTac")
        ma = kh["ma_hd"]
        # dòng HĐ phát sinh mới (N7) — chép dòng cuối để giữ công thức
        for d in kh["dong_hd_moi"]:
            last = w7.Cells(w7.Rows.Count, 1).End(XL_UP).Row; r = last + 1
            w7.Range(f"A{last}:O{last}").Copy(w7.Range(f"A{r}"))
            for col, v in (("A", ma), ("B", d["stt"]), ("D", "NGOAI_HD"), ("E", d["noi_dung"]), ("F", d["dvt"]), ("G", None),
                           ("H", d["don_gia"]), ("J", ""), ("K", ""), ("O", f"webapp — phát sinh từ {ten_file[:40]}")):
                w7.Range(f"{col}{r}").Value = v
        # dòng thanh toán (N9)
        for x in kh["dong_tt"]:
            last = w9.Cells(w9.Rows.Count, 1).End(XL_UP).Row; r = last + 1
            w9.Range(f"A{last}:R{last}").Copy(w9.Range(f"A{r}"))
            ngay = dt.datetime(x["ngay"].year, x["ngay"].month, x["ngay"].day) if x["ngay"] else None
            dong = x["stt"] is not None
            vals = {"A": ma, "B": x["dot"], "C": ngay, "D": x["loai"], "E": x["stt"] if dong else None,
                    "F": f"=IFERROR(VLOOKUP(A{r}&\"|\"&E{r},'N7_HD_DoiTac_ChiTiet'!$C:$E,3,0),\"\")" if dong else x["ghi"],
                    "G": f"=IFERROR(VLOOKUP(A{r}&\"|\"&E{r},'N7_HD_DoiTac_ChiTiet'!$C:$F,4,0),\"\")" if dong else None,
                    "H": x["kl"], "I": x["dg"], "J": x["so_tien"], "P": None, "R": f"webapp · {ten_file[:50]}" + (f" · {x['ghi']}" if x["ghi"] and dong else "")}
            for col, v in vals.items():
                c = w9.Range(f"{col}{r}")
                if isinstance(v, str) and v.startswith("="): c.Formula = v
                else: c.Value = v
        app.CalculateFullRebuild()
        # tự kiểm: lũy kế thực hiện của HĐ trong khung = lũy kế trên HSTT
        lk = app.WorksheetFunction.SumIfs(w9.Range("K2:K5000"), w9.Range("A2:A5000"), ma, w9.Range("D2:D5000"), "THUC_HIEN") + \
             app.WorksheetFunction.SumIfs(w9.Range("K2:K5000"), w9.Range("A2:A5000"), ma, w9.Range("D2:D5000"), "DIEU_CHINH")
        if abs(lk - kh["lk_hstt"]) > NGUONG:
            wb.Close(False); wb = None
            return dict(ok=False, ly_do=f"Sau khi ghi, lũy kế trong khung {lk:,.0f} ≠ HSTT {kh['lk_hstt']:,.0f} — KHÔNG lưu, file giữ nguyên", backup=bk)
        wb.Save(); wb.Close(False); wb = None
        return dict(ok=True, luy_ke_khung=lk, so_dong_tt=len(kh["dong_tt"]), so_dong_hd_moi=len(kh["dong_hd_moi"]), backup=bk)
    finally:
        if wb is not None: wb.Close(False)
        app.Quit()
