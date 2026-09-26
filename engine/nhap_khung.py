"""GHI HỒ SƠ NỀN ĐÃ DUYỆT VÀO KHUNG (Excel COM) — hợp đồng CĐT / đối tác: N1 (đối tác mới, thông tin dự án) + N4/N6 (dòng HĐ) + N5/N7 (dòng đơn giá, có công thức).
Quy tắc như ghi sổ HSTT: backup trước · DispatchEx (không đụng Excel đang mở, CẤM taskkill) · tính lại · TỰ KIỂM (giá trị HĐ = AI đọc, Σ dòng = Σ bảng)
· lệch ⇒ đóng KHÔNG LƯU. Khung trống không có dòng mẫu ⇒ công thức dựng từ cong_thuc.py."""
import os, shutil, datetime as dt
import pythoncom, win32com.client as w32
import cong_thuc as T
XL_UP = -4162
N1_LOAI = {"DTC": "ĐTC", "NTP": "NTP", "NCC": "NCC", "DVK": "DVK", "CDT": "CĐT"}
TIEN, PT, NG = '#,##0;[Red]-#,##0;"–"', "0.0%", "dd/mm/yyyy"

def _ngay(s):
    """Ngày ⇒ số serial Excel. KHÔNG truyền datetime qua COM: pywin32 coi là UTC ⇒ lùi 7 giờ ⇒ hiện ngày hôm trước."""
    try: return (dt.datetime.strptime(str(s)[:10], "%Y-%m-%d").date() - dt.date(1899, 12, 30)).days
    except Exception: return None
def _dong_cuoi(ws, col="A", tu=1):
    r = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
    return max(r, tu)
def _tim(ws, col, gt, tu, den=5000):
    for r in range(tu, _dong_cuoi(ws, col, tu) + 1):
        if str(ws.Range(f"{col}{r}").Value or "").strip() == gt: return r
    return None

def dong_hop_le(bang):
    """Dòng lá của bảng đơn giá (bỏ dòng nhóm / dòng cộng). Dòng chỉ có thành tiền ⇒ KL 1, đơn giá = thành tiền (gói)."""
    import re; out = []
    for x in bang or []:
        if not (x.get("dvt") or x.get("kl") is not None) or re.match(r"^\s*(cộng|tổng)", str(x.get("noi_dung") or ""), re.I): continue
        kl, dg, tt = x.get("kl"), x.get("don_gia"), x.get("thanh_tien"); ghi = ""
        if kl is None and dg is None and tt is not None: kl, dg = 1, tt
        v = kl * dg if isinstance(kl, (int, float)) and isinstance(dg, (int, float)) else None
        if isinstance(tt, (int, float)) and (v is None or abs(v - tt) > max(1000, abs(tt) * 0.001)):   # KL×ĐG ≠ thành tiền in trên HĐ ⇒ tin THÀNH TIỀN
            ghi = f"HĐ ghi KL {kl} × ĐG {dg}; thành tiền theo HĐ {tt:,.0f}"; kl, dg = 1, tt
        out.append(dict(stt=str(x.get("stt") or len(out) + 1), noi_dung=x.get("noi_dung"), dvt=x.get("dvt") or "gói", kl=kl, don_gia=dg, ghi=ghi))
    return out

def ghi_hd(khung, da, rec, sua, thu_muc_backup):
    """rec = bản ghi hồ sơ nền (loai HD_CDT / HD_DOI_TAC, có rec['ai']). sua = dict ghi đè trường AI (anh điền chỗ trống). Trả dict kết quả."""
    a = dict(rec["ai"] or {}); a.update({k: v for k, v in (sua or {}).items() if v not in (None, "")})
    import nen; bad = [c for c in nen.chuan_hoa(a) + nen.danh_gia(a, rec["loai"], None, {})[0] if c["muc"] == "CHAN"]                # chuẩn hoá lại lần nữa ngay trước khi ghi (phòng thủ)
    if bad: return dict(ok=False, ly_do="; ".join(c["mo_ta"] for c in bad))
    cdt = rec["loai"] == "HD_CDT"; sh_hd, sh_ct = ("N4_HD_CDT", "N5_BOQ_CDT") if cdt else ("N6_HD_DoiTac", "N7_HD_DoiTac_ChiTiet")
    ma_dt = rec.get("ma_doi_tac") or ""; loai_dt = "CDT" if cdt else (rec.get("loai_doi_tac") or a.get("loai_doi_tac"))
    if not ma_dt: return dict(ok=False, ly_do="Chưa xác định mã đối tác")
    if not cdt and loai_dt not in ("DTC", "NTP", "NCC", "DVK"): return dict(ok=False, ly_do="Chưa xác định loại đối tác (đội / thầu phụ / NCC / dịch vụ)")
    dong = dong_hop_le(a.get("bang")); tong = sum((d["kl"] or 0) * (d["don_gia"] or 0) for d in dong)
    os.makedirs(thu_muc_backup, exist_ok=True)
    bk = os.path.join(thu_muc_backup, f"{os.path.splitext(os.path.basename(khung))[0]}_truoc_nhap_HD_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"); shutil.copy2(khung, bk)
    pythoncom.CoInitialize(); app = w32.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False; wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(khung)); w1, wh, wc = wb.Worksheets("N1_DanhMuc"), wb.Worksheets(sh_hd), wb.Worksheets(sh_ct)
        ma_hd = "HD-CDT" if cdt else f"HD-{ma_dt}"
        if _tim(wh, "A", ma_hd, 2):
            wb.Close(False); wb = None
            return dict(ok=False, ly_do=f"{ma_hd} đã có trong khung — phụ lục/điều chỉnh HĐ sẽ xử lý ở bước sau, không ghi đè", backup=bk)
        # N1: thông tin dự án (khung trống) + đối tác mới
        if str(w1.Range("A3").Value or "") in ("", "MAU_TRONG"): w1.Range("A3").Value = da
        if str(w1.Range("B3").Value or "") in ("", "(tên dự án)") and a.get("du_an"): w1.Range("B3").Value = a["du_an"]
        if cdt and not w1.Range("C3").Value: w1.Range("C3").Value = ma_dt
        dt_moi = not _tim(w1, "G", ma_dt, 3)
        if dt_moi:
            r = _dong_cuoi(w1, "G", 2) + 1
            for col, v in (("G", ma_dt), ("H", a.get("doi_tac_ten") or ma_dt), ("I", N1_LOAI.get(loai_dt, loai_dt)), ("J", f"thêm khi nhập HĐ {a.get('so_hd') or ''}".strip())):
                w1.Range(f"{col}{r}").Value = v
        # N4/N6: dòng hợp đồng
        r = _dong_cuoi(wh, "A", 1) + 1; cot = T.cot(T.COT_HD[sh_hd])
        gt = {"ma_hd": ma_hd, "loai_ghi": "GOC", "ma_doi_tac": ma_dt, "ma_goi": rec.get("goi") or "", "so_hd": a.get("so_hd"), "ngay_ky": _ngay(a.get("ngay_ky")),
              "noi_dung": a.get("noi_dung"), "dang_hd": a.get("dang_hd"), "gia_tri_truoc_vat": a.get("gia_tri_truoc_vat"), "vat_pct": a.get("vat_pct"),
              "pct_tam_ung": a.get("pct_tam_ung"), "pct_tt_dot": a.get("pct_tt_dot"), "pct_tt_quyet_toan": a.get("pct_tt_quyet_toan"),
              "han_tt_ngay": a.get("han_tt_ngay"), "don_vi_han": a.get("don_vi_han"), "han_qt_ngay": a.get("han_qt_ngay"), "han_tra_gl_ngay": a.get("han_tra_gl_ngay"),
              "bao_hanh_thang": a.get("bao_hanh_thang"), "trang_thai": "DANG_TH",
              "nguon": a.get("nguon") or f"AI đọc {rec['ten'][:60]} · anh duyệt {dt.datetime.now():%d/%m/%Y}", "ghi_chu": (a.get("ghi_chu") or "")[:250],
              "ho_so_thieu": a.get("ho_so_thieu")}
        for k, v in gt.items():
            if k in cot and v not in (None, ""): wh.Range(f"{cot[k]}{r}").Value = v
        for k, f in (("ngay_ky", NG), ("gia_tri_truoc_vat", TIEN), ("vat_pct", PT), ("pct_tam_ung", PT), ("pct_tt_dot", PT), ("pct_tt_quyet_toan", PT)):
            wh.Range(f"{cot[k]}{r}").NumberFormat = f
        # N5/N7: dòng đơn giá + công thức
        c = T.cot(T.CT_COLS); r0 = _dong_cuoi(wc, "A", 1) + 1
        for j, d in enumerate(dong):
            rr = r0 + j
            for k, v in (("ma_hd", ma_hd), ("stt", d["stt"]), ("pham_vi", "TRONG_HD"), ("noi_dung", d["noi_dung"]), ("dvt", d["dvt"]), ("kl_hd", d["kl"]),
                         ("don_gia", d["don_gia"]), ("nguon", f"AI đọc {rec['ten'][:40]}" + (f" · {d['ghi']}" if d.get("ghi") else ""))):
                if v not in (None, ""): wc.Range(f"{c[k]}{rr}").Value = v
            for col, f in T.ct_cong_thuc(sh_ct, rr).items(): wc.Range(f"{col}{rr}").Formula = f
            wc.Range(f"{c['don_gia']}{rr}").NumberFormat = TIEN; wc.Range(f"{c['thanh_tien']}{rr}").NumberFormat = TIEN
        app.CalculateFullRebuild()
        # TỰ KIỂM
        f = app.WorksheetFunction
        gt_khung = wh.Range(f"{cot['gia_tri_truoc_vat']}{r}").Value or 0
        tong_khung = f.SumIfs(wc.Range(f"{c['thanh_tien']}2:{c['thanh_tien']}5000"), wc.Range("A2:A5000"), ma_hd)
        loi = []
        if a.get("gia_tri_truoc_vat") and abs(gt_khung - a["gia_tri_truoc_vat"]) > 1: loi.append(f"giá trị HĐ trong khung {gt_khung:,.0f} ≠ đọc được {a['gia_tri_truoc_vat']:,.0f}")
        if abs(tong_khung - tong) > 1: loi.append(f"Σ dòng trong khung {tong_khung:,.0f} ≠ Σ bảng {tong:,.0f}")
        g = a.get("gia_tri_truoc_vat")
        if g and dong and abs(tong_khung - g) > g * 0.005: loi.append(f"Σ dòng {tong_khung:,.0f} ≠ giá trị HĐ {g:,.0f} (lệch {abs(tong_khung - g) / g:.1%})")
        if loi:
            wb.Close(False); wb = None
            return dict(ok=False, ly_do="Tự kiểm sau khi ghi KHÔNG KHỚP (" + "; ".join(loi) + ") — KHÔNG lưu, file giữ nguyên", backup=bk)
        wb.Save(); wb.Close(False); wb = None
        return dict(ok=True, ma_hd=ma_hd, doi_tac_moi=dt_moi, so_dong=len(dong), tong_dong=tong_khung, gia_tri=gt_khung, backup=bk,
                    thong_bao=f"Đã ghi {ma_hd} vào khung: {len(dong)} dòng đơn giá, Σ {tong_khung:,.0f}" + (" · thêm đối tác mới vào N1" if dt_moi else "") + " · đã backup")
    finally:
        if wb is not None: wb.Close(False)
        app.Quit()

def ghi_phu_luc(khung, da, rec, thu_muc_backup):
    """PHỤ LỤC HĐ: N4/N6 thêm dòng loai_ghi PHU_LUC dưới mã HĐ gốc; mỗi đơn giá của PL ⇒ 1 dòng N5/N7 MỚI (STT PLn.xx, nội dung '— giá <số PL> từ <ngày>'),
    giữ nguyên dòng giá cũ để đối chiếu HSTT theo thời điểm. HĐ gốc phải có sẵn; cùng số PL đã ghi ⇒ không ghi lần 2."""
    a = dict(rec["ai"] or {}); cdt = rec["loai"] == "HD_CDT"
    sh_hd, sh_ct = ("N4_HD_CDT", "N5_BOQ_CDT") if cdt else ("N6_HD_DoiTac", "N7_HD_DoiTac_ChiTiet")
    ma_dt = rec.get("ma_doi_tac") or ""; ma_hd = "HD-CDT" if cdt else f"HD-{ma_dt}"; so_pl = str(a.get("so_hd") or "PL").strip()
    ngay = _ngay(a.get("ngay_ky")); ngay_txt = dt.datetime.strptime(str(a.get("ngay_ky"))[:10], "%Y-%m-%d").strftime("%d/%m/%Y") if ngay else "?"
    dong = [d for d in dong_hop_le(a.get("bang")) if d["don_gia"] is not None]
    os.makedirs(thu_muc_backup, exist_ok=True)
    bk = os.path.join(thu_muc_backup, f"{os.path.splitext(os.path.basename(khung))[0]}_truoc_nhap_PL_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"); shutil.copy2(khung, bk)
    pythoncom.CoInitialize(); app = w32.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False; wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(khung)); wh, wc = wb.Worksheets(sh_hd), wb.Worksheets(sh_ct); cot = T.cot(T.COT_HD[sh_hd])
        if not _tim(wh, "A", ma_hd, 2):
            wb.Close(False); wb = None; return dict(ok=False, ly_do=f"Chưa có HĐ gốc {ma_hd} trong khung — duyệt HĐ gốc trước", backup=bk)
        so_pl_cu = [str(wh.Range(f"{cot['so_hd']}{r}").Value or "").strip() for r in range(2, _dong_cuoi(wh, "A", 1) + 1)
                    if wh.Range(f"A{r}").Value == ma_hd and wh.Range(f"B{r}").Value == "PHU_LUC"]
        if so_pl in so_pl_cu:
            wb.Close(False); wb = None; return dict(ok=False, ly_do=f"Phụ lục {so_pl} của {ma_hd} đã ghi trước đó — không ghi lần 2", backup=bk)
        n = len(so_pl_cu) + 1; r = _dong_cuoi(wh, "A", 1) + 1
        for k, v in (("ma_hd", ma_hd), ("loai_ghi", "PHU_LUC"), ("ma_doi_tac", ma_dt), ("so_hd", so_pl), ("ngay_ky", ngay), ("noi_dung", (a.get("noi_dung") or "")[:250]),
                     ("dang_hd", a.get("dang_hd")), ("gia_tri_truoc_vat", a.get("gia_tri_truoc_vat")), ("vat_pct", a.get("vat_pct")), ("trang_thai", "DANG_TH"),
                     ("nguon", f"AI đọc {rec['ten'][:60]} · anh duyệt {dt.datetime.now():%d/%m/%Y}"), ("ghi_chu", (a.get("ghi_chu") or "")[:250])):
            if k in cot and v not in (None, ""): wh.Range(f"{cot[k]}{r}").Value = v
        wh.Range(f"{cot['ngay_ky']}{r}").NumberFormat = NG; wh.Range(f"{cot['gia_tri_truoc_vat']}{r}").NumberFormat = TIEN
        c = T.cot(T.CT_COLS); r0 = _dong_cuoi(wc, "A", 1) + 1
        for j, d in enumerate(dong):
            rr = r0 + j
            for k, v in (("ma_hd", ma_hd), ("stt", f"PL{n}.{d['stt']}"), ("pham_vi", "TRONG_HD"), ("noi_dung", f"{d['noi_dung']} — giá {so_pl} từ {ngay_txt}"),
                         ("dvt", d["dvt"]), ("kl_hd", d["kl"] if d["kl"] != 1 or d.get("ghi") else None), ("don_gia", d["don_gia"]), ("nguon", f"PL {so_pl} · AI đọc {rec['ten'][:40]}")):
                if v not in (None, ""): wc.Range(f"{c[k]}{rr}").Value = v
            for col, f in T.ct_cong_thuc(sh_ct, rr).items(): wc.Range(f"{col}{rr}").Formula = f
            wc.Range(f"{c['don_gia']}{rr}").NumberFormat = TIEN
        app.CalculateFullRebuild()
        so = sum(1 for rr in range(r0, r0 + len(dong)) if wc.Range(f"A{rr}").Value == ma_hd and str(wc.Range(f"B{rr}").Value).startswith(f"PL{n}."))
        if so != len(dong) or wh.Range(f"B{r}").Value != "PHU_LUC":
            wb.Close(False); wb = None; return dict(ok=False, ly_do=f"Tự kiểm không khớp ({so}/{len(dong)} dòng) — KHÔNG lưu", backup=bk)
        wb.Save(); wb.Close(False); wb = None
        return dict(ok=True, ma_hd=ma_hd, phu_luc=so_pl, so_dong=len(dong), backup=bk,
                    thong_bao=f"Đã ghi phụ lục {so_pl} vào {ma_hd}: dòng PHU_LUC (N{'4' if cdt else '6'}) + {len(dong)} đơn giá mới (STT PL{n}.xx) · đã backup")
    finally:
        if wb is not None: wb.Close(False)
        app.Quit()
