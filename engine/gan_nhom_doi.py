"""Gán NHÓM CV (⇒ mã NS) cho dòng HĐ đối tác chưa gán (N7 cột K) — bảng luật anh Phương duyệt 28/09/2026.
Nhóm mới DOI_<mã NS>_<ĐVT> (ĐVT trùng đúng dòng HĐ ⇒ không 'LỆCH ĐVT'). Chạy: python gan_nhom_doi.py <khung.xlsx> [--ghi]"""
import sys, os, re, shutil, datetime as dt, unicodedata, pythoncom, win32com.client as w32
sys.stdout.reconfigure(encoding="utf-8")
na = lambda s: re.sub(r"\s+", " ", unicodedata.normalize("NFD", str(s or "").replace("Đ", "D").replace("đ", "d")).encode("ascii", "ignore").decode().upper()).strip()
LUAT = [(None, r"^GACH OP LAT|^CA MAY"),                                   # chưa có mục NS khớp ⇒ để CHƯA GÁN
        ("DTC_CanNen", r"LANG NEN"),                       # TRƯỚC luật xi măng: láng nền vữa xi măng là CÔNG TÁC
        ("NCC_KeoChaRon", r"\bKEO\b"),
        ("NCC_Gach", r"^GACH "),
        ("NCC_VLXD", r"^CAT\b|XI MANG|DA 1X2"),
        ("NCC_VTP", r"HOAN UNG BCH"),
        ("Prelim_9.1", r"TRAC DA"),
        ("Prelim_7", r"HO TRO"),
        ("Prelim_9.2", r"CONG NHAT|DOI TRUONG|TANG CA"),
        ("Prelim_1", r"VAN PHONG BCH|MAI CHE VAN THANG|NHA VAN PHONG"),
        ("Prelim_3", r"HO SO AN TOAN"),
        ("DTC_LTBT", r"LANH TO|BO TRU|THEP DAI"),
        ("DTC_Trat", r"TRAT|LUOI CHONG NUT"),
        ("DTC_VanChuyen", r"XA BAN"),
        ("DTC_Khac1", r"CHONG NONG|LUOI THEP HAN|XUONG CA|MAI CAU THANG|MAI BE TONG"),
        ("DTC_OpLat", r"LAT GACH|OP LAT|LEN CHAN TUONG"),
        ("NTP_Xay", r"XAY TUONG|XAY GACH|CHAN KE LAN CAN")]
def ma_ns(ten):
    t = na(ten)
    for m, rx in LUAT:
        if re.search(rx, t): return m, True
    return None, False

def nhom_cho(noi_dung, dvt, hd_noi_dung=None):
    """(mã nhóm, dòng N1 Q:T) cho 1 dòng HĐ đội mới — None nếu bảng luật chưa có mục khớp.
    HĐ VẬN CHUYỂN (nội dung HĐ có 'vận chuyển'): dòng mang tên vật tư (xi măng, cát, gạch…) là CƯỚC ⇒ DTC_VanChuyen."""
    m, _ = ma_ns(noi_dung)
    if hd_noi_dung and "VAN CHUYEN" in na(hd_noi_dung) and m and m.startswith("NCC_"): m = "DTC_VanChuyen"
    if m is None and hd_noi_dung and "VAN CHUYEN" in na(hd_noi_dung) and re.search(r"GACH|XI MANG|CAT|DA |KEO", na(noi_dung)): m = "DTC_VanChuyen"
    if not m: return None, None
    nh = f"DOI_{m}_" + re.sub(r"[^A-Za-z0-9]", "", na(dvt))
    return nh, (nh, f"Đội/NTP → {m} ({dvt})", str(dvt or "").strip(), m)

if __name__ == "__main__":
    khung = sys.argv[1]; ghi = "--ghi" in sys.argv
    if ghi: os.makedirs(os.path.join(os.path.dirname(khung), "_backup"), exist_ok=True); shutil.copy2(khung, os.path.join(os.path.dirname(khung), "_backup", f"{os.path.splitext(os.path.basename(khung))[0]}_truoc_gan_nhom_doi_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"))
    pythoncom.CoInitialize(); xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False; wb = xl.Workbooks.Open(os.path.abspath(khung))
    try:
        w1, w7, w9 = wb.Worksheets("N1_DanhMuc"), wb.Worksheets("N7_HD_DoiTac_ChiTiet"), wb.Worksheets("N9_TT_DoiTac"); f = xl.WorksheetFunction
        ten_ns = {str(w1.Range(f"L{r}").Value): str(w1.Range(f"M{r}").Value) for r in range(3, 80) if w1.Range(f"L{r}").Value}
        co_nhom = {str(w1.Range(f"Q{r}").Value) for r in range(3, w1.Cells(w1.Rows.Count, 17).End(-4162).Row + 1)}
        them, gan, bo = {}, 0, []
        for r in range(2, w7.Cells(w7.Rows.Count, 1).End(-4162).Row + 1):
            if not w7.Range(f"A{r}").Value or w7.Range(f"K{r}").Value: continue
            nd, dvt = w7.Range(f"E{r}").Value, str(w7.Range(f"F{r}").Value or "").strip()
            m, khop = ma_ns(nd)
            if not m: bo.append(f"{w7.Range(f'A{r}').Value}·{str(nd)[:30]}"); continue
            nh = ("HU_" if na(nd).startswith("HOAN UNG BCH") else "DOI_") + m + "_" + re.sub(r"[^A-Za-z0-9]", "", na(dvt)) if not na(nd).startswith("HOAN UNG BCH") else f"HU_{m}"
            if nh not in co_nhom and nh not in them: them[nh] = (nh, f"{'Hoàn ứng BCH' if nh.startswith('HU_') else 'Đội/NTP'} → {ten_ns.get(m, m)} ({dvt})", dvt, m)
            w7.Range(f"K{r}").Value = nh; gan += 1
            if nh.startswith("HU_") and "CHỜ ANH" in str(nd).upper(): w7.Range(f"E{r}").Value = "Hoàn ứng BCH — Vật tư phụ / dụng cụ thi công"
        r1 = w1.Cells(w1.Rows.Count, 17).End(-4162).Row + 1
        for j, v in enumerate(them.values()):
            for col, x in zip("QRST", v): w1.Range(f"{col}{r1 + j}").Value = x
        xl.CalculateFullRebuild()
        loi = [f"N7 dòng {r}: {w7.Range(f'N{r}').Value}" for r in range(2, w7.Cells(w7.Rows.Count, 1).End(-4162).Row + 1)
               if str(w7.Range(f"K{r}").Value or "").startswith(("DOI_", "HU_")) and str(w7.Range(f"N{r}").Value or "") in ("NHÓM LẠ", "LỆCH ĐVT", "CHƯA GÁN NHÓM")]
        tong = sum(f.SumIfs(w9.Range("K2:K5000"), w9.Range("D2:D5000"), t) for t in ("THUC_HIEN", "DIEU_CHINH"))
        chua = sum(f.SumIfs(w9.Range("K2:K5000"), w9.Range("D2:D5000"), t, w9.Range("M2:M5000"), "") for t in ("THUC_HIEN", "DIEU_CHINH"))
        ns = {}
        for r in range(2, w9.Cells(w9.Rows.Count, 1).End(-4162).Row + 1):
            if w9.Range(f"D{r}").Value in ("THUC_HIEN", "DIEU_CHINH") and w9.Range(f"M{r}").Value: ns[w9.Range(f"M{r}").Value] = ns.get(w9.Range(f"M{r}").Value, 0) + (w9.Range(f"K{r}").Value or 0)
        print(f"gán {gan} dòng · thêm {len(them)} nhóm · để CHƯA GÁN {len(bo)}: {bo}")
        print(f"chi phí thực hiện {tong:,.0f} = có mã NS {sum(ns.values()):,.0f} + chưa gán {chua:,.0f}  ⇒ {'KHỚP' if abs(tong - sum(ns.values()) - chua) < 2 else 'LỆCH'}")
        for k_, v in sorted(ns.items(), key=lambda x: -x[1]): print(f"   {k_:<14} {v:>16,.0f}")
        print("lỗi kiểm tra N7:", loi[:8] or "không")
        if ghi and not loi and abs(tong - sum(ns.values()) - chua) < 2: wb.Save(); print("ĐÃ LƯU")
        elif ghi: print("KHÔNG LƯU (tự kiểm chưa đạt)")
    finally: wb.Close(False); xl.Quit()
