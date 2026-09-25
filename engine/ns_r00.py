"""NGÂN SÁCH R00 theo MẪU NỘI BỘ công ty — bộ đọc TẤT ĐỊNH (không dùng AI) + ghi vào khung (Excel COM).
Mẫu: sheet 'BCTC' (cột C mã NS · D tên · E doanh thu phân bổ · F ngân sách chi phí; dòng 'Lợi Nhuận…' = tổng) và
sheet '01. PhanTichBOQ' (B CODE công việc · C mã NS · F nội dung · Q ĐVT · R KL · S đơn giá · T thành tiền theo NS).
Phần NS của một mã chưa chi tiết theo CODE ⇒ thêm 1 dòng cân bằng ⇒ Σ theo từng mã NS luôn = BCTC (tự kiểm tới đồng)."""
import os, re, shutil, datetime as dt, unicodedata, openpyxl, warnings
warnings.filterwarnings("ignore")
XL_UP = -4162
TIEN = '#,##0;[Red]-#,##0;"–"'

def _so(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool): return float(v)
    try: return float(str(v).replace(",", "").strip())
    except Exception: return None
def _kd(s): return unicodedata.normalize("NFD", str(s or "").replace("Đ", "D").replace("đ", "d")).encode("ascii", "ignore").decode().lower()

def la_mau_r00(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True); ok = "BCTC" in wb.sheetnames and "01. PhanTichBOQ" in wb.sheetnames; wb.close(); return ok
    except Exception: return False

def doc_r00(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ma_ns, nhom_b, tong, tong_dt = [], "B.1", None, None
    for r in wb["BCTC"].iter_rows(min_row=9, max_row=120, max_col=8, values_only=True):
        t = _kd(r[3])
        if t.startswith("chi phi vat tu"): nhom_b = "B.2"
        elif t.startswith("chi phi gian tiep"): nhom_b = "B.3"
        elif t.startswith("chi phi du phong"): nhom_b = "B.4"
        if t.startswith("loi nhuan"): tong, tong_dt = _so(r[5]), _so(r[4])
        if r[2] and re.match(r"^[A-Za-z]+_", str(r[2]).strip()):
            m = str(r[2]).strip(); ma_ns.append(dict(ma=m, ten=str(r[3] or "").strip(), nhom=nhom_b, dt=_so(r[4]) or 0, ns=_so(r[5]) or 0))
    dong = []
    for i, r in enumerate(wb["01. PhanTichBOQ"].iter_rows(min_row=10, max_row=2000, max_col=21, values_only=True), 10):
        cv, mn, tt = r[1], r[2], _so(r[19])
        if cv and mn and tt is not None and re.match(r"^[A-Za-z]+_", str(mn).strip()):
            dong.append(dict(ma_ns=str(mn).strip(), ma_cv=str(cv).strip(), ds=str(r[5] or "").strip(), dvt=str(r[16] or r[6] or "").strip().replace("bao 50kg", "bao"),
                             kl=_so(r[17]), dg=_so(r[18]), tt=tt, nguon=f"R00 01.PhanTichBOQ dòng {i}"))
    wb.close()
    ds_ma = {m["ma"] for m in ma_ns}; lech_cv = sorted({d["ma_ns"] for d in dong} - ds_ma)
    for m in ma_ns:                                                   # cân bằng: phần NS chưa chi tiết theo CODE (chi phí gián tiếp, dòng nhóm…)
        s = sum(d["tt"] for d in dong if d["ma_ns"] == m["ma"]); con = round(m["ns"] - s, 2)
        if abs(con) >= 1: dong.append(dict(ma_ns=m["ma"], ma_cv=m["ma"], ds=m["ten"] if not s else f"{m['ten']} — phần NS chưa chi tiết theo mã công việc",
                                           dvt="lot", kl=1, dg=con, tt=con, nguon="R00 BCTC cột F (cân bằng theo mã NS)"))
    return dict(ma_ns=ma_ns, dong=dong, tong=tong, tong_dt=tong_dt, ma_la=lech_cv, tong_dong=sum(d["tt"] for d in dong))

def nhom_cv(dong):
    """Nhóm công việc = CODE công việc (mỗi CODE 1 nhóm) — khoá nối NS ↔ dòng HĐ."""
    out = {}
    for d in dong: out.setdefault(d["ma_cv"], dict(nhom=d["ma_cv"], ten=d["ds"][:80], dvt=d["dvt"], ma_ns=d["ma_ns"]))
    return list(out.values())

def ket_qua_ai_gia(path):
    """Kết quả dạng 'AI' để đi chung luồng hồ sơ nền (thẻ, cờ) — nhưng số lấy từ bộ đọc tất định."""
    k = doc_r00(path)
    return dict(loai="NGAN_SACH", ly_do_loai="Mẫu ngân sách R00 nội bộ (sheet BCTC + 01. PhanTichBOQ) — bộ đọc riêng, KHÔNG dùng AI",
                bang=[dict(stt=d["ma_cv"], noi_dung=d["ds"], dvt=d["dvt"], kl=d["kl"], don_gia=d["dg"], thanh_tien=d["tt"]) for d in k["dong"]],
                tong_ghi_tren_file=k["tong"], khong_chac=[f"Mã NS có trong BOQ nhưng không có trong BCTC: {', '.join(k['ma_la'])}"] if k["ma_la"] else [],
                ghi_chu=f"{len(k['ma_ns'])} mã NS · {len(k['dong'])} dòng · tổng NS {k['tong']:,.0f} · doanh thu phân bổ {k['tong_dt'] or 0:,.0f}", nguon={}), k

def ghi_r00(khung, path, ngay_hl, thu_muc_backup):
    import pythoncom, win32com.client as w32
    k = doc_r00(path); nh = nhom_cv(k["dong"])
    if not k["tong"] or abs(k["tong_dong"] - k["tong"]) > 1: return dict(ok=False, ly_do=f"Σ dòng {k['tong_dong']:,.0f} ≠ tổng BCTC {k['tong'] or 0:,.0f}")
    os.makedirs(thu_muc_backup, exist_ok=True)
    bk = os.path.join(thu_muc_backup, f"{os.path.splitext(os.path.basename(khung))[0]}_truoc_nhap_NS_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"); shutil.copy2(khung, bk)
    serial = (ngay_hl - dt.date(1899, 12, 30)).days
    pythoncom.CoInitialize(); app = w32.DispatchEx("Excel.Application"); app.Visible = False; app.DisplayAlerts = False; wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(khung)); w1, w2 = wb.Worksheets("N1_DanhMuc"), wb.Worksheets("N2_NganSach")
        if w1.Range("L3").Value or w2.Range("A2").Value:
            wb.Close(False); wb = None; return dict(ok=False, ly_do="Khung đã có ngân sách — điều chỉnh NS (R01, R02…) xử lý ở bước sau, không ghi đè R00", backup=bk)
        for j, m in enumerate(k["ma_ns"]):                            # N1 L:P mã NS
            r = 3 + j; ma = m["ma"]
            pp = "KL" if ma.startswith(("NTP_", "DTC_")) and any(d["ma_ns"] == ma and d["dvt"] != "lot" and (d["kl"] or 0) > 0 for d in k["dong"]) else "NS"
            for col, v in (("L", ma), ("M", m["ten"]), ("N", m["nhom"]), ("O", pp), ("P", m["dt"] or None)):
                if v not in (None, ""): w1.Range(f"{col}{r}").Value = v
            w1.Range(f"P{r}").NumberFormat = TIEN
        for j, n in enumerate(nh):                                    # N1 Q:T nhóm công việc
            r = 3 + j
            for col, v in (("Q", n["nhom"]), ("R", n["ten"]), ("S", n["dvt"]), ("T", n["ma_ns"])): w1.Range(f"{col}{r}").Value = v
        NH = "'N1_DanhMuc'!$Q$3:$T$400"
        for j, d in enumerate(k["dong"]):                             # N2 dòng NS R00
            r = 2 + j; khop = d["kl"] and d["dg"] and abs(d["kl"] * d["dg"] - d["tt"]) < 1
            for col, v in (("A", "R00"), ("B", serial), ("C", "GOC"), ("D", d["ma_ns"]), ("E", d["ma_cv"]), ("F", d["ma_cv"]), ("H", d["ds"]), ("I", d["dvt"]),
                           ("J", d["kl"] if khop else (d["kl"] or None)), ("K", d["dg"] if khop else None), ("N", d["nguon"])):
                if v not in (None, ""): w2.Range(f"{col}{r}").Value = v
            if khop: w2.Range(f"L{r}").Formula = f"=J{r}*K{r}"
            else: w2.Range(f"L{r}").Value = d["tt"]
            w2.Range(f"O{r}").Formula = (f'=IF(F{r}="","CHƯA GÁN NHÓM",IF(ISNA(VLOOKUP(F{r},{NH},3,0)),"NHÓM LẠ",IF(G{r}="","CHƯA GÁN GÓI",'
                                         f'IF(AND(I{r}<>"lot",I{r}<>VLOOKUP(F{r},{NH},3,0)),"LỆCH ĐVT",""))))')
            w2.Range(f"B{r}").NumberFormat = "dd/mm/yyyy"
            for col in ("J", "K", "L"): w2.Range(f"{col}{r}").NumberFormat = TIEN if col != "J" else "#,##0.00"
        app.CalculateFullRebuild()
        f = app.WorksheetFunction; loi = []
        tong = f.SumIfs(w2.Range("L2:L5000"), w2.Range("C2:C5000"), "GOC")
        if abs(tong - k["tong"]) > 1: loi.append(f"Σ N2 {tong:,.0f} ≠ BCTC {k['tong']:,.0f}")
        for m in k["ma_ns"]:
            s = f.SumIfs(w2.Range("L2:L5000"), w2.Range("D2:D5000"), m["ma"])
            if abs(s - m["ns"]) > 1: loi.append(f"{m['ma']}: {s:,.0f} ≠ {m['ns']:,.0f}")
        if loi:
            wb.Close(False); wb = None; return dict(ok=False, ly_do="Tự kiểm KHÔNG KHỚP — không lưu: " + "; ".join(loi[:5]), backup=bk)
        wb.Save(); wb.Close(False); wb = None
        return dict(ok=True, ma_hd="NS-R00", so_dong=len(k["dong"]), so_ma_ns=len(k["ma_ns"]), so_nhom=len(nh), tong=tong, backup=bk,
                    thong_bao=f"Đã ghi ngân sách R00: {len(k['ma_ns'])} mã NS · {len(nh)} nhóm công việc · {len(k['dong'])} dòng · Σ {tong:,.0f} (= BCTC) · đã backup")
    finally:
        if wb is not None: wb.Close(False)
        app.Quit()
