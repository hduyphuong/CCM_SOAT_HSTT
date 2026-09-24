"""Đọc 1 file HSTT (đội / NTP / NCC) — tự dò cột theo tiêu đề, chạy được các dạng đã gặp:
05.Giá trị (mới + cũ có mã CV) · BẢNG KL (Tình, NCC chia khung giá / đợt giao). Không ghi gì, chỉ đọc."""
import openpyxl, re, unicodedata, datetime as dt, hashlib, warnings
warnings.filterwarnings("ignore")

def na(s):
    s = str(s if s is not None else "").replace("Đ", "D").replace("đ", "d")
    return re.sub(r"\s+", " ", unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode().upper()).strip()
def num(v): return float(v) if isinstance(v, (int, float)) else 0.0
def van_tay(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""): h.update(b)
    return h.hexdigest()

def _ngay(v):
    if isinstance(v, dt.datetime): return v.date()
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(v or ""))
    return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None

def doc_cover(wb):
    """Tên đơn vị · số HĐ · số đợt · ngày HSTT (vị trí ô) từ sheet COVER; ngày năm < 2026 ⇒ lấy 'ĐẾN NGÀY' trên BIA."""
    kq = dict(ten_don_vi=None, so_hd=None, dot=None, ngay=None, o={})
    sc = next((s for s in wb.sheetnames if "COVER" in na(s)), None)
    if sc:
        for r_i, row in enumerate(wb[sc].iter_rows(min_row=1, max_row=12, max_col=10, values_only=True), 1):
            cells = [c for c in row]
            for j, c in enumerate(cells):
                t = na(c)
                if t.startswith("TEN NTP") and kq["ten_don_vi"] is None:
                    kq["ten_don_vi"] = str(c).split(":", 1)[-1].strip(); kq["o"]["ten_don_vi"] = f"{sc}!{openpyxl.utils.get_column_letter(j+1)}{r_i}"
                if t.startswith("SO HOP DONG") and kq["so_hd"] is None:
                    v = next((x for x in cells[j+1:] if x not in (None, "")), None)
                    if v: kq["so_hd"] = str(v).strip(); kq["o"]["so_hd"] = f"{sc}!{openpyxl.utils.get_column_letter(cells.index(v)+1)}{r_i}"
                if t in ("SO:", "SO :") and kq["dot"] is None:
                    v = next((x for x in cells[j+1:] if x not in (None, "")), None)
                    try: kq["dot"] = int(str(v).strip()); kq["o"]["dot"] = f"{sc}!{r_i}"
                    except (TypeError, ValueError): pass
                if t.startswith("NGAY") and kq["ngay"] is None and r_i <= 4:
                    v = next((x for x in cells[j+1:] if x not in (None, "")), None); d = _ngay(v)
                    if d and d.year >= 2026: kq["ngay"] = d
    if kq["ngay"] is None:
        sb = next((s for s in wb.sheetnames if "BIA" in na(s)), None)
        if sb:
            for row in wb[sb].iter_rows(min_row=1, max_row=40, max_col=11, values_only=True):
                for c in row:
                    m = re.search(r"DEN NGAY (\d{1,2})/(\d{1,2})/(\d{4})", na(c))
                    if m: kq["ngay"] = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))); break
                if kq["ngay"]: break
    return kq

def doc_file(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    import cdt
    if cdt.la_claim_cdt(wb):                                   # hồ sơ thanh toán GỬI CĐT (doanh thu)
        kq = cdt.doc_claim(path, wb); wb.close(); return kq
    cv = doc_cover(wb)
    sn = next((s for s in wb.sheetnames if na(s) in ("05.GIA TRI", "BANG KL") or re.fullmatch(r"\d+\.BANG KL", na(s))), None)
    if not sn: wb.close(); raise ValueError("Không thấy sheet bảng giá trị / bảng KL (05.Giá trị hoặc 5.BẢNG KL)")
    rows = [list(r) for r in wb[sn].iter_rows(min_row=1, max_row=200, max_col=30, values_only=True)]
    txt = []
    sc = next((s for s in wb.sheetnames if "COVER" in na(s)), None)
    for r in (wb[sc].iter_rows(min_row=1, max_row=4, max_col=10, values_only=True) if sc else []): txt += [na(c) for c in r if c]
    wb.close()
    h = next(i for i, r in enumerate(rows) if any(na(c) == "DIEN GIAI" for c in r))
    txt += [na(c) for r in rows[:min(h, 6)] for c in r if c and any(k in na(c) for k in ("DU AN", "CONG TRINH"))]
    H, S = [na(c) for c in rows[h]], [na(c) for c in rows[h + 1]]
    col = lambda lab, arr: next((j for j, c in enumerate(arr) if c.startswith(lab)), None)
    c_stt, c_ds, c_dg = col("STT", H), col("DIEN GIAI", H), col("DON GIA", H)
    c_dv = next(j for j, c in enumerate(H) if c == "DVT" or c.startswith("DON VI"))
    kt = [j for j, c in enumerate(S) if c in ("KI TRUOC", "KY TRUOC")]; kn = [j for j, c in enumerate(S) if c in ("KI NAY", "KY NAY")]
    lk = [j for j, c in enumerate(S) if c == "LUY KE"]
    c_klhd = col("KL (HD)", H); c_klhd = kt[0] - 1 if c_klhd is None else c_klhd
    kq = dict(file=path, sheet=sn, cover=cv, lines=[], tong=None, vat=None, tu=0.0, hu=0.0, tu_k=0.0, hu_k=0.0, gl=None)
    khung, ngoai = "", False
    for i, r in enumerate(rows[h + 2:], h + 3):
        a = na(r[c_stt]); ds = r[c_ds]; dsn = na(ds); tag = a if a.startswith("(") else ""
        if tag == "(DNTT)" or dsn.startswith("DE NGHI THANH TOAN"): break      # hết bảng — phía dưới là chữ ký + bảng phụ (Bê tông/Coffa/Cốt thép)
        if dsn.startswith("TONG CONG (CHUA") or a == "GTTH": kq["tong"] = (num(r[kt[1]]), num(r[kn[1]]), num(r[lk[1]])); continue
        if tag:
            if tag == "(VAT)": kq["vat"] = next((x for x in r[c_dv:c_dv + 3] if isinstance(x, (int, float))), None)
            if tag in ("(GL)", "(BH)"): kq["gl"] = num(r[lk[1]])
            if tag == "(TU)": kq["tu"] = num(r[lk[1]]); kq["tu_k"] = num(r[kn[1]])
            if tag == "(HU)": kq["hu"] = num(r[lk[1]]); kq["hu_k"] = num(r[kn[1]])
            continue
        co_so = any(num(r[j]) for j in (kt[0], kn[0], lk[0], kn[1], lk[1]))
        if ds and not r[c_dv] and not isinstance(r[c_dg], (int, float)):
            khung = str(ds).strip()
            if any(k in dsn for k in ("NGOAI HOP DONG", "PHU LUC", "PHAT SINH")): ngoai = True
            elif a in ("III", "C", "D", "E"): ngoai = False
            continue
        if ds and isinstance(r[c_dg], (int, float)) and (r[c_dv] or (r[c_dg] > 0 and co_so)) and not a.isalpha():   # dòng tổng nhóm: ĐG=0, không ĐVT ⇒ bỏ
            kq["lines"].append(dict(dong=i, khung=khung, stt=a, ds=str(ds).strip(), dvt=str(r[c_dv] or "").strip(), kl_hd=r[c_klhd] if isinstance(r[c_klhd], (int, float)) else None,
                dg=num(r[c_dg]), kl_kt=num(r[kt[0]]), kl_kn=num(r[kn[0]]), kl_lk=num(r[lk[0]]),
                tt_kt=num(r[kt[1]]), tt_kn=num(r[kn[1]]), tt_lk=num(r[lk[1]]), ngoai=ngoai))
    kq["du_tru"] = round(kq["tu"] + kq["hu"]); kq["du_an_text"] = " ".join(txt)
    return kq
