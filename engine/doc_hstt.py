"""Đọc 1 file HSTT (đội / NTP / NCC) — tự dò cột theo tiêu đề, chạy được các dạng đã gặp:
05.Giá trị (mới + cũ có mã CV) · BẢNG KL (Tình, NCC chia khung giá / đợt giao). Không ghi gì, chỉ đọc."""
import openpyxl, re, os, unicodedata, datetime as dt, hashlib, warnings
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
    if not sn and la_hoan_ung(wb):                             # hồ sơ HOÀN ỨNG quỹ BCH (tờ trình + bảng kê + phiếu tạm ứng + giấy thanh toán tạm ứng)
        kq = doc_hoan_ung(path, wb); wb.close(); return kq
    if not sn: wb.close(); raise ValueError("Không thấy sheet bảng giá trị / bảng KL (05.Giá trị hoặc 5.BẢNG KL)")
    rows = [list(r) for r in wb[sn].iter_rows(min_row=1, max_row=200, max_col=30, values_only=True)]
    txt = []
    sc = next((s for s in wb.sheetnames if "COVER" in na(s)), None)
    for r in (wb[sc].iter_rows(min_row=1, max_row=4, max_col=10, values_only=True) if sc else []): txt += [na(c) for c in r if c]
    wb.close()
    h = next(i for i, r in enumerate(rows) if any(na(c) == "DIEN GIAI" for c in r))
    txt += [na(c) for r in rows[:min(h, 6)] for c in r if c and any(k in na(c) for k in ("DU AN", "CONG TRINH"))]
    H, S = [na(c) for c in rows[h]], [na(c) for c in rows[h + 1]]
    if "CONG TAC" in H and any(c.startswith("NGAY") for c in H):      # mẫu HSTT CÔNG NHẬT ⇒ bộ đọc riêng
        return doc_cong_nhat(path, sn, rows, h, H, cv, txt)
    col = lambda lab, arr: next((j for j, c in enumerate(arr) if c.startswith(lab)), None)
    c_stt, c_ds, c_dg = col("STT", H), col("DIEN GIAI", H), col("DON GIA", H)
    c_dv = next(j for j, c in enumerate(H) if c == "DVT" or c.startswith("DON VI"))
    kt = [j for j, c in enumerate(S) if c in ("KI TRUOC", "KY TRUOC")]; kn = [j for j, c in enumerate(S) if c in ("KI NAY", "KY NAY")]
    lk = [j for j, c in enumerate(S) if c == "LUY KE"]
    if not (kt and kn and lk):                                 # mẫu khác (VD HSTT CÔNG NHẬT: bảng KL theo ngày + bảng chấm công) ⇒ báo rõ, không lỗi kỹ thuật
        cn = any("CHAM CONG" in na(x) for x in (wb.sheetnames if False else [])) or any("NGAY" == c for c in H)
        raise ValueError(("Mẫu HSTT CÔNG NHẬT (bảng KL ghi theo ngày, có bảng chấm công) — " if cn else "Mẫu HSTT chưa hỗ trợ — ")
                         + f"sheet '{sn}' không có cột KỲ TRƯỚC / KỲ NÀY / LŨY KẾ nên app chưa đọc được. Hồ sơ CHƯA vào sổ nạp (không soát, không ghi sổ); báo em để thêm bộ đọc mẫu này.")
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


def doc_cong_nhat(path, sn, rows, h, H, cv, txt):
    """HSTT CÔNG NHẬT: bảng KL ghi theo NGÀY (chức danh · ngày · công tác · ĐVT · KL · ĐG · thành tiền · %TT) + bảng chấm công.
    Gom theo CHỨC DANH + ĐƠN GIÁ thành dòng HSTT. Nhóm 'Đợt N' = đợt hiện tại ⇒ KỲ NÀY, nhóm đợt trước ⇒ KỲ TRƯỚC; không có nhóm ⇒ tất cả kỳ này.
    Lũy kế đối chiếu với 'Giá trị thực hiện lũy kế' trên COVER (lk_cover)."""
    col = lambda lab: next((j for j, c in enumerate(H) if c.startswith(lab)), None)
    c_stt, c_ds, c_ng, c_dv, c_kl, c_dg, c_tt = col("STT"), col("DIEN GIAI"), col("NGAY"), col("DVT"), col("KL"), col("DON GIA"), col("THANH TIEN")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True); lk_cover = None; dot_data = ngay_data = so_hd_data = ten_data = None
    if "DATA" in wb.sheetnames:                                     # sheet DATA: thông tin đợt / HĐ / nhà thầu
        for r in wb["DATA"].iter_rows(min_row=1, max_row=40, max_col=4, values_only=True):
            t = na(r[0])
            if t.startswith("DOT THANH TOAN SO"):
                try: dot_data = int(str(r[1]).strip())
                except (TypeError, ValueError): pass
            elif t.startswith("DEN NGAY"): ngay_data = _ngay(r[1])
            elif t.startswith("SO HOP DONG") and r[1]: so_hd_data = str(r[1]).strip()
    sc = next((x for x in wb.sheetnames if "COVER" in na(x)), None)
    if sc:
        for r in wb[sc].iter_rows(min_row=1, max_row=40, max_col=8, values_only=True):
            if na(r[0]).startswith("GIA TRI THUC HIEN LUY KE"): lk_cover = next((num(x) for x in r[1:] if isinstance(x, (int, float))), None)
    wb.close()
    dot = cv.get("dot") or dot_data
    if not cv.get("so_hd") and so_hd_data: cv["so_hd"] = so_hd_data
    if not cv.get("dot") and dot_data: cv["dot"] = dot_data
    if not cv.get("ngay") and ngay_data: cv["ngay"] = ngay_data
    gom, thu_tu, nhom_dot, ngoai = {}, [], None, False
    for i, r in enumerate(rows[h + 1:], h + 2):
        a, ds = na(r[c_stt]), r[c_ds]; dsn = na(ds)
        if dsn.startswith("DE NGHI THANH TOAN") or a == "(DNTT)": break
        m = re.match(r"DOT\s*(\d+)", dsn)
        if a and not isinstance(r[c_dg], (int, float)) and m: nhom_dot = int(m.group(1)); continue          # nhóm 'Đợt N từ ngày…'
        if a and dsn and not isinstance(r[c_dg], (int, float)):
            if any(k in dsn for k in ("NGOAI HOP DONG", "PHAT SINH", "PHU LUC")): ngoai = True
            continue
        if not (ds and isinstance(r[c_dg], (int, float)) and r[c_dg] > 0 and isinstance(r[c_ng], (dt.date, dt.datetime))): continue
        kl, tt = num(r[c_kl]), num(r[c_tt]) if c_tt is not None else num(r[c_kl]) * num(r[c_dg])
        k = (na(ds), round(num(r[c_dg]), 2), ngoai)
        if k not in gom:
            gom[k] = dict(dong=i, khung="Công nhật", stt=str(len(gom) + 1), ds=str(ds).strip(), dvt=str(r[c_dv] or "công").strip(), kl_hd=None, dg=num(r[c_dg]),
                          kl_kt=0.0, kl_kn=0.0, kl_lk=0.0, tt_kt=0.0, tt_kn=0.0, tt_lk=0.0, ngoai=ngoai); thu_tu.append(k)
        g = gom[k]; ky = "kn" if (nhom_dot is None or dot is None or nhom_dot == dot) else "kt"
        g[f"kl_{ky}"] += kl; g[f"tt_{ky}"] += tt; g["kl_lk"] += kl; g["tt_lk"] += tt
    lines = [gom[k] for k in thu_tu]
    tong = (sum(l["tt_kt"] for l in lines), sum(l["tt_kn"] for l in lines), sum(l["tt_lk"] for l in lines))
    return dict(file=path, sheet=sn, cover=cv, lines=lines, tong=tong, vat=None, tu=0.0, hu=0.0, tu_k=0.0, hu_k=0.0, gl=None, du_tru=0,
                du_an_text=" ".join(txt), mau="CONG_NHAT", lk_cover=lk_cover)


HU_MA_NS = [   # HOÀN ỨNG BCH: diễn giải dòng chi tiết → mã NS Prelims (đối chiếu sheet 'C. ChiPhiGianTiep' của NS R00 · anh chốt 26/09). Thứ tự = ưu tiên.
    ("Prelim_3", "An toàn lao động", r"PHAN QUANG|BAO HO|AN TOAN|HSAT|THE AT|GKSK|NON BAO|DAY DAI"),
    ("Prelim_9.2", "Chi phí công nhật cơ hữu", r"CONG NHAT"),
    ("Prelim_9.3", "Chi phí khởi công dự án", r"KHOI CONG"),
    ("Prelim_6", "Chi phí ngoại giao", r"NGOAI GIAO|TIEP KHACH"),
    ("Prelim_2", "Hệ thống điện, nước tạm", r"TIEN DIEN|TIEN NUOC|DIEN NUOC"),
    (None, "Dụng cụ thi công — CHỜ ANH CHỌN MÃ NS", r"THUOC|MANG HO|CO RUA|CO LE|ONG DIEU|XENG|NHO XAY|MUI KHOAN|MUI DUC|CHOI DOT NHUA|VE SINH SAN|BUA |KIM HAN"),
    ("Prelim_5", "Chi phí sinh hoạt BCH", r"CHOI|XUC RAC|THUNG SON|NUOC UONG|SINH HOAT|NHA TRO|THUE NHA"),
    ("Prelim_1", "Tiện ích văn phòng tạm + kho bãi", r"VAN PHONG|GHE|QUAT|WIFI|INTERNET|CUA CHINH|CUA SO|SIMILI|TAM OP|PHOTO|KHO|PALLET|BAI|HOC VAT TU|CHI MAY|^BAO$|MAY LANH"),
]
HU_NHOM_MAC_DINH = {1: "Prelim_5", 3: "Prelim_1", 4: "Prelim_6", 5: "Prelim_6"}   # không khớp từ khoá ⇒ theo nhóm La Mã của mẫu BCH (II 'phục vụ thi công' ⇒ chờ chọn)
def ma_ns_hoan_ung(dien_giai, nhom_la_ma):
    t = na(dien_giai).strip()
    for ma, ten, rx in HU_MA_NS:
        if re.search(rx, t): return ma, ten
    ma = HU_NHOM_MAC_DINH.get(nhom_la_ma)
    return (ma, next(t_ for m_, t_, _r in HU_MA_NS if m_ == ma)) if ma else (None, "CHƯA GÁN MÃ NS — chờ anh chọn")

LA_MA = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
def la_hoan_ung(wb):
    t = [na(s) for s in wb.sheetnames]
    return any("BANG KE" in s for s in t) and any("TO TRINH" in s or "H.U" in s or "PHIEU T.U" in s for s in t)

def _so(v):
    if isinstance(v, (int, float)): return float(v)
    s = re.sub(r"[^\d]", "", str(v or "")); return float(s) if s else None

def doc_hoan_ung(path, wb):
    """HOÀN ỨNG QUỸ BCH: 'Bảng kê' chia A (có hoá đơn) / B (không hoá đơn) × nhóm chi phí La Mã (I…V) × dòng chi tiết (ĐVT·KL·ĐG·thành tiền·ngày).
    Mỗi NHÓM CHI PHÍ = 1 dòng HSTT (gộp A+B), tính theo tiền (ĐVT 'đ', ĐG 1). File chỉ có ĐỢT NÀY ⇒ kỳ trước lấy từ khung lúc kiểm.
    Tự kiểm: KL×ĐG từng dòng · Σ chi tiết = dòng nhóm · Σ nhóm = dòng phần A/B · tổng = Tờ trình (từng nhóm) = Phiếu tạm ứng = Giấy thanh toán tạm ứng."""
    kr = []; add = lambda *x: kr.append(x)
    sk = next(s for s in wb.sheetnames if "BANG KE" in na(s))
    rows = [list(r) + [None] * 12 for r in wb[sk].iter_rows(min_row=1, max_row=3000, max_col=12, values_only=True)]
    h = next(i for i, r in enumerate(rows) if any(na(c) == "DIEN GIAI" for c in r))
    H = [na(c) for c in rows[h]]; col = lambda lab: next(j for j, c in enumerate(H) if c.startswith(lab))
    cS, cD, cK, cG, cT = col("STT"), col("DIEN GIAI"), col("KHOI LUONG"), col("DON GIA"), col("THANH TIEN")
    cC = next((j for j, c in enumerate(H) if c.startswith("GHI CHU")), None)
    dau = " ".join(na(c) for r in rows[:h] for c in r if c)
    phan = nhom = None; gom, ten_nhom, dong_nhom, t_phan, t_nhom, ct = {}, {}, {}, {}, {}, {}
    net_a, gross_a, gross_b, ts_khac = {}, {}, {}, []                    # A = có hoá đơn ⇒ TÁCH VAT (anh chốt 26/09: chi phí nhập vào khung là TRƯỚC VAT)
    theo_ns, ten_ns, chua_gan = {}, {}, []                                # (mã NS, 'H'/'') → [trước VAT, đã trả] · gán THEO TỪNG DÒNG chi tiết
    for i, r in enumerate(rows[h + 1:], h + 2):
        s = na(r[cS]).strip(); ds = str(r[cD] or "").strip(); g = r[cT]
        if s in ("A", "B") and ds: phan = s; t_phan[s] = (i, _so(g) or 0.0); continue
        if s in LA_MA and ds and phan:
            nhom = LA_MA[s]; t_nhom[(phan, nhom)] = (i, _so(g) or 0.0); ten_nhom.setdefault(nhom, ds); dong_nhom.setdefault(nhom, i); gom.setdefault(nhom, 0.0); continue
        if re.fullmatch(r"[IVX]+\.\d+", s): continue                          # nhóm con (II.1 THIẾT BỊ …) — chỉ là dòng cộng
        if not (phan and nhom and re.fullmatch(r"\d+", s) and isinstance(g, (int, float)) and g): continue
        e, f = r[cK], r[cG]
        if isinstance(e, (int, float)) and isinstance(f, (int, float)) and abs(e * f - g) > 1:
            add("CHAN", "Số học", f"{sk}!dòng {i}", round(g), round(e * f), f"Thành tiền ≠ KL × ĐG: {ds[:40]}")
        gom[nhom] += g; ct[(phan, nhom)] = ct.get((phan, nhom), 0.0) + g
        ts = 0.0
        if phan == "A":                                                   # thuế suất: ghi chú dòng ghi 'x%' / 'KCT' ⇒ theo đó, không ghi ⇒ mặc định 8%
            gc = na(r[cC]) if cC is not None else ""; mv = re.search(r"(\d+(?:[.,]\d+)?)\s*%", gc)
            ts = 0.0 if "KCT" in gc or "KHONG CHIU THUE" in gc else (float(mv.group(1).replace(",", ".")) / 100 if mv else 0.08)
            if ts != 0.08: ts_khac.append((i, ts))
            net_a[nhom] = net_a.get(nhom, 0.0) + g / (1 + ts); gross_a[nhom] = gross_a.get(nhom, 0.0) + g
        else: gross_b[nhom] = gross_b.get(nhom, 0.0) + g
        ma_ns, tn_ = ma_ns_hoan_ung(ds, nhom); key = (ma_ns or "CHUA_GAN", "H" if phan == "A" else ""); ten_ns[key[0]] = tn_
        v_ = theo_ns.setdefault(key, [0.0, 0.0]); v_[0] += g / (1 + ts); v_[1] += g
        if not ma_ns: chua_gan.append((i, ds, g))
    for (p, n_), (i, v) in t_nhom.items():
        if abs(ct.get((p, n_), 0.0) - v) > 1: add("CHAN", "Số học", f"{sk}!dòng {i}", round(ct.get((p, n_), 0.0)), round(v), f"Σ chi tiết ≠ dòng nhóm {ten_nhom[n_][:35]} (phần {p})")
    for p, (i, v) in t_phan.items():
        sv = sum(v_ for (p_, _n), (_i, v_) in t_nhom.items() if p_ == p)
        if abs(sv - v) > 1: add("CHAN", "Số học", f"{sk}!dòng {i}", round(sv), round(v), f"Σ các nhóm ≠ dòng tổng phần {p}")
    tong = sum(gom.values())
    # Tờ trình: bảng I (từng nhóm: có HĐ + không HĐ = tổng) · tên dự án
    st = next((s for s in wb.sheetnames if "TO TRINH" in na(s)), None); txt = [dau]; so_tt = None
    if st:
        tr = [list(r) + [None] * 8 for r in wb[st].iter_rows(min_row=1, max_row=120, max_col=8, values_only=True)]
        so_tt = next((str(r[0]).split(":", 1)[1].strip() for r in tr if na(r[0]).startswith("SO:")), None)
        ct_da = re.search(r"CONG TRINH:\s*(.+?)(KY THU|$)", dau)
        for i, r in enumerate(tr, 1):
            a = na(r[0])
            if a.startswith("TONG GIA TRI") and isinstance(r[5], (int, float)):
                if abs(r[5] - tong) > 1: add("CHAN", "Số học", f"{st}!dòng {i}", round(r[5]), round(tong), "Tổng Tờ trình ≠ Σ Bảng kê")
                break
            n_ = next((n for n, t in ten_nhom.items() if na(t) == na(r[2])), None)
            if n_ and isinstance(r[5], (int, float)) and abs(r[5] - gom[n_]) > 1:
                add("CHAN", "Số học", f"{st}!dòng {i}", round(r[5]), round(gom[n_]), f"Tờ trình nhóm {ten_nhom[n_][:35]} ≠ Bảng kê")
        for i, r in enumerate(tr, 1):
            a = na(r[0]); txt.append(a)
            if a.startswith("BCH DU AN") and ct_da and not all(w in a for w in ct_da.group(1).split()[:3]):
                add("LUU_Y", "Hồ sơ", f"{st}!dòng {i}", str(r[0])[:45], ct_da.group(1).strip()[:30], "Tờ trình ghi tên dự án KHÁC (sót từ mẫu cũ?) — anh soát lại trước khi trình ký")
    # Phiếu đề nghị tạm ứng · Giấy thanh toán tạm ứng
    ngay = None; chua_hu = None
    for s in wb.sheetnames:
        if "PHIEU T.U" in na(s) or "H.U" in na(s):
            for i, r in enumerate(wb[s].iter_rows(min_row=1, max_row=60, max_col=9, values_only=True), 1):
                a = na(r[0]); v = next((_so(c) for c in r[1:] if _so(c)), None) if r else None
                m = re.search(r"NGAY:\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", a)
                if m and ngay is None:
                    try: ngay = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                    except ValueError: pass
                if (a.startswith("SO TIEN DE NGHI") or "SO TIEN DA CHI" in a) and v is not None and abs(v - tong) > 1:
                    add("CHAN", "Số học", f"{s}!dòng {i}", round(v), round(tong), f"{str(r[0]).strip()[:35]} ≠ Σ Bảng kê")
                if "SO TIEN CHUA HOAN UNG" in a: chua_hu = v
    if chua_hu: add("LUU_Y", "Hồ sơ", "Giấy thanh toán tạm ứng", round(chua_hu), "—", "Còn tạm ứng CHƯA hoàn từ đợt trước — kiểm chênh lệch trước khi chi")
    if ngay is None:
        m = re.match(r"(\d{2})(\d{2})(\d{2})", os.path.basename(path).split("_", 1)[-1])
        try: ngay = dt.date(2000 + int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None
        except ValueError: ngay = None
    md = re.search(r"KY THU\s*(\d+)", dau) or re.search(r"DOT\s*(\d+)", " ".join(txt))
    cv = dict(ten_don_vi="Ban chỉ huy công trường", so_hd=None, dot=int(md.group(1)) if md else None, ngay=ngay, so_to_trinh=so_tt,
              o={"dot": f"{sk}!đầu trang", "ten_don_vi": "Tờ trình", "so_hd": "Tờ trình"})
    lines = []                                                            # 1 dòng / MÃ NS × (không HĐ | 'H' có HĐ: TRƯỚC VAT, tiền chi = đã gồm VAT)
    for (ma, h_), (net, gross) in sorted(theo_ns.items()):
        nh = None if ma == "CHUA_GAN" else f"HU_{ma}"
        lines.append(dict(dong=h + 2, khung="", stt=ma + h_, ds=f"Hoàn ứng BCH — {ten_ns[ma]}" + (" — có hoá đơn (trước VAT)" if h_ else ""), dvt="đ", kl_hd=None, dg=1.0,
                          kl_kt=0.0, kl_kn=net, kl_lk=net, tt_kt=0.0, tt_kn=net, tt_lk=net, ngoai=False, nhom=nh,
                          nhom_moi=(nh, f"Hoàn ứng BCH → {ten_ns[ma]}", "đ", ma) if nh else None, **({"vat_rieng": gross / net - 1} if h_ and net else {})))
    if chua_gan:
        add("LUU_Y", "Theo HĐ", f"{sk}!dòng {chua_gan[0][0]}…", round(sum(x[2] for x in chua_gan)), len(chua_gan),
            "Dòng CHƯA GÁN MÃ NS (dụng cụ thi công…): " + ", ".join(x[1][:18] for x in chua_gan[:8]) + (" …" if len(chua_gan) > 8 else "") + " — anh chọn mã NS để em thêm luật")
    if sum(gross_a.values()) > 1:
        add("LUU_Y", "Theo HĐ", f"{sk} phần A", round(sum(gross_a.values())), round(sum(net_a.values())),
            "Chi phí CÓ hoá đơn đã TÁCH VAT (mặc định 8%" + (f"; {len(ts_khac)} dòng theo thuế suất ghi chú" if ts_khac else "") + ") — ghi chi phí trước VAT, tiền chi = đã gồm VAT")
    tn = sum(l["tt_kn"] for l in lines)
    return dict(file=path, sheet=sk, cover=cv, lines=lines, tong=(0.0, tn, tn), vat=0, tu=0.0, hu=0.0, tu_k=0.0, hu_k=0.0, gl=None, du_tru=0,
                du_an_text=" ".join(txt), mau="HOAN_UNG_BCH", kiem_rieng=kr)
