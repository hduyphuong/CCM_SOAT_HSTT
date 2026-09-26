"""PHÍA CĐT (doanh thu): đọc hồ sơ thanh toán gửi CĐT · kiểm với BOQ HĐ CĐT (N5) + đợt đã ghi (N8) · kế hoạch ghi sổ.
Khấu trừ / phạt / cấp vật tư của CĐT = CHI của mình ⇒ ghi bên ĐỐI TÁC (N9, HĐ khấu trừ CĐT) — đúng bản chất THU–CHI."""
import openpyxl, re, datetime as dt
from collections import defaultdict
from doc_hstt import na, num, _ngay
NGUONG = 10
MA_KT = "HD-CDT-KT"                 # HĐ "CĐT khấu trừ" bên đối tác (N6) — dòng 1 = tiện ích, 2 = phạt, 3 = cấp vật tư
MA_PHAT = "HD-CDT-PHAT"             # HĐ "CĐT phạt" bên đối tác (N6) — VAT 0% (phạt không có hoá đơn VAT) · anh chốt 26/09
MA_VT = "HD-CDT-VT"                 # HĐ "CĐT cấp vật tư" bên đối tác (N6) — mỗi loại vật tư 1 dòng; CĐT KHÔNG trừ trên phiếu mà tự cắt tiền khi chuyển
VT_NHOM = [("VT_CDT_GACH", "CĐT cấp gạch xây — CĐT cắt tiền", "Viên", "NCC_Gach", "GACH"),          # (nhóm, tên, ĐVT, mã NS, từ khoá tên vật tư)
           ("VT_CDT_XIMANG", "CĐT cấp xi măng — CĐT cắt tiền", "Bao", "NCC_VLXD", "XI MANG")]
def nhom_vt(ten): return next((n[0] for n in VT_NHOM if n[4] in na(ten)), "")

def _doc_vat_tu(wb):
    """Sheet 'TH VATTU' (bảng tổng hợp vật tư CĐT cấp): mỗi dòng KL kỳ trước / đợt này / lũy kế + ĐG + thành tiền; dòng TỔNG CỘNG TRƯỚC THUẾ."""
    sn = next((s for s in wb.sheetnames if "VATTU" in na(s).replace(" ", "") and "HAO" not in na(s)), None)
    if not sn: return None
    rows = [list(r) for r in wb[sn].iter_rows(min_row=1, max_row=80, max_col=16, values_only=True)]
    h = next((i for i, r in enumerate(rows) if any(na(c).startswith("NOI DUNG KHAU TRU") for c in r)), None)
    if h is None: return None
    H, S = [na(c) for c in rows[h]], [na(c) for c in rows[h + 1]]
    c_stt, c_ten = next(j for j, c in enumerate(H) if c == "STT"), next(j for j, c in enumerate(H) if c.startswith("NOI DUNG"))
    c_dvt, c_dg = next(j for j, c in enumerate(H) if c == "DVT"), next(j for j, c in enumerate(H) if c.startswith("DON GIA"))
    kt = [j for j, c in enumerate(S) if c.startswith("LUY KE DEN HET KY TRUOC")]; kn = [j for j, c in enumerate(S) if c == "DOT NAY"]
    lk = [j for j, c in enumerate(S) if c.startswith("LUY KE DEN HET DOT NAY")]
    if len(kt) < 2 or len(kn) < 2 or len(lk) < 2: raise ValueError(f"{sn}: không nhận ra cột kỳ trước / đợt này / lũy kế của bảng vật tư CĐT cấp")
    out = dict(sheet=sn, items=[], tong=None)
    for i, r in enumerate(rows[h + 2:], h + 3):
        if na(r[c_ten]).startswith("TONG CONG TRUOC THUE"): out["tong"] = (num(r[kt[1]]), num(r[kn[1]]), num(r[lk[1]])); continue
        if re.fullmatch(r"\d+", na(r[c_stt]).replace(".0", "")) and r[c_ten]:
            x = dict(dong=i, stt=na(r[c_stt]).replace(".0", ""), ten=str(r[c_ten]).strip(), dvt=str(r[c_dvt] or "").strip(), dg=num(r[c_dg]),
                     kl_kt=num(r[kt[0]]), kl_kn=num(r[kn[0]]), tt_kt=num(r[kt[1]]), tt_kn=num(r[kn[1]]), tt_lk=num(r[lk[1]]))
            x["kl_lk"] = x["kl_kt"] + x["kl_kn"] if r[lk[0]] is None else num(r[lk[0]]); out["items"].append(x)
    return out

def _doc_phat(wb):
    """Sheet 'TH PHẠT': dòng TỔNG CỘNG GIÁ TRỊ PHẠT ⇒ (lũy kế đợt trước, đợt hiện tại, lũy kế). Phạt KHÔNG cấn trừ trên phiếu — chuyển thẳng."""
    sn = next((s for s in wb.sheetnames if "PHAT" in na(s).split() or na(s).endswith("PHAT")), None)
    if not sn: return None
    rows = [list(r) for r in wb[sn].iter_rows(min_row=1, max_row=80, max_col=12, values_only=True)]
    h = next((i for i, r in enumerate(rows) if any(na(c) == "DOT HIEN TAI" for c in r)), None)
    if h is None: return None
    S = [na(c) for c in rows[h]]; c_kt = next(j for j, c in enumerate(S) if c.startswith("LUY KE DOT TRUOC")); c_kn = S.index("DOT HIEN TAI"); c_lk = S.index("LUY KE")
    for i, r in enumerate(rows[h + 1:], h + 2):
        if any(na(c).startswith("TONG CONG GIA TRI PHAT") for c in r): return dict(sheet=sn, dong=i, tong=(num(r[c_kt]), num(r[c_kn]), num(r[c_lk])))
    return None

def la_claim_cdt(wb): return any("BANG GIA TRI KL" in na(s) for s in wb.sheetnames)

def doc_claim(path, wb):
    sn = next(s for s in wb.sheetnames if "BANG GIA TRI KL" in na(s))
    rows = [list(r) for r in wb[sn].iter_rows(min_row=1, max_row=200, max_col=16, values_only=True)]
    h = next(i for i, r in enumerate(rows) if any(na(c) == "NOI DUNG CONG VIEC" for c in r))
    H, S = [na(c) for c in rows[h]], [na(c) for c in rows[h + 1]]
    col = lambda lab, arr: next((j for j, c in enumerate(arr) if c.startswith(lab)), None)
    c_stt, c_ds, c_dv, c_klhd, c_dg = col("STT", H), col("NOI DUNG", H), col("DON VI", H), col("KHOI LUONG HD", H), col("DON GIA", H)
    kt = [j for j, c in enumerate(S) if c in ("KI TRUOC", "KY TRUOC")]; kn = [j for j, c in enumerate(S) if c in ("KI NAY", "KY NAY")]
    lk = [j for j, c in enumerate(S) if c == "LUY KE"]
    txt = " ".join(na(c) for r in rows[:h] for c in r if c)
    so_hd = next((str(c).split(":", 1)[1].strip() for r in rows[:h] for c in r if na(c).startswith("SO HOP DONG:")), None)
    m = re.search(r"DOT\s*(\d+)", txt); dot = int(m.group(1)) if m else None
    kq = dict(loai="CDT", file=path, sheet=sn, lines=[], tong=None, vat=None, tu=0.0, hu=0.0, tu_k=0.0, hu_k=0.0, gl=None, du_tru=0)
    for i, r in enumerate(rows[h + 2:], h + 3):
        a = na(r[c_stt])
        if a == "B": kq["tong"] = (num(r[kt[1]]), num(r[kn[1]]), num(r[lk[1]])); continue
        if a == "C":
            mv = re.search(r"(\d+(\.\d+)?)\s*%", str(r[c_ds] or "")); kq["vat"] = float(mv.group(1)) / 100 if mv else None; break
        if re.fullmatch(r"\d+", a) and r[c_dv] and not str(r[c_dv]).startswith(("B =", "C =")):
            kq["lines"].append(dict(dong=i, khung="", stt=a, ds=str(r[c_ds]).strip(), dvt=str(r[c_dv]).strip(),
                kl_hd=r[c_klhd] if isinstance(r[c_klhd], (int, float)) else None, dg=num(r[c_dg]), kl_kt=num(r[kt[0]]), kl_kn=num(r[kn[0]]),
                kl_lk=num(r[kt[0]]) + num(r[kn[0]]) if r[lk[0]] is None else num(r[lk[0]]), tt_kt=num(r[kt[1]]), tt_kn=num(r[kn[1]]), tt_lk=num(r[lk[1]]), ngoai=False))
    tien = {}
    sp = next((s for s in wb.sheetnames if "PHIEU" in na(s)), None)
    if sp:
        for r in wb[sp].iter_rows(min_row=10, max_row=40, max_col=6, values_only=True):
            t = na(r[0]); v = r[4] if isinstance(r[4], (int, float)) else None
            for k_, lab in (("th_vat", "TONG CONG"), ("thu_hoi", "GIA TRI THU HOI"), ("khau_tru", "GIA TRI KHAU TRU"), ("giu_lai", "GIA TRI GIU LAI"), ("de_nghi", "GIA TRI DE NGHI")):
                if t.startswith(lab) and v is not None: tien[k_] = v
    ngay = None
    mf = re.match(r"(\d{2})(\d{2})(\d{2})", path.split("\\")[-1].split("_", 1)[-1])
    if mf:
        try: ngay = dt.date(2000 + int(mf.group(1)), int(mf.group(2)), int(mf.group(3)))
        except ValueError: ngay = None
    if ngay is None:
        sb = next((s for s in wb.sheetnames if "BIA" in na(s)), None)
        bia = " ".join(na(c) for r in (wb[sb].iter_rows(min_row=1, max_row=80, max_col=12, values_only=True) if sb else []) for c in r if c)
        mm = re.search(r"THANG (\d{1,2}) ?/(\d{4})", bia)
        if mm: ngay = dt.date(int(mm.group(2)), int(mm.group(1)), 25)
    kq["cover"] = dict(ten_don_vi="CĐT — hồ sơ thanh toán gửi CĐT", so_hd=so_hd, dot=dot, ngay=ngay, o={"so_hd": f"{sn}!đầu trang", "dot": f"{sn}!đầu trang"})
    kq["tien"] = tien; kq["du_an_text"] = txt
    kq["vat_tu"] = _doc_vat_tu(wb); kq["phat"] = _doc_phat(wb)                 # 2 khoản CĐT tự cắt tiền, KHÔNG nằm trên phiếu ĐNTT
    return kq

def doc_ben_cdt(wb, nhom_ns):
    """Đọc N4 (HĐ CĐT) · N5 (BOQ CĐT) · N8 (thanh toán CĐT)."""
    b = dict(hd={}, dong={}, lk_kl=defaultdict(float), lk_tien=defaultdict(float), lk_tien_dong=defaultdict(float), dot_cuoi={}, tu_treo=defaultdict(float))
    w = wb["N4_HD_CDT"]
    for r in range(2, w.max_row + 1):
        m = w[f"A{r}"].value
        if not m: continue
        rec = b["hd"].setdefault(m, dict(ma_hd=m, so_hd=str(w[f"D{r}"].value or ""), gia_tri=0.0, vat=num(w[f"I{r}"].value), pct_tu=num(w[f"J{r}"].value), pct_tt_dot=w[f"K{r}"].value))
        rec["gia_tri"] += num(w[f"H{r}"].value)
    w = wb["N5_BOQ_CDT"]
    for r in range(2, w.max_row + 1):
        m = w[f"A{r}"].value
        if m: b["dong"].setdefault(m, []).append(dict(row=r, stt=str(w[f"B{r}"].value), pham_vi=w[f"D{r}"].value, noi_dung=str(w[f"E{r}"].value or ""),
                                                      dvt=str(w[f"F{r}"].value or ""), kl_hd=w[f"G{r}"].value, don_gia=num(w[f"H{r}"].value), nhom=w[f"K{r}"].value, ma_ns=""))
    w = wb["N8_TT_CDT"]
    for r in range(2, w.max_row + 1):
        m = w[f"A{r}"].value
        if not m: continue
        loai, st = w[f"D{r}"].value, str(w[f"E{r}"].value) if w[f"E{r}"].value is not None else ""
        kl, dg, tien = w[f"H{r}"].value, num(w[f"I{r}"].value), num(w[f"J{r}"].value)
        gt = kl * dg if isinstance(kl, (int, float)) else tien
        if isinstance(w[f"B{r}"].value, (int, float)) and w[f"B{r}"].value: b["dot_cuoi"][m] = max(b["dot_cuoi"].get(m, 0), int(w[f"B{r}"].value))
        if loai in ("THUC_HIEN", "DIEU_CHINH"):
            if isinstance(kl, (int, float)): b["lk_kl"][(m, st)] += kl
            b["lk_tien"][m] += gt; b["lk_tien_dong"][(m, st)] += gt
        if loai in ("TAM_UNG", "HOAN_UNG"): b["tu_treo"][m] += tien
    return b

def _khop(l, ds):
    c = [d for d in ds if d["stt"] == l["stt"] and na(d["noi_dung"]) == na(l["ds"])] or [d for d in ds if d["stt"] == l["stt"]]
    return c[0] if c else None

def kiem_cdt(hs, k, van_tay_da_co=()):
    b = k["cdt"]; co = []; add = lambda *x: co.append(x)
    cv, S, t, tien = hs["cover"], hs["sheet"], hs["tong"] or (0, 0, 0), hs["tien"]
    ma = next((m for m, h in b["hd"].items() if cv["so_hd"] and re.sub(r"[^A-Z0-9]", "", na(h["so_hd"])) == re.sub(r"[^A-Z0-9]", "", na(cv["so_hd"]))), None)
    if ma is None and len(b["hd"]) == 1: ma = next(iter(b["hd"]))
    if hs.get("van_tay") in van_tay_da_co: add("CHAN", "Hồ sơ", "file", hs["van_tay"][:12], "đã nạp", "File trùng nội dung với hồ sơ đã nạp trước")
    if not ma: add("CHAN", "Hồ sơ", cv["o"]["so_hd"], cv["so_hd"], "—", "Không tìm thấy HĐ CĐT trong khung (N4)")
    elif cv["so_hd"] and re.sub(r"[^A-Z0-9]", "", na(b["hd"][ma]["so_hd"])) != re.sub(r"[^A-Z0-9]", "", na(cv["so_hd"])):
        add("LUU_Y", "Hồ sơ", cv["o"]["so_hd"], cv["so_hd"], b["hd"][ma]["so_hd"], "Số HĐ trên hồ sơ khác số HĐ CĐT trong khung")
    if cv["ngay"] is None: add("LUU_Y", "Hồ sơ", "BIA", "—", "—", "Không đọc được ngày hồ sơ")
    kn = sum(l["tt_kn"] for l in hs["lines"])
    if abs(kn - t[1]) > NGUONG: add("CHAN", "Số học", f"{S} dòng B", round(kn), round(t[1]), "Σ các dòng kỳ này ≠ TỔNG GIÁ TRỊ TRƯỚC THUẾ kỳ này")
    for l in hs["lines"]:
        for ky in ("kt", "kn", "lk"):                                  # dòng tính theo giá trị (VD 'Chi phí quản lý 2,5%'): KL trống, có tiền ⇒ KL = tiền ÷ ĐG
            if not l[f"kl_{ky}"] and abs(l[f"tt_{ky}"]) > NGUONG and l["dg"]:
                l[f"kl_{ky}"] = l[f"tt_{ky}"] / l["dg"]
                if ky == "kn": add("LUU_Y", "Số học", f"{S}!dòng {l['dong']}", round(l["tt_kn"]), round(l["kl_kn"], 8), f"KL kỳ này suy = tiền ÷ ĐG ({l['kl_kn']:.8f}): {l['ds'][:40]}")
        if abs(l["kl_kn"] * l["dg"] - l["tt_kn"]) > NGUONG: add("CHAN", "Số học", f"{S}!dòng {l['dong']}", round(l["tt_kn"]), round(l["kl_kn"] * l["dg"]), f"Tiền ≠ KL × ĐG: {l['ds'][:40]}")
    vat = hs["vat"] if hs["vat"] is not None else (b["hd"][ma]["vat"] if ma else 0)
    if tien.get("th_vat") is not None and abs(t[1] * (1 + vat) - tien["th_vat"]) > NGUONG + 1:
        add("CHAN", "Số học", "PHIẾU ĐNTT", round(tien["th_vat"]), round(t[1] * (1 + vat)), "Giá trị có VAT trên phiếu ≠ kỳ này × (1 + VAT)")
    tom = dict(ky_nay=t[1], luy_ke=t[2], ky_truoc=t[0], so_dong=len(hs["lines"]), don_vi=cv["ten_don_vi"], so_hd=cv["so_hd"], dot=cv["dot"],
               ngay=cv["ngay"].isoformat() if cv["ngay"] else None, vat=vat, du_tru_tam_ung=None, tien=tien)
    pl = dict(ma_hd=ma, loai_doi_tac="CĐT", loai_hs="DOANH_THU")
    if not ma: return dict(phan_loai=pl, co=_dang(co), tom_tat=tom, khop=[])
    h = b["hd"][ma]; ds = b["dong"].get(ma, [])
    if cv["dot"] is not None and cv["dot"] <= b["dot_cuoi"].get(ma, 0):
        add("CHAN", "Hồ sơ", cv["o"]["dot"], cv["dot"], b["dot_cuoi"][ma], f"Đợt {cv['dot']} đã ghi sổ phía CĐT — chống ghi doanh thu 2 lần")
    elif cv["dot"] is not None and cv["dot"] > b["dot_cuoi"].get(ma, 0) + 1: add("LUU_Y", "Hồ sơ", cv["o"]["dot"], cv["dot"], b["dot_cuoi"].get(ma, 0) + 1, "Thiếu đợt ở giữa")
    khop = []
    for l in hs["lines"]:
        d = _khop(l, ds); khop.append((l, d["stt"] if d else None)); vt = f"{S}!dòng {l['dong']}"
        if d is None:
            if l["kl_lk"]: add("LUU_Y", "Theo HĐ", vt, l["ds"][:40], "—", "Dòng không có trong BOQ HĐ CĐT — phát sinh, cần PLHĐ với CĐT")
            continue
        if d["don_gia"] and abs(l["dg"] - d["don_gia"]) > 0.5: add("LUU_Y", "Theo HĐ", vt, l["dg"], d["don_gia"], f"ĐG khác BOQ HĐ CĐT: {l['ds'][:40]}")
        if isinstance(d["kl_hd"], (int, float)) and d["kl_hd"] >= 0 and d["don_gia"] > 0 and l["kl_lk"] > d["kl_hd"] * 1.0001 + 0.001:
            add("LUU_Y", "Theo HĐ", vt, l["kl_lk"], d["kl_hd"], f"KL nghiệm thu vượt KL HĐ ⇒ phát sinh CĐT chưa PLHĐ: {l['ds'][:40]}")
        lk0 = b["lk_kl"].get((ma, d["stt"]), 0.0)
        if abs(l["kl_kt"] - lk0) > 1e-6: add("LUU_Y", "Đợt trước", vt, round(l["kl_kt"], 4), round(lk0, 4), f"KL kỳ trước ≠ lũy kế đã ghi ⇒ sẽ ghi ĐIỀU CHỈNH: {l['ds'][:40]}")
    if abs(t[0] - b["lk_tien"][ma]) > NGUONG: add("LUU_Y", "Đợt trước", f"{S} dòng B", round(t[0]), round(b["lk_tien"][ma]), "Giá trị kỳ trước ≠ lũy kế doanh thu đã ghi sổ")
    th = tien.get("th_vat", t[1] * (1 + vat)); ptt = num(h["pct_tt_dot"])
    if tien.get("giu_lai") is not None and ptt and abs(tien["giu_lai"] - th * (1 - ptt)) > NGUONG: add("LUU_Y", "Theo HĐ", "PHIẾU ĐNTT", round(tien["giu_lai"]), round(th * (1 - ptt)), "Giữ lại khác % HĐ")
    tu_con = b["tu_treo"][ma]
    if tien.get("thu_hoi", 0) > tu_con + NGUONG: add("LUU_Y", "Theo HĐ", "PHIẾU ĐNTT", round(tien["thu_hoi"]), round(tu_con), "Thu hồi tạm ứng VƯỢT tạm ứng còn lại")
    if tien.get("de_nghi") is not None and ptt:
        tinh = th * ptt - tien.get("thu_hoi", 0) - tien.get("khau_tru", 0)
        if abs(tien["de_nghi"] - tinh) > NGUONG: add("CHAN", "Số học", "PHIẾU ĐNTT", round(tien["de_nghi"]), round(tinh), "Đề nghị TT ≠ có VAT × %TT − thu hồi − khấu trừ")
    if tien.get("khau_tru") and MA_KT not in k["hd"]: add("LUU_Y", "Hồ sơ", "PHIẾU ĐNTT", round(tien["khau_tru"]), "—", f"Khung chưa có HĐ khấu trừ {MA_KT} — bấm Đồng ý thì app TỰ TẠO rồi ghi khấu trừ bên CHI (÷ 1+VAT)")
    vt_ = hs.get("vat_tu")                                             # VẬT TƯ CĐT CẤP ⇒ CHI (HD-CDT-VT) · kỳ trước từng loại = lũy kế đã ghi sổ
    if vt_:
        sv, its, tg = vt_["sheet"], vt_["items"], vt_["tong"]; tom["vat_tu_cdt"] = tg[1] if tg else None
        if tg and abs(sum(x["tt_kn"] for x in its) - tg[1]) > NGUONG: add("CHAN", "Số học", sv, round(sum(x["tt_kn"] for x in its)), round(tg[1]), "Σ vật tư CĐT cấp đợt này ≠ TỔNG CỘNG TRƯỚC THUẾ")
        for x in its:
            if abs(x["kl_kn"] * x["dg"] - x["tt_kn"]) > NGUONG: add("CHAN", "Số học", f"{sv}!dòng {x['dong']}", round(x["tt_kn"]), round(x["kl_kn"] * x["dg"]), f"Vật tư CĐT cấp: tiền ≠ KL × ĐG: {x['ten'][:40]}")
            if MA_VT in k["hd"] and abs(x["kl_kt"] - k["lk_kl"].get((MA_VT, x["stt"]), 0.0)) > 1e-6:
                add("LUU_Y", "Đợt trước", f"{sv}!dòng {x['dong']}", round(x["kl_kt"], 4), round(k["lk_kl"].get((MA_VT, x["stt"]), 0.0), 4), f"Vật tư CĐT cấp: KL kỳ trước ≠ lũy kế đã ghi sổ: {x['ten'][:40]}")
            if abs(x["kl_kn"]) > 1e-9 and not nhom_vt(x["ten"]): add("LUU_Y", "Theo HĐ", f"{sv}!dòng {x['dong']}", x["ten"][:30], "—", "Vật tư CĐT cấp chưa có nhóm / mã NS (chỉ tự nhận gạch, xi măng) — sẽ ghi 'CHƯA GÁN NHÓM'")
        if tg and abs(tg[1]) > NGUONG and MA_VT not in k["hd"]:
            add("LUU_Y", "Hồ sơ", sv, round(tg[1]), "—", f"Khung chưa có HĐ vật tư CĐT cấp {MA_VT} — bấm Đồng ý thì app TỰ TẠO rồi ghi CHI {tg[1]:,.0f} trước VAT (CĐT tự cắt tiền)")
    ph_ = hs.get("phat")
    if ph_ and abs(ph_["tong"][1]) > NGUONG:
        tom["phat_cdt"] = ph_["tong"][1]
        if MA_PHAT not in k["hd"]:
            add("LUU_Y", "Hồ sơ", f"{ph_['sheet']}!dòng {ph_['dong']}", round(ph_["tong"][1]), "—", f"Khung chưa có HĐ phạt CĐT {MA_PHAT} — bấm Đồng ý thì app TỰ TẠO (VAT 0%) rồi ghi CHI {ph_['tong'][1]:,.0f}")
    if ph_ and MA_PHAT in k["hd"] and abs(ph_["tong"][0] - k["lk_tien"][MA_PHAT]) > NGUONG:     # kỳ trước phạt = lũy kế phạt đã ghi sổ
        add("LUU_Y", "Đợt trước", f"{ph_['sheet']}!dòng {ph_['dong']}", round(ph_["tong"][0]), round(k["lk_tien"][MA_PHAT]), "Phạt CĐT: lũy kế đợt trước ≠ lũy kế phạt đã ghi sổ")
    return dict(phan_loai=pl, co=_dang(co), tom_tat=tom, khop=khop)

def phan_phat(hs, k):
    """Phần ghi sổ PHẠT CĐT: 1 dòng N9 THUC_HIEN (dòng HĐ 1) = phạt đợt này, VAT 0% ⇒ chi phí = tiền chi. None nếu không có."""
    ph_, cv = hs.get("phat"), hs["cover"]
    if not ph_ or MA_PHAT not in k["hd"] or abs(ph_["tong"][1]) <= NGUONG: return None
    return dict(tt="N9_TT_DoiTac", ct="N7_HD_DoiTac_ChiTiet", ma_hd=MA_PHAT, dong_moi=[], lk_hstt=ph_["tong"][2],
                dong_tt=[dict(loai="THUC_HIEN", stt="1", kl=1, dg=ph_["tong"][1], so_tien=None, ma_hd=MA_PHAT, dot=cv["dot"], ngay=cv["ngay"],
                              ghi="CĐT phạt — CHI (VAT 0%, chuyển thẳng, không trừ trên phiếu)")])

def phan_vat_tu(hs, k):
    """Phần ghi sổ VẬT TƯ CĐT CẤP (N7 dòng mới cho loại vật tư chưa có + N9 THUC_HIEN theo KL đợt này × ĐG). None nếu không có."""
    vt_, cv = hs.get("vat_tu"), hs["cover"]
    if not vt_ or MA_VT not in k["hd"] or not any(abs(x["kl_kn"]) > 1e-9 for x in vt_["items"]): return None
    co_ = {str(d["stt"]) for d in k["dong"].get(MA_VT, [])}
    moi = [dict(stt=x["stt"], noi_dung=x["ten"], dvt=x["dvt"], don_gia=x["dg"], pham_vi="TRONG_HD", nhom=nhom_vt(x["ten"]),
                nguon=f"webapp · {vt_['sheet']} đợt {cv['dot']}") for x in vt_["items"] if x["stt"] not in co_]
    return dict(tt="N9_TT_DoiTac", ct="N7_HD_DoiTac_ChiTiet", ma_hd=MA_VT, dong_moi=moi, lk_hstt=vt_["tong"][2] if vt_["tong"] else None,
                dong_tt=[dict(loai="THUC_HIEN", stt=x["stt"], kl=x["kl_kn"], dg=x["dg"], so_tien=None, ma_hd=MA_VT, dot=cv["dot"], ngay=cv["ngay"],
                              ghi="CĐT cấp vật tư — CHI (trước VAT, CĐT tự cắt tiền)") for x in vt_["items"] if abs(x["kl_kn"]) > 1e-9])

def ke_hoach_cdt(hs, kq, k):
    b = k["cdt"]; ma = kq["phan_loai"]["ma_hd"]; cv = hs["cover"]; tt, gop = [], {}
    for l, st in kq["khop"]:
        if st is None: continue                                   # dòng lạ: chưa tự thêm vào BOQ CĐT — anh bổ sung PLHĐ trước
        g = gop.setdefault(st, dict(kt=0.0, vt=0.0)); g["kt"] += l["kl_kt"]; g["vt"] += l["tt_lk"] - l["tt_kn"]
    for st, g in gop.items():
        dkl = g["kt"] - b["lk_kl"].get((ma, st), 0.0); dv = g["vt"] - b["lk_tien_dong"].get((ma, st), 0.0)
        if abs(dkl) > 1e-6: tt.append(dict(loai="DIEU_CHINH", stt=st, kl=dkl, dg=round(dv / dkl, 6), so_tien=None, ghi="sửa kỳ trước"))
        elif abs(dv) > 1: tt.append(dict(loai="DIEU_CHINH", stt=st, kl=None, dg=None, so_tien=round(dv, 2), ghi="đổi giá kỳ trước"))
    for l, st in kq["khop"]:
        if st and abs(l["kl_kn"]) > 1e-12: tt.append(dict(loai="THUC_HIEN", stt=st, kl=l["kl_kn"], dg=l["dg"], so_tien=None, ghi=""))
    tien = hs["tien"]
    if tien.get("thu_hoi"): tt.append(dict(loai="HOAN_UNG", stt=None, kl=None, dg=None, so_tien=-tien["thu_hoi"], ghi="Thu hồi tạm ứng"))
    for x in tt: x.update(ma_hd=ma, dot=cv["dot"], ngay=cv["ngay"])
    phan = [dict(tt="N8_TT_CDT", ct="N5_BOQ_CDT", ma_hd=ma, dong_moi=[], dong_tt=tt, lk_hstt=(hs["tong"] or (0, 0, 0))[2])]
    if tien.get("khau_tru") and MA_KT in k["hd"]:
        vat_kt = k["hd"][MA_KT]["vat"]
        phan.append(dict(tt="N9_TT_DoiTac", ct="N7_HD_DoiTac_ChiTiet", ma_hd=MA_KT, dong_moi=[], lk_hstt=None,
                         dong_tt=[dict(loai="THUC_HIEN", stt="1", kl=1, dg=round(tien["khau_tru"] / (1 + vat_kt), 2), so_tien=None, ma_hd=MA_KT,
                                       dot=cv["dot"], ngay=cv["ngay"], ghi="CĐT khấu trừ — CHI (trước VAT)")]))
    for pv in (phan_vat_tu(hs, k), phan_phat(hs, k)):
        if pv: phan.append(pv)
    return dict(ma_hd=ma, phan=phan, dong_hd_moi=[], dong_tt=tt + [x for p in phan[1:] for x in p["dong_tt"]], lk_hstt=phan[0]["lk_hstt"])

def _dang(co): return [dict(muc=m, lop=lop, vi_tri=vt, hstt=a, doi_chieu=b_, mo_ta=mt) for m, lop, vt, a, b_, mt in co]
