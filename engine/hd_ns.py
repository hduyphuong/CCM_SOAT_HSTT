"""KIỂM HỢP ĐỒNG ↔ NGÂN SÁCH — nghiệp vụ ĐỘC LẬP (anh Phương 29/09/2026): CHỈ ĐỌC khung, không ghi gì, không nối vào luồng HSTT.
Nguồn HĐ: (1) HĐ đã có trong khung (N6/N7) · (2) HĐ / báo giá AI đã đọc ở Hồ sơ nền nhưng CHƯA nhập khung (kiểm trước khi ký).
Mỗi dòng HĐ → mã NS (khung, hoặc gợi ý theo luật gan_nhom_doi) → dòng NS khớp (cùng mã NS + ĐVT, giống nội dung nhất) ⇒ so ĐƠN GIÁ.
Mỗi mã NS · ĐVT ⇒ so KHỐI LƯỢNG: KL NS ↔ KL HĐ này ↔ KL các HĐ đã ký khác. Mỗi mã NS ⇒ so GIÁ TRỊ.
Mức cờ: DO = đơn giá HĐ vượt NS (so đúng dòng) · VANG = KL/giá trị vượt NS, hoặc ĐG vượt khi chỉ so được bình quân mã
        · XAM = không so được (chưa có mã NS / ĐVT không có trong NS) · XANH = đạt.
Chạy thử: python hd_ns.py <khung.xlsx> <mã HĐ>"""
import math, re, sys, unicodedata
from collections import defaultdict
import openpyxl
import gan_nhom_doi as GN
import nen as NEN

NGUONG_DG = 0.005          # ĐG HĐ vượt NS quá 0,5% mới cờ (né làm tròn)
LECH_BQ = 1.2              # mã NS có ĐG các dòng chênh nhau > 20% ⇒ ĐG bình quân chỉ để tham khảo
KHOP_TOI_THIEU = 0.34      # độ giống nội dung tối thiểu để nhận "đúng dòng NS"
SAN_SAN = 0.15             # dòng NS có điểm cách dòng khớp ≤ 0,15 = "gần giống" ⇒ HĐ vượt ĐG dòng gần giống RẺ hơn thì cờ vàng (chống lọt)
DOI_NGHIA = (("trong", "ngoài"),)
DONG_NGHIA = {"xmcl": ("không", "nung"), "demi": ("mi",), "đemi": ("mi",), "đinh": ("thẻ",), "rảnh": ("rãnh",)}   # gạch XMCL = không nung · demi = mi · gạch đinh = gạch thẻ
TRON_GOI = {"lot", "goi", "%", "trongoi", "ls", "lumpsum", "khoan"}
BO_TU = set("công tác thi và các loại bằng theo cho phần trên dưới hạng mục vị trí khu vực bao gồm".split())   # GIỮ trong/ngoài/nhà: trát trong 60k ↔ ngoài 105k

def _kd(s): return unicodedata.normalize("NFD", str(s or "").replace("Đ", "D").replace("đ", "d")).encode("ascii", "ignore").decode().lower()
def dvt(s): return re.sub(r"[\s.]", "", _kd(str(s or "").replace("²", "2").replace("³", "3"))).replace("mdai", "md")
def _so(v): return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None
def tu(s):
    """Từ khoá CÓ DẤU (bẫy đã biết: bỏ dấu ⇒ 'mài' cầu thang = 'mái', 'đá' = 'đã'). Tách số khỏi chữ: 100mm ⇒ 100 · mm."""
    t = unicodedata.normalize("NFC", str(s or "")).lower()
    t = re.sub(r"(\d)(\D)", r"\1 \2", re.sub(r"(\D)(\d)", r"\1 \2", t))
    out = set()
    for w in re.split(r"[\W_]+", t):
        if w and w not in BO_TU and (len(w) > 1 or w.isdigit()): out.update(DONG_NGHIA.get(w, (w,)))
    return out
def chon_dong(noi_dung, ung_vien):
    """Xếp hạng dòng NS giống dòng HĐ trong các ứng viên (cùng mã NS + ĐVT). Trọng số từ = ĐỘ HIẾM trong nhóm ứng viên (IDF):
    'ngạch', 'cửa' nặng; 'gạch', 'lát' (dòng nào cũng có) nhẹ; con số ×1,5. Điểm = F2 (nặng phủ-dòng-HĐ) · đối nghĩa trong↔ngoài ×0,5 · lệch 'không (nung)' ×0,7 (chỉ tính từ có trong nhóm).
    Lệch con số (100 ↔ 200) ⇒ điểm ×0,5. Trả [(điểm, dòng)] giảm dần."""
    A = tu(noi_dung); T = [tu(n["noi_dung"]) for n in ung_vien]
    if not A or not T: return []
    df = {}
    for t in T:
        for x in t: df[x] = df.get(x, 0) + 1
    w = lambda x: math.log(1 + len(T) / df[x]) * (1.5 if x.isdigit() else 1.0)
    A_v = {x for x in A if x in df}; wa = sum(w(x) for x in A_v); so_a = {x for x in A if x.isdigit()}
    xh = []
    for n, B in zip(ung_vien, T):
        if not B or not wa: continue
        ov = sum(w(x) for x in A_v & B); p, r = ov / sum(w(x) for x in B), ov / wa
        f = 5 * p * r / (4 * p + r) if p + r else 0.0             # F2: ưu tiên PHỦ HẾT chữ đặc trưng của dòng HĐ (ngạch cửa ≠ lát sàn)
        for x, y in DOI_NGHIA:
            if (x in A and y in B and x not in B) or (y in A and x in B and y not in B): f *= 0.5
        if ("không" in A) != ("không" in B): f *= 0.7                 # gạch nung ≠ gạch không nung
        so_b = {x for x in B if x.isdigit()}
        if so_a and so_b and not so_a & so_b: f *= 0.5
        xh.append((round(f, 2), n, B))
    xh.sort(key=lambda x: -x[0])
    if xh:                                                            # chữ ĐẶC TRƯNG nhất (hiếm nhất) dòng khớp phủ được — dòng thiếu nó không tính "gần giống"
        chung = A_v & xh[0][2]; dac_trung = max(chung, key=w) if chung else None
        xh = [xh[0]] + [x for x in xh[1:] if dac_trung is None or dac_trung in x[2]]
    return [(f, n) for f, n, _ in xh]

def ma_ns_luat(noi_dung, hd_noi_dung):
    """Mã NS gợi ý theo bảng luật gan_nhom_doi (28/09). Luật 'HĐ vận chuyển ⇒ cước' KHÔNG áp cho HĐ CUNG CẤP vật tư
    (HĐ NCC hay ghi 'giá đã gồm vận chuyển' ⇒ gạch 3.000đ bị so với cước 70đ)."""
    m = GN.ma_ns(noi_dung)[0]
    t = GN.na(hd_noi_dung)
    if m and m.startswith("NCC_") and "VAN CHUYEN" in t and "CUNG CAP" not in t: m = "DTC_VanChuyen"
    return m or ""

def doc_khung(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        v = lambda sh, r0: [r for r in wb[sh].iter_rows(min_row=r0, values_only=True) if r and r[0] is not None]
        n1 = [r for r in wb["N1_DanhMuc"].iter_rows(min_row=3, max_col=20, values_only=True)]
        ten_dt = {r[6]: r[7] or "" for r in n1 if len(r) > 7 and r[6]}
        ten_ns = {r[11]: r[12] or "" for r in n1 if len(r) > 12 and r[11]}
        nhom_ns = {r[16]: r[19] or "" for r in n1 if len(r) > 19 and r[16]}
        ns = [dict(phien_ban=r[0], loai=r[2], ma_ns=r[3], ma_cv=r[4], noi_dung=str(r[7] or ""), dvt=str(r[8] or ""), kl=_so(r[9]), dg=_so(r[10]),
                   gt=(_so(r[9]) or 0) * (_so(r[10]) or 0) if _so(r[9]) is not None and _so(r[10]) is not None else (_so(r[11]) or 0))
              for r in wb["N2_NganSach"].iter_rows(min_row=2, values_only=True) if r and r[3]]
        hd = {}
        for r in v("N6_HD_DoiTac", 2):
            h = hd.setdefault(r[0], dict(ma_hd=r[0], ma_doi_tac=r[2], doi_tac=ten_dt.get(r[2], ""), so_hd=str(r[4] or ""), ngay_ky=r[5], noi_dung=str(r[6] or ""),
                                         dang_hd=r[7] or "", gia_tri=0.0, so_phu_luc=0))
            h["gia_tri"] += _so(r[8]) or 0
            if str(r[1] or "").upper().startswith("PH"): h["so_phu_luc"] += 1
        boq = [dict(ma_hd=r[0], stt=str(r[1] if r[1] is not None else ""), noi_dung=str(r[4] or ""), dvt=str(r[5] or ""), kl=_so(r[6]), dg=_so(r[7]))
               for r in v("N5_BOQ_CDT", 2) if _so(r[7])]
        dong = defaultdict(list)
        for r in v("N7_HD_DoiTac_ChiTiet", 2):
            dong[r[0]].append(dict(stt=str(r[1] if r[1] is not None else ""), pham_vi=r[3] or "", noi_dung=str(r[4] or ""), dvt=str(r[5] or ""), kl=_so(r[6]),
                                   dg=_so(r[7]), tien=_so(r[8]) if _so(r[8]) is not None else (_so(r[6]) or 0) * (_so(r[7]) or 0),
                                   ma_ns=r[11] or nhom_ns.get(r[10], "") or "", nhom=r[10] or ""))
    finally: wb.close()
    return dict(ten_dt=ten_dt, ten_ns=ten_ns, ns=ns, hd=hd, dong=dict(dong), boq=boq)

def la_hd_chi_phi(ma): return bool(ma) and not str(ma).startswith("HD-CDT")

def ds_nguon(khung, so_nen):
    """Danh sách HĐ chọn để kiểm: HĐ chi phí trong khung + hồ sơ nền AI đã đọc bảng mà chưa nhập khung."""
    k = doc_khung(khung)
    ds = [dict(nguon="khung", ma=m, ten=f"{m} · {h['doi_tac'] or h['ma_doi_tac'] or ''}", so_hd=h["so_hd"], gia_tri=h["gia_tri"], so_dong=len(k["dong"].get(m, [])))
          for m, h in k["hd"].items() if la_hd_chi_phi(m)]
    for i, r in so_nen.items():
        bang = NEN.dong_la((r.get("ai") or {}).get("bang"))
        if bang and r.get("loai") in ("HD_DOI_TAC", "BAO_GIA", "CHON_THAU", "QUYET_TOAN") and not r.get("da_nhap_khung"):
            ds.append(dict(nguon="nen", ma=i, ten=f"{r.get('loai_ten') or r.get('loai')} · {(r.get('ai') or {}).get('doi_tac_ten') or r['ten']}", so_hd=(r.get("ai") or {}).get("so_hd") or "",
                           gia_tri=sum(NEN.tien_dong(x) for x in bang), so_dong=len(bang), file=r["ten"], trang_thai=r.get("trang_thai")))
    return ds

def _dong_tu_nen(rec):
    ai = rec.get("ai") or {}; out = []
    for x in NEN.dong_la(ai.get("bang")):
        out.append(dict(stt=str(x.get("stt") or ""), pham_vi="", noi_dung=str(x.get("noi_dung") or ""), dvt=str(x.get("dvt") or ""), kl=_so(x.get("kl")),
                        dg=_so(x.get("don_gia")), tien=NEN.tien_dong(x), ma_ns="", nhom=""))
    h = dict(ma_hd=None, ma_doi_tac=rec.get("ma_doi_tac"), doi_tac=ai.get("doi_tac_ten") or "", so_hd=ai.get("so_hd") or "", ngay_ky=ai.get("ngay_ky"),
             noi_dung=ai.get("noi_dung") or "", dang_hd=ai.get("dang_hd") or "", gia_tri=ai.get("gia_tri_truoc_vat") or sum(d["tien"] for d in out), so_phu_luc=0,
             file=rec["ten"], loai=rec.get("loai_ten") or rec.get("loai"))
    return h, out

def kiem(khung, ma_hd=None, rec_nen=None, k=None):
    k = k or doc_khung(khung)
    if rec_nen: h, dong = _dong_tu_nen(rec_nen)
    else:
        if ma_hd not in k["hd"]: raise ValueError(f"Không có HĐ {ma_hd} trong khung")
        h, dong = dict(k["hd"][ma_hd]), [dict(d) for d in k["dong"].get(ma_hd, [])]
    ghi_chu = []
    # ── ngân sách theo (mã NS, ĐVT)
    ns_md = defaultdict(lambda: dict(kl=0.0, gt=0.0, dong=[]))
    ns_ma = defaultdict(float)
    for n in k["ns"]:
        ns_ma[n["ma_ns"]] += n["gt"]
        g = ns_md[(n["ma_ns"], dvt(n["dvt"]))]; g["dong"].append(n)
        if n["kl"] is not None and n["dg"] is not None: g["kl"] += n["kl"]; g["gt"] += n["kl"] * n["dg"]
    # ── cam kết các HĐ chi phí KHÁC đã có trong khung
    khac_md, khac_ma, khac_ds = defaultdict(float), defaultdict(float), defaultdict(set)
    for m, ds in k["dong"].items():
        if m == ma_hd or not la_hd_chi_phi(m): continue
        for d in ds:
            mn = d["ma_ns"] or ma_ns_luat(d["noi_dung"], k["hd"].get(m, {}).get("noi_dung"))
            if not mn: continue
            khac_ma[mn] += d["tien"] or 0; khac_ds[mn].add(m)
            if d["kl"] is not None: khac_md[(mn, dvt(d["dvt"]))] += d["kl"]
    # ── từng dòng HĐ
    for d in dong:
        d["nguon_ma"] = "khung" if d["ma_ns"] else ""
        if not d["ma_ns"]:
            mn = ma_ns_luat(d["noi_dung"], h.get("noi_dung"))
            if mn: d["ma_ns"], d["nguon_ma"] = mn, "luật (gợi ý)"
        d["ten_ns"] = k["ten_ns"].get(d["ma_ns"], "")
        d.update(ns_noi_dung=None, ns_ma_cv=None, dg_ns=None, cach_so=None, do_giong=None, chenh_dg=None, tien_vuot=0.0)
        if not d["ma_ns"]: d.update(muc="XAM", ly_do="Chưa gán được mã NS — không so được"); continue
        g = ns_md.get((d["ma_ns"], dvt(d["dvt"])))
        if d["ma_ns"] not in ns_ma: d.update(muc="XAM", ly_do=f"Mã {d['ma_ns']} không có trong ngân sách"); continue
        if not g: d.update(muc="XAM", ly_do=f"ĐVT '{d['dvt']}' không có trong NS mã {d['ma_ns']} (NS dùng: {', '.join(sorted({n['dvt'] for n in k['ns'] if n['ma_ns'] == d['ma_ns']}))})"); continue
        if d["dg"] is None: d.update(muc="XAM", ly_do="Dòng HĐ không có đơn giá"); continue
        co_dg = [n for n in g["dong"] if n["dg"] is not None and n["kl"]]
        if dvt(d["dvt"]) in TRON_GOI: d.update(muc="XAM", ly_do="Khoản trọn gói — so theo GIÁ TRỊ ở bảng mã NS, không so đơn giá"); continue
        xh = chon_dong(d["noi_dung"], co_dg); diem, tot = xh[0] if xh else (0.0, None)
        dg_min, dg_max = min((n["dg"] for n in co_dg), default=0), max((n["dg"] for n in co_dg), default=0)
        if tot is not None and diem >= KHOP_TOI_THIEU:
            d.update(dg_ns=tot["dg"], ns_noi_dung=tot["noi_dung"], ns_ma_cv=tot["ma_cv"], cach_so="đúng dòng NS", do_giong=diem)
            chac = True
        elif g["kl"]:
            d.update(dg_ns=g["gt"] / g["kl"], do_giong=diem, cach_so=f"bình quân mã NS ({len(co_dg)} dòng)", ns_noi_dung=f"{d['ma_ns']} · {d['dvt']} — ĐG NS từ {dg_min:,.0f} đến {dg_max:,.0f}")
            chac = dg_min > 0 and dg_max / dg_min <= LECH_BQ
        else: d.update(muc="XAM", ly_do="NS mã này không có đơn giá theo dòng"); continue
        if not d["dg_ns"]: d.update(muc="XAM", ly_do="Đơn giá NS = 0"); continue
        d["chenh_dg"] = d["dg"] / d["dg_ns"] - 1
        if d["chenh_dg"] > NGUONG_DG:
            d["tien_vuot"] = (d["dg"] - d["dg_ns"]) * (d["kl"] or 0)
            d.update(muc="DO" if chac else "VANG", ly_do=f"ĐG HĐ vượt NS {d['chenh_dg']:+.1%}" + ("" if chac else " — NS không có dòng tương ứng, so BÌNH QUÂN mã NS gồm nhiều công việc ⇒ chỉ tham khảo"))
        else:
            d.update(muc="XANH", ly_do="ĐG HĐ ≤ NS" if d["chenh_dg"] < -NGUONG_DG else "ĐG HĐ = NS")
            gan = [n for f, n in xh[1:] if f >= diem - SAN_SAN and n["dg"] and d["dg"] > n["dg"] * (1 + NGUONG_DG)] if chac and d["cach_so"] == "đúng dòng NS" and diem < 0.85 else []
            if gan:
                r = min(gan, key=lambda n: n["dg"]); d["ns_gan"] = dict(noi_dung=r["noi_dung"], dg=r["dg"], ma_cv=r["ma_cv"])
                d.update(muc="VANG", ly_do=f"Khớp dòng NS CHƯA CHẮC (độ giống {diem:.0%}): ĐG HĐ ≤ dòng khớp nhưng VƯỢT dòng gần giống '{r['noi_dung'][:40]}' ({r['dg']:,.0f}) — anh xem đúng dòng nào")
        if d.get("pham_vi") == "NGOAI_HD": d["ly_do"] += " · dòng NGOÀI HĐ (phát sinh)"
    # ── khối lượng theo (mã NS, ĐVT)
    kl = []
    for key in sorted({(d["ma_ns"], dvt(d["dvt"])) for d in dong if d["ma_ns"] and d["kl"] is not None}):
        g = ns_md.get(key); nay = sum(d["kl"] for d in dong if (d["ma_ns"], dvt(d["dvt"])) == key and d["kl"] is not None)
        x = dict(ma_ns=key[0], ten_ns=k["ten_ns"].get(key[0], ""), dvt=next(d["dvt"] for d in dong if (d["ma_ns"], dvt(d["dvt"])) == key), kl_ns=g["kl"] if g else None,
                 kl_nay=nay, kl_khac=khac_md.get(key, 0.0))
        x["kl_tong"] = x["kl_nay"] + x["kl_khac"]
        if not g or not g["kl"]: x.update(muc="XAM", ly_do="NS không có KL cho ĐVT này")
        elif x["kl_nay"] > g["kl"] * (1 + NGUONG_DG): x.update(muc="VANG", ly_do=f"KL HĐ này vượt KL NS ({x['kl_nay'] / g['kl']:.0%} NS)")
        elif x["kl_tong"] > g["kl"] * (1 + NGUONG_DG): x.update(muc="VANG", ly_do=f"Cộng các HĐ đã ký vượt KL NS ({x['kl_tong'] / g['kl']:.0%} NS)")
        else: x.update(muc="XANH", ly_do=f"Cộng dồn {x['kl_tong'] / g['kl']:.0%} KL NS")
        kl.append(x)
    # ── giá trị theo mã NS
    gt = []
    for mn in sorted({d["ma_ns"] for d in dong if d["ma_ns"]}):
        nay = sum(d["tien"] or 0 for d in dong if d["ma_ns"] == mn)
        x = dict(ma_ns=mn, ten_ns=k["ten_ns"].get(mn, ""), gt_ns=ns_ma.get(mn), gt_nay=nay, gt_khac=khac_ma.get(mn, 0.0), hd_khac=sorted(khac_ds.get(mn, ())))
        x["gt_tong"] = x["gt_nay"] + x["gt_khac"]
        if not x["gt_ns"]: x.update(muc="XAM", ly_do="Mã không có ngân sách")
        elif x["gt_nay"] > x["gt_ns"] * (1 + NGUONG_DG): x.update(muc="DO", ly_do=f"Riêng HĐ này đã vượt NS mã ({x['gt_nay'] / x['gt_ns']:.0%})")
        elif x["gt_tong"] > x["gt_ns"] * (1 + NGUONG_DG): x.update(muc="VANG", ly_do=f"Cộng {len(x['hd_khac'])} HĐ đã ký + HĐ này = {x['gt_tong'] / x['gt_ns']:.0%} NS")
        else: x.update(muc="XANH", ly_do=f"Cộng dồn {x['gt_tong'] / x['gt_ns']:.0%} NS")
        gt.append(x)
    if "DON_GIA" in str(h.get("dang_hd")).upper() or "NGUYEN_TAC" in str(h.get("dang_hd")).upper():
        ghi_chu.append("HĐ đơn giá / nguyên tắc: KL trong HĐ có thể là KL tạm tính ⇒ cờ KL / giá trị (vàng) để anh cân nhắc; cờ ĐƠN GIÁ (đỏ) là cờ chắc chắn.")
    if any(d["nguon_ma"] == "luật (gợi ý)" for d in dong): ghi_chu.append("Một số dòng chưa gán nhóm trong khung — mã NS là GỢI Ý theo bảng luật 28/09, chưa ghi vào khung.")
    dem = defaultdict(int)
    for d in dong: dem[d["muc"]] += 1
    tong = dict(tien=sum(d["tien"] or 0 for d in dong), tien_co_ns=sum(d["tien"] or 0 for d in dong if d["muc"] != "XAM"), so_dong=len(dong), dem=dict(dem),
                tien_vuot_dg=sum(d["tien_vuot"] for d in dong if d["muc"] == "DO"), tien_vuot_tk=sum(d["tien_vuot"] for d in dong if d["muc"] == "VANG"),
                ma_ns_vuot=sum(1 for x in gt if x["muc"] in ("DO", "VANG")))
    return dict(hd=h, tong=tong, dong=dong, kl=kl, gt=gt, ghi_chu=ghi_chu, nguong=dict(dg=NGUONG_DG, lech_bq=LECH_BQ, khop=KHOP_TOI_THIEU))

def ptln(khung, nguon, so_nen):
    """PTLN — so giá NHIỀU đơn vị nhận thầu ↔ ngân sách ↔ BoQ CĐT. nguon = [("khung", mã HĐ) | ("nen", id hồ sơ nền)].
    Hàng = DÒNG NGÂN SÁCH mà ít nhất 1 đơn vị báo giá ghép được (lõi ghép dòng dùng chung với kiem()).
    Tổng mỗi đơn vị = Σ ĐG × KL NS — xếp hạng chỉ trên các dòng MỌI đơn vị cùng báo (so công bằng).
    BoQ CĐT: ghép theo nội dung + ĐVT; KHÔNG tính TSLN vì giá CĐT có thể gồm vật tư, giá đội chỉ nhân công ⇒ chỉ hiện tỷ lệ giá ĐV / giá bán."""
    k = doc_khung(khung); dv, hang, le = [], {}, []
    for i, (ng, ma) in enumerate(nguon):
        r = kiem(khung, rec_nen=so_nen[ma], k=k) if ng == "nen" else kiem(khung, ma, k=k)
        h = r["hd"]; dv.append(dict(i=i, nguon=ng, ma=ma, ten=h.get("doi_tac") or h.get("ma_hd") or ma, ma_hd=h.get("ma_hd"), so_hd=h.get("so_hd"), gia_tri=h.get("gia_tri"), so_dong=r["tong"]["so_dong"]))
        for d in r["dong"]:
            if d.get("pham_vi") == "NGOAI_HD": continue                   # phát sinh ngoài HĐ không phải giá chào thầu
            if d["cach_so"] == "đúng dòng NS" and d["muc"] != "XAM":
                key = (d["ma_ns"], d["ns_ma_cv"], d["ns_noi_dung"], dvt(d["dvt"]))
                x = hang.setdefault(key, dict(ma_ns=d["ma_ns"], ten_ns=d["ten_ns"], ma_cv=d["ns_ma_cv"], noi_dung=d["ns_noi_dung"], dvt=d["dvt"], dg_ns=d["dg_ns"], gia={}))
                x["gia"].setdefault(i, []).append(dict(dg=d["dg"], noi_dung=d["noi_dung"], stt=d["stt"], giong=d["do_giong"]))
            elif d["dg"]: le.append(dict(i=i, stt=d["stt"], noi_dung=d["noi_dung"], dvt=d["dvt"], dg=d["dg"], ma_ns=d["ma_ns"], ly_do=d.get("ly_do"), dg_ns=d["dg_ns"], cach_so=d["cach_so"]))
    kl_ns = defaultdict(float)
    for n in k["ns"]:
        if n["kl"] is not None: kl_ns[(n["ma_ns"], n["ma_cv"], n["noi_dung"], dvt(n["dvt"]))] += n["kl"]
    boq = defaultdict(list)
    for b in k["boq"]: boq[dvt(b["dvt"])].append(b)
    ds = []
    for key, x in hang.items():
        x["kl_ns"] = kl_ns.get(key)
        xh = chon_dong(x["noi_dung"], boq.get(key[3], []))
        if xh and xh[0][0] >= 0.5: x.update(dg_boq=xh[0][1]["dg"], boq_noi_dung=xh[0][1]["noi_dung"], boq_giong=xh[0][0])
        else: x.update(dg_boq=None, boq_noi_dung=None, boq_giong=None)
        for i, g in x["gia"].items():
            for y in g: y["chenh"] = y["dg"] / x["dg_ns"] - 1 if x["dg_ns"] else None
        gm = {i: max(y["dg"] for y in g) for i, g in x["gia"].items()}       # 1 ĐV nhiều dòng cùng ghép 1 dòng NS ⇒ lấy giá CAO nhất (thận trọng)
        x["dg_dv"] = gm; x["thap_nhat"] = min(gm.values()) if gm else None
        x["dv_thap"] = [i for i, v in gm.items() if v == x["thap_nhat"]]
        ds.append(x)
    ds.sort(key=lambda x: (x["ma_ns"] or "", x["ma_cv"] or "", x["noi_dung"]))
    chung = [x for x in ds if len(x["dg_dv"]) == len(dv)]
    for d in dv:
        i = d["i"]; co = [x for x in ds if i in x["dg_dv"]]
        d.update(so_hang=len(co), vuot_ns=sum(1 for x in co if x["dg_ns"] and x["dg_dv"][i] > x["dg_ns"] * (1 + NGUONG_DG)),
                 gt_theo_kl_ns=sum(x["dg_dv"][i] * (x["kl_ns"] or 0) for x in co), gt_chung=sum(x["dg_dv"][i] * (x["kl_ns"] or 0) for x in chung),
                 gt_ns_chung=sum((x["dg_ns"] or 0) * (x["kl_ns"] or 0) for x in chung), so_le=sum(1 for l in le if l["i"] == i))
    if len(dv) > 1 and chung:
        for d in dv: d["hang"] = 1 + sum(1 for e in dv if e["gt_chung"] < d["gt_chung"] - 1)   # bằng giá ⇒ đồng hạng
    ghi_chu = [f"Xếp hạng trên {len(chung)} dòng NS mà cả {len(dv)} đơn vị cùng báo giá (Σ ĐG × KL ngân sách)." if len(dv) > 1 else "Chọn ≥ 2 đơn vị để xếp hạng."]
    if any(x["dg_boq"] for x in ds): ghi_chu.append("Giá BoQ CĐT có thể GỒM vật tư + nhân công; giá đơn vị nhận thầu thường chỉ nhân công ⇒ chỉ so TỶ LỆ, không tính TSLN.")
    return dict(don_vi=dv, hang=ds, le=le, so_chung=len(chung), ghi_chu=ghi_chu)

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    r = kiem(sys.argv[1], sys.argv[2])
    print(r["hd"]["ma_hd"], r["hd"]["doi_tac"], r["tong"])
    for d in r["dong"]: print(f"{d['muc']:5} {d['stt']:>5} {d['noi_dung'][:38]:38} {d['dvt']:5} ĐG {d['dg'] or 0:>11,.0f} | NS {d['dg_ns'] or 0:>11,.0f} {d['cach_so'] or '':22} {(d['ns_noi_dung'] or '')[:34]:34} | {d['ly_do']}")
    for x in r["kl"]: print("KL", x["muc"], x["ma_ns"], x["dvt"], x["kl_ns"], x["kl_nay"], x["kl_khac"], x["ly_do"])
    for x in r["gt"]: print("GT", x["muc"], x["ma_ns"], round(x["gt_ns"] or 0), round(x["gt_nay"]), round(x["gt_khac"]), x["ly_do"])
