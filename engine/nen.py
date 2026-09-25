"""HỒ SƠ NỀN — người dùng CHỈ NẠP FILE; app tự nhận loại, tự đọc (Claude Code gói Pro qua ai_doc), tự đối chiếu, tự xếp folder.
Luồng: nạp → lưu tạm _HE_THONG\\cho_phan_loai (chỉ đọc) → hàng đợi AI (1 luồng) → cờ tính bằng CODE → chuyển file đúng folder →
CHO_DUYET (HĐ / BoQ / ngân sách / gói thầu: chờ anh duyệt để ghi khung) hoặc DA_LUU (chứng từ). Engine khởi động lại ⇒ đọc tiếp việc dở.
HSTT PDF = bản ký gắn vào HSTT Excel cùng HĐ + đợt (dinh_kem / xep_dinh_kem). Sổ: <DATA>\\<dự án>\\_HE_THONG\\ho_so_nen.json."""
import os, re, json, hashlib, shutil, stat, threading, queue, traceback, unicodedata, datetime as dt, openpyxl, warnings
warnings.filterwarnings("ignore")
DUOI = (".pdf", ".xlsx", ".xlsm", ".xls", ".doc", ".docx", ".jpg", ".jpeg", ".png")
AI_DOC_DUOC = (".pdf", ".xlsx", ".xlsm")
LOAI = {   # mã: (tên hiển thị, thư mục tương đối — {dt} = <LOAI>_<mã đối tác>, {goi} = mã gói)
    "BOQ_CDT":   ("BoQ / dự toán HĐ với CĐT", "01_THIET_LAP/01_BOQ_CDT"),
    "NGAN_SACH": ("Ngân sách (R00 / điều chỉnh)", "01_THIET_LAP/02_NGAN_SACH"),
    "GOI_THAU":  ("Phân chia gói thầu", "01_THIET_LAP/03_GOI_THAU"),
    "CHON_THAU": ("Báo giá chọn thầu (chưa có HĐ)", "01_THIET_LAP/03_GOI_THAU/CHON_THAU/{goi}"),
    "HD_CDT":    ("Hợp đồng / PLHĐ với CĐT", "10_THU_CDT/HOP_DONG"),
    "HD_DOI_TAC": ("Hợp đồng / PLHĐ đối tác", "20_CHI_DOI_TAC/{dt}/HOP_DONG"),
    "BAO_GIA":   ("Báo giá / bảng giá đối tác", "20_CHI_DOI_TAC/{dt}/BAO_GIA"),
    "QUYET_TOAN": ("Hồ sơ quyết toán đối tác", "20_CHI_DOI_TAC/{dt}/QUYET_TOAN"),
    "KHAC":      ("Hồ sơ khác (biên bản, tờ trình…)", "01_THIET_LAP/99_KHAC"),
    "HSTT_KY":   ("HSTT bản ký (PDF) kèm HSTT Excel", "")}
AI_SANG_LOAI = {"HD_CDT": "HD_CDT", "PLHD_CDT": "HD_CDT", "HD_DOI_TAC": "HD_DOI_TAC", "PLHD_DOI_TAC": "HD_DOI_TAC", "BOQ_CDT": "BOQ_CDT",
                "NGAN_SACH": "NGAN_SACH", "GOI_THAU": "GOI_THAU", "BAO_GIA": "BAO_GIA", "QUYET_TOAN": "QUYET_TOAN",
                "BIEN_BAN": "KHAC", "TO_TRINH": "KHAC", "HSTT": "KHAC", "KHAC": "KHAC"}
CAN_NHAP = {"HD_CDT", "HD_DOI_TAC", "BOQ_CDT", "NGAN_SACH", "GOI_THAU"}          # loại phải ghi vào khung ⇒ chờ anh duyệt
LOAI_DT = {"ĐTC": "DTC", "DTC": "DTC", "NTP": "NTP", "NCC": "NCC", "DVK": "DVK"}
KHOA = threading.RLock()

def _sach(s): return re.sub(r"[^A-Za-z0-9_.-]", "", str(s or "")) or "KHAC"
def khong_dau(s): return unicodedata.normalize("NFD", str(s or "").replace("Đ", "D").replace("đ", "d")).encode("ascii", "ignore").decode()
def van_tay(raw): return hashlib.sha256(raw).hexdigest()
def so_path(data, da): return os.path.join(data, da, "_HE_THONG", "ho_so_nen.json")
def doc_so(data, da):
    p = so_path(data, da); return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}
def ghi_so_nen(data, da, s):
    os.makedirs(os.path.dirname(so_path(data, da)), exist_ok=True)
    tmp = so_path(data, da) + ".tmp"; json.dump(s, open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=1); os.replace(tmp, so_path(data, da))
def sua_rec(data, da, i, **kv):
    with KHOA:
        s = doc_so(data, da); s[i].update(kv); ghi_so_nen(data, da, s); return s[i]

def danh_muc(khung):
    """Danh sách chọn: đối tác (N1), HĐ (N4/N6), gói (N3) — đọc từ file khung."""
    wb = openpyxl.load_workbook(khung, read_only=True, data_only=True)
    dt_ = [dict(ma=r[6], ten=r[7] or "", loai=r[8] or "") for r in wb["N1_DanhMuc"].iter_rows(min_row=3, max_col=9, values_only=True) if r[6]]
    hd = [dict(ma_hd=r[0], ben="CĐT", ma_doi_tac=r[2], so_hd=r[3]) for r in wb["N4_HD_CDT"].iter_rows(min_row=2, max_col=4, values_only=True) if r[0]]
    hd += [dict(ma_hd=r[0], ben="ĐỐI TÁC", ma_doi_tac=r[2], so_hd=r[4]) for r in wb["N6_HD_DoiTac"].iter_rows(min_row=2, max_col=5, values_only=True) if r[0]]
    goi = [dict(ma=r[0], ten=r[1] or "") for r in wb["N3_GoiThau"].iter_rows(min_row=2, max_col=2, values_only=True) if r[0]]
    wb.close()
    return dict(doi_tac=dt_, hop_dong=hd, goi=goi, loai=[dict(ma=k, ten=v[0]) for k, v in LOAI.items() if k != "HSTT_KY"])

TIEN_TO = r"^(to doi|doi thi cong|doi|cong ty co phan|cong ty tnhh mtv|cong ty tnhh|cong ty cp|cong ty|ctcp|cty|tnhh|mtv|sx|tm|dv|xd|thuong mai|xay dung|kien truc|san xuat|dau tu)\s+"
def ma_de_xuat(ten):
    """Quy ước mã cũ: người/tổ đội ⇒ chữ đầu các từ + từ cuối ('Tổ đội Trần Thành Thắng' ⇒ TTThang);
    công ty ⇒ ghép 2 từ cuối ('Công ty TNHH TM DV XD Mỹ Kim' ⇒ MyKim); có viết tắt trong ngoặc ⇒ dùng luôn ('…(DECOFI)' ⇒ DECOFI)."""
    m = re.search(r"[\([\[]([A-Z0-9&]{3,12})[\)\]]", khong_dau(ten))
    if m: return m.group(1).replace("&", "")
    t = re.sub(r"\(.*?\)", " ", khong_dau(ten).lower()).strip(); cong_ty = bool(re.match(r"(cong ty|ctcp|cty|doanh nghiep)", t))
    for _ in range(6): t = re.sub(TIEN_TO, "", t.strip())
    if not cong_ty: t = re.sub(r"^(ong|ba|anh|chi)\s+", "", t)
    w = [x for x in re.split(r"[^a-z0-9]+", t) if x]
    if not w: return "DT_MOI"
    if cong_ty: return "".join(x.capitalize() for x in w[-2:])[:16]
    return ("".join(x[0].upper() for x in w[:-1]) + w[-1].capitalize())[:16]
def khop_doi_tac(ten, ds):
    """Tìm đối tác đã có trong khung theo tên (bỏ dấu, bỏ tiền tố). Trả dict hoặc None."""
    k = ma_de_xuat(ten).lower(); t = set(re.split(r"[^a-z0-9]+", khong_dau(ten).lower())) - {""}
    for d in ds:
        if d["ma"].lower() == k: return d
        td = set(re.split(r"[^a-z0-9]+", khong_dau(d["ten"]).lower())) - {"", "to", "doi", "cong", "ty"}
        if td and len(td & t) >= max(2, len(td) - 1): return d
    return None

def luu_file(data, da, rel_dir, ten, raw):
    """Ghi file vào <DATA>/<da>/<rel_dir>/ (chỉ đọc). Trùng tên khác nội dung ⇒ thêm _v2, _v3…"""
    d = os.path.normpath(os.path.join(data, da, rel_dir)); os.makedirs(d, exist_ok=True)
    goc, duoi = os.path.splitext(ten); dich = os.path.join(d, ten); k = 1
    while os.path.exists(dich):
        if van_tay(open(dich, "rb").read()) == van_tay(raw): return dich
        k += 1; dich = os.path.join(d, f"{goc}_v{k}{duoi}")
    open(dich, "wb").write(raw); os.chmod(dich, 0o444); return dich

def thu_muc_dich(loai, ma_dt=None, loai_dt=None, goi=None):
    mau = LOAI[loai][1]
    if "{dt}" in mau:
        l = LOAI_DT.get(loai_dt or "")
        if not (ma_dt and l): return None                                     # chưa biết đối tác ⇒ chưa xếp được
        mau = mau.replace("{dt}", f"{l}_{_sach(ma_dt)}")
    if "{goi}" in mau: mau = mau.replace("{goi}", _sach(goi) if goi else "CHUA_GAN_GOI")
    return mau

def nap_nen(data, da, khung, ten, raw, loai=None, ma_dt=None, loai_dt=None, goi=None, ghi_chu=""):
    """Lưu file + đăng ký. loai rỗng ⇒ app tự nhận (AI). Trả bản ghi (trung=True nếu đã có)."""
    ten = os.path.basename(ten)
    if not ten.lower().endswith(DUOI): raise ValueError(f"'{ten}': chỉ nhận {', '.join(DUOI)}")
    if loai and loai not in LOAI: raise ValueError("Loại hồ sơ không hợp lệ")
    vt = van_tay(raw)
    with KHOA:
        s = doc_so(data, da); cu = next((v for v in s.values() if v["van_tay"] == vt), None)
        if cu: return dict(cu, trung=True)
        if loai and ma_dt and not loai_dt:
            loai_dt = {d["ma"]: d["loai"] for d in danh_muc(khung)["doi_tac"]}.get(ma_dt)
        rel = thu_muc_dich(loai, ma_dt, loai_dt, goi) if loai else None
        dich = luu_file(data, da, rel or "_HE_THONG/cho_phan_loai", ten if rel else f"{vt[:12]}_{ten}", raw)
        ai = ten.lower().endswith(AI_DOC_DUOC)
        rec = dict(id=vt[:12], van_tay=vt, ten=ten, loai=loai or None, loai_ten=LOAI[loai][0] if loai else "🤖 app đang nhận loại", loai_chon=loai or None,
                   ma_doi_tac=ma_dt, loai_doi_tac=LOAI_DT.get(loai_dt or "", loai_dt), goi=goi, ghi_chu=ghi_chu, duong_dan=os.path.relpath(dich, data),
                   luc=dt.datetime.now().isoformat(timespec="seconds"), trang_thai="CHO_AI" if ai else ("DA_LUU" if rel else "CHO_PHAN_LOAI"),
                   da_nhap_khung=False, can_nhap=bool(loai in CAN_NHAP), co=[], ai=None)
        s[rec["id"]] = rec; ghi_so_nen(data, da, s)
    if ai: HANG.put((data, da, rec["id"]))
    return rec

# ───────── đánh giá kết quả AI bằng CODE (không tin mù) ─────────
def dong_la(bang):
    """Bỏ dòng tiêu đề nhóm / dòng cộng: không ĐVT và không KL."""
    return [r for r in (bang or []) if (r.get("dvt") or r.get("kl") is not None) and not re.match(r"^\s*(cộng|tổng)", str(r.get("noi_dung") or ""), re.I)]
def tien_dong(r):
    """Tiền 1 dòng — CÙNG quy tắc với nhap_khung.dong_hop_le: có thành tiền in trên HĐ thì tin thành tiền."""
    if isinstance(r.get("thanh_tien"), (int, float)): return r["thanh_tien"]
    if isinstance(r.get("kl"), (int, float)) and isinstance(r.get("don_gia"), (int, float)): return r["kl"] * r["don_gia"]
    return 0
def danh_gia(kq, loai, loai_chon, dm):
    co = []; C = lambda m, t: co.append(dict(muc=m, mo_ta=t))
    la = dong_la(kq.get("bang")); tong = sum(tien_dong(r) for r in la)
    if loai_chon and AI_SANG_LOAI.get(kq.get("loai")) != loai_chon: C("LUU_Y", f"Anh chọn loại {LOAI[loai_chon][0]} nhưng app đọc thấy là {kq.get('loai')} — {kq.get('ly_do_loai') or ''}")
    if loai in ("HD_CDT", "HD_DOI_TAC"):
        for f, t in (("so_hd", "số hợp đồng"), ("ngay_ky", "ngày ký"), ("doi_tac_ten", "tên đối tác")):
            if not kq.get(f): C("CHAN", f"Không đọc được {t} — anh kiểm lại PDF")
        if kq.get("dang_hd") != "NGUYEN_TAC" and not kq.get("gia_tri_truoc_vat"): C("CHAN", "Không đọc được giá trị HĐ trước VAT")
        if kq.get("pct_tt_dot") is None: C("LUU_Y", "Không thấy % thanh toán mỗi đợt")
        if kq.get("han_tt_ngay") is None: C("LUU_Y", "Không thấy hạn thanh toán (ngày)")
        g = kq.get("gia_tri_truoc_vat")
        if g and la and abs(tong - g) > g * 0.005: C("CHAN", f"Σ bảng đơn giá {tong:,.0f} ≠ giá trị HĐ {g:,.0f} (lệch {tong - g:,.0f}, {abs(tong - g) / g:.1%}) — AI có thể đọc sai dòng; anh kiểm bảng rồi bấm Đọc lại")
        elif g and la and abs(tong - g) > 1000: C("LUU_Y", f"Σ bảng đơn giá {tong:,.0f} ≠ giá trị HĐ {g:,.0f} (lệch {tong - g:,.0f}, làm tròn)")
        miss = [str(x.get("stt") or x.get("noi_dung"))[:20] for x in la if x.get("kl") is None and x.get("thanh_tien") is None]
        if miss: C("LUU_Y", f"{len(miss)} dòng không đọc được KL / thành tiền: {', '.join(miss[:6])}")
        if kq.get("co_chu_ky") is False: C("LUU_Y", "Bản PDF chưa thấy chữ ký")
        if kq.get("co_dong_dau") is False: C("LUU_Y", "Bản PDF chưa thấy đóng dấu")
    if loai in ("BOQ_CDT", "NGAN_SACH", "GOI_THAU"):
        if not la: C("CHAN", "Không đọc được dòng nào trong bảng")
        t = kq.get("tong_ghi_tren_file")
        if t and la and abs(tong - t) > max(1000, t * 0.0005): C("LUU_Y", f"Σ các dòng {tong:,.0f} ≠ tổng ghi trên file {t:,.0f} (lệch {tong - t:,.0f})")
    for x in kq.get("khong_chac") or []: C("LUU_Y", f"AI đọc không chắc: {x}")
    if kq.get("loai") == "HSTT": C("LUU_Y", "Đây là HSTT — nạp ở mục ① Nạp & duyệt (Excel để soát, PDF làm bản ký)")
    return co, dict(so_dong=len(la), tong_dong=tong)

PCT = ("vat_pct", "pct_tam_ung", "pct_tt_dot", "pct_tt_quyet_toan", "pct_giu_lai")
def tu_khoa_cty(cty):
    """'VELA (Công ty CP Kỹ thuật Xây dựng VELA)' ⇒ ['VELA'] — từ khoá nhận ra công ty người dùng trong HĐ."""
    goc = str(cty or "").split("(")[0].strip()
    return [khong_dau(goc).upper()] if goc and not goc.startswith("(chưa") else []
def la_cty_minh(ten, kw): t = khong_dau(ten).upper(); return any(k and re.search(r"\b" + re.escape(k) + r"\b", t) for k in kw)
def chuan_hoa(kq, cty=None):
    """KHÔNG tin AI về định dạng / phân loại: % dạng 90 ⇒ 0.9 (ngoài 0..1 ⇒ CHẶN); công ty mình là BÊN NHẬN ⇒ HĐ phía CĐT, đối tác = bên giao;
    đối tác trùng công ty mình ⇒ CHẶN; tổ đội / giao khoán ⇒ DTC. Trả danh sách cờ."""
    co = []; kw = tu_khoa_cty(cty if cty is not None else _CFG.get("cty"))
    if kw and kq.get("loai") in ("HD_CDT", "HD_DOI_TAC", "PLHD_CDT", "PLHD_DOI_TAC"):
        nhan, giao = la_cty_minh(kq.get("ben_nhan"), kw), la_cty_minh(kq.get("ben_giao"), kw)
        if nhan and not giao and kq["loai"] in ("HD_DOI_TAC", "PLHD_DOI_TAC"):
            kq["loai"] = kq["loai"].replace("DOI_TAC", "CDT"); co.append(dict(muc="LUU_Y", mo_ta="Công ty mình là BÊN NHẬN thầu ⇒ app đổi thành HĐ phía CĐT (doanh thu)"))
        if nhan and not giao: kq["doi_tac_ten"], kq["loai_doi_tac"] = kq.get("ben_giao"), "CDT"
        if giao and not nhan and kq["loai"] in ("HD_CDT", "PLHD_CDT"):
            kq["loai"] = kq["loai"].replace("CDT", "DOI_TAC"); co.append(dict(muc="LUU_Y", mo_ta="Công ty mình là BÊN GIAO việc ⇒ app đổi thành HĐ đối tác (chi phí)"))
        if giao and not nhan: kq["doi_tac_ten"] = kq.get("ben_nhan")
    if kq.get("doi_tac_ten"):                                         # bỏ phần người đại diện: "… - Ông Chu Quang Huân, P.TGĐ (ủy quyền…)"
        kq["doi_tac_ten"] = re.split(r"\s+-\s+(?:Ông|Bà|Ong|Ba)\b|\s*\((?:đại diện|dai dien)", kq["doi_tac_ten"])[0].strip(" -,")
        if la_cty_minh(kq.get("doi_tac_ten"), kw): co.append(dict(muc="CHAN", mo_ta="Đối tác trùng tên công ty mình — không xác định được bên nào là đối tác"))
    for k in PCT:
        v = kq.get(k)
        if isinstance(v, (int, float)) and 1.0001 < v <= 100: kq[k] = round(v / 100, 6); co.append(dict(muc="LUU_Y", mo_ta=f"{k}: AI trả {v} ⇒ app đổi thành {kq[k]:.2%}"))
        elif isinstance(v, (int, float)) and not (0 <= kq[k] <= 1): co.append(dict(muc="CHAN", mo_ta=f"{k} = {v} không hợp lệ (phải 0–100%)"))
    chu = khong_dau(" ".join(str(kq.get(x) or "") for x in ("doi_tac_ten", "ben_nhan", "so_hd", "noi_dung"))).lower()
    if kq.get("loai_doi_tac") in ("NTP", None) and re.search(r"\bto doi\b|doi thi cong|giao khoan|khoan nhan cong|hdgk", chu):
        co.append(dict(muc="LUU_Y", mo_ta=f"AI xếp loại đối tác {kq.get('loai_doi_tac')} nhưng HĐ là giao khoán / tổ đội ⇒ app đổi thành Đội thi công (DTC)")); kq["loai_doi_tac"] = "DTC"
    return co
def xu_ly_ai(data, da, i, khung, cty):
    import ai_doc
    s = doc_so(data, da); rec = s[i]; f = os.path.join(data, rec["duong_dan"])
    try: kq, meta = ai_doc.doc(f, cty=cty)
    except Exception as e:
        sua_rec(data, da, i, trang_thai="LOI_AI", loi=str(e)[:400]); return
    ap_ket_qua(data, da, i, khung, kq, meta)
def ap_ket_qua(data, da, i, khung, kq, meta):
    """Áp kết quả AI: chuẩn hoá, đánh giá cờ, nhận đối tác, xếp file. Tách riêng để áp lại được mà không phải đọc lại."""
    s = doc_so(data, da); rec = s[i]; co0 = chuan_hoa(kq)
    dm = danh_muc(khung)
    loai = rec.get("loai_chon") or AI_SANG_LOAI.get(kq.get("loai"), "KHAC")
    ma_dt, loai_dt = rec.get("ma_doi_tac"), rec.get("loai_doi_tac")
    if LOAI[loai][1].find("{dt}") >= 0 and not ma_dt:
        d = khop_doi_tac(kq.get("doi_tac_ten") or "", dm["doi_tac"])
        ma_dt, loai_dt = (d["ma"], LOAI_DT.get(d["loai"], d["loai"])) if d else (ma_de_xuat(kq.get("doi_tac_ten") or ""), kq.get("loai_doi_tac"))
    elif ma_dt and not loai_dt: loai_dt = kq.get("loai_doi_tac")
    if loai == "HD_CDT" and not ma_dt:                                  # HĐ doanh thu: đối tác = CĐT / thầu chính
        d = khop_doi_tac(kq.get("doi_tac_ten") or "", dm["doi_tac"]); ma_dt, loai_dt = (d["ma"] if d else ma_de_xuat(kq.get("doi_tac_ten") or "")), "CDT"
    if loai == "BAO_GIA" and not (ma_dt and LOAI_DT.get(loai_dt or "")): loai = "CHON_THAU"
    co, tk = danh_gia(kq, loai, rec.get("loai_chon"), dm); co = co0 + co
    moi_dt = bool(ma_dt) and not any(d["ma"] == ma_dt for d in dm["doi_tac"])
    if moi_dt and loai in ("HD_DOI_TAC", "BAO_GIA", "QUYET_TOAN"): co.append(dict(muc="LUU_Y", mo_ta=f"Đối tác mới (chưa có trong khung) — app đề xuất mã {ma_dt}, loại {loai_dt}"))
    rel = thu_muc_dich(loai, ma_dt, loai_dt, rec.get("goi")) or "_HE_THONG/cho_phan_loai"
    with KHOA:
        s = doc_so(data, da); rec = s[i]; cu = os.path.join(data, rec["duong_dan"])
        if not rel.startswith("_HE_THONG"):
            dich = luu_file(data, da, rel, rec["ten"], open(cu, "rb").read())
            if os.path.normcase(dich) != os.path.normcase(cu) and "cho_phan_loai" in cu:
                os.chmod(cu, stat.S_IWRITE); os.remove(cu)
            rec["duong_dan"] = os.path.relpath(dich, data)
        rec.update(ai=kq, ai_meta=meta, co=co, thong_ke=tk, loai=loai, loai_ten=LOAI[loai][0], ma_doi_tac=ma_dt, loai_doi_tac=loai_dt, doi_tac_moi=moi_dt,
                   can_nhap=loai in CAN_NHAP, trang_thai="DA_NHAP" if rec.get("da_nhap_khung") else ("CHO_DUYET" if loai in CAN_NHAP else "DA_LUU"), loi=None, luc_ai=dt.datetime.now().isoformat(timespec="seconds"))
        s[i] = rec; ghi_so_nen(data, da, s)

HANG = queue.Queue(); _CFG = {}
def _tho():
    while True:
        data, da, i = HANG.get()
        try:
            cfg = _CFG["du_an"]()[da]
            xu_ly_ai(data, da, i, cfg["khung"], _CFG.get("cty", "(chưa khai báo)"))
        except Exception: traceback.print_exc()
        finally: HANG.task_done()
def khoi_dong(data, lay_du_an, cty):
    """Gọi 1 lần khi engine chạy: bật luồng đọc AI + xếp lại việc dở (CHO_AI) của mọi dự án."""
    _CFG.update(du_an=lay_du_an, cty=cty); threading.Thread(target=_tho, daemon=True).start()
    for da in lay_du_an():
        for i, r in doc_so(data, da).items():
            if r.get("trang_thai") == "CHO_AI": HANG.put((data, da, i))
def doc_lai(data, da, i):
    r = sua_rec(data, da, i, trang_thai="CHO_AI", loi=None); HANG.put((data, da, i)); return r

# ───────── HSTT PDF = bản ký đi kèm HSTT Excel cùng HĐ + đợt ─────────
def _chuan(s): return re.sub(r"\s+", " ", re.sub(r"[^0-9a-zà-ỹđ ]", " ", str(s or "").lower())).strip()
def doan_hstt(ten_pdf, ds_hs):
    """Đoán HSTT Excel khớp với PDF: cùng số đợt trong tên file + nhiều từ tên đơn vị trùng nhất. Trả [(điểm, id)] giảm dần."""
    t = _chuan(os.path.splitext(ten_pdf)[0]); dot_pdf = re.findall(r"(?:đợt|dot|đ)\s*0?(\d{1,2})", t)
    kq = []
    for r in ds_hs:
        tt = r.get("tom_tat") or {}; ten_x = _chuan(os.path.splitext(r["ten"])[0]) + " " + _chuan(tt.get("don_vi"))
        diem = len(set(w for w in t.split() if len(w) > 2) & set(ten_x.split()))
        if dot_pdf and str(tt.get("dot")) in dot_pdf: diem += 5
        if diem: kq.append((diem, r["id"]))
    return sorted(kq, reverse=True)

def dinh_kem(data, da, rec, ten, raw):
    """Gắn PDF vào hồ sơ HSTT Excel. Đã ghi sổ ⇒ chép luôn vào folder đợt; chưa ⇒ giữ ở _HE_THONG/nap, sẽ xếp khi ghi sổ."""
    ten = os.path.basename(ten)
    if not ten.lower().endswith(".pdf"): raise ValueError("Bản ký HSTT phải là PDF")
    vt = van_tay(raw); dk = rec.setdefault("dinh_kem", [])
    if any(x["van_tay"] == vt for x in dk): return rec
    thu_muc = os.path.dirname(rec["luu_tru"]) if rec.get("luu_tru") else os.path.join(data, da, "_HE_THONG", "nap")
    dich = luu_file(data, da, os.path.relpath(thu_muc, os.path.join(data, da)), (ten if rec.get("luu_tru") else f"{vt[:12]}_{ten}"), raw)
    dk.append(dict(ten=ten, van_tay=vt, file=dich, da_xep=bool(rec.get("luu_tru")), luc=dt.datetime.now().isoformat(timespec="seconds")))
    return rec

def xep_dinh_kem(rec):
    """Gọi sau khi HSTT Excel ghi sổ + đã xếp: chép các PDF đính kèm vào cùng folder đợt."""
    if not rec.get("luu_tru"): return
    d = os.path.dirname(rec["luu_tru"])
    for x in rec.get("dinh_kem", []):
        if x.get("da_xep") or not os.path.exists(x["file"]): continue
        dich = os.path.join(d, x["ten"])
        if not os.path.exists(dich): shutil.copy2(x["file"], dich); os.chmod(dich, 0o444)
        x["file"], x["da_xep"] = dich, True
