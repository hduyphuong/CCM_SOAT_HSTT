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
    "BAO_GIA_CDT": ("Báo giá / đề xuất giá gửi CĐT", "10_THU_CDT/BAO_GIA"),
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

TIEN_TO = r"^(to doi|doi thi cong|thi cong|doi|dich vu|thuong mai dich vu|tu van|thiet ke|cong ty co phan|cong ty tnhh mtv|cong ty tnhh|cong ty cp|cong ty|ctcp|cty|tnhh|mtv|sx|tm|dv|xd|thuong mai|xay dung|kien truc|san xuat|dau tu)\s+"
def ma_de_xuat(ten):
    """Quy ước mã cũ: người/tổ đội ⇒ chữ đầu các từ + từ cuối ('Tổ đội Trần Thành Thắng' ⇒ TTThang);
    công ty ⇒ ghép 2 từ cuối ('Công ty TNHH TM DV XD Mỹ Kim' ⇒ MyKim); có viết tắt trong ngoặc ⇒ dùng luôn ('…(DECOFI)' ⇒ DECOFI)."""
    m = re.search(r"[\([\[]([A-Z0-9&]{3,12})[\)\]]", khong_dau(ten))
    if m: return m.group(1).replace("&", "")
    t = re.sub(r"[^a-z0-9]+", " ", re.sub(r"[\(\[].*?[\)\]]", " ", khong_dau(ten).lower())).strip(); cong_ty = bool(re.match(r"(cong ty|ctcp|cty|doanh nghiep)", t))
    for _ in range(6): t = re.sub(TIEN_TO, "", t.strip())
    if not cong_ty: t = re.sub(r"^(ong|ba|anh|chi)\s+", "", t)
    w = [x for x in re.split(r"[^a-z0-9]+", t) if x]
    if not w: return "DT_MOI"
    if cong_ty and len(w) <= 2: return "".join(x.capitalize() for x in w)[:16]          # Mỹ Kim ⇒ MyKim · Ong Vàng ⇒ OngVang
    return ("".join(x[0].upper() for x in w[:-1]) + w[-1].capitalize())[:16]
def loi_ten(ten):
    """Phần tên riêng (bỏ dấu, bỏ tiền tố tổ đội / công ty / loại hình): 'Tổ đội thi công Nguyễn Công Danh' ⇒ ['nguyen','cong','danh']."""
    t = re.sub(r"[^a-z0-9]+", " ", re.sub(r"[\(\[].*?[\)\]]", " ", khong_dau(ten).lower())).strip()
    for _ in range(6): t = re.sub(TIEN_TO, "", t.strip())
    return [w for w in re.split(r"[^a-z0-9]+", t) if w]
def khop_doi_tac(ten, ds):
    """Đối tác đã có trong khung: TRÙNG MÃ đề xuất, hoặc phần tên riêng TRÙNG HẲN (cùng tên cuối, cùng số chữ). Không khớp mờ."""
    k = ma_de_xuat(ten).lower(); lt = loi_ten(ten)
    for d in ds:
        if d["ma"].lower() == k: return d
        ld = loi_ten(d["ten"])
        if lt and ld and lt == ld: return d
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
    if loai not in ("HD_CDT", "HD_DOI_TAC"):                          # % thanh toán chỉ có nghĩa với hợp đồng ⇒ bỏ cờ nhiễu
        co[:] = [c for c in co if not re.match(r"^(vat_pct|pct_\w+): AI trả", c["mo_ta"])]
    if loai in ("HD_CDT", "HD_DOI_TAC"):
        for f, t in (("so_hd", "số hợp đồng"), ("ngay_ky", "ngày ký"), ("doi_tac_ten", "tên đối tác")):
            if not kq.get(f): C("CHAN", f"Không đọc được {t} — anh kiểm lại PDF")
        if kq.get("dang_hd") != "NGUYEN_TAC" and not kq.get("gia_tri_truoc_vat"):
            if any(isinstance(x.get("don_gia"), (int, float)) for x in la): C("LUU_Y", "HĐ không có giá trị cố định (HĐ đơn giá / nguyên tắc) — ghi khung với giá trị 0, thanh toán theo đơn giá")
            else: C("CHAN", "Không đọc được giá trị HĐ trước VAT")
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
        if t and la and abs(tong - t) > t * 0.005: C("CHAN", f"Σ các dòng {tong:,.0f} ≠ tổng ghi trên file {t:,.0f} (lệch {abs(tong - t) / t:.0%}) — file nhiều sheet, AI có thể gom lẫn dòng; cần đọc đúng sheet")
        elif t and la and abs(tong - t) > 1000: C("LUU_Y", f"Σ các dòng {tong:,.0f} ≠ tổng ghi trên file {t:,.0f} (làm tròn)")
    for x in kq.get("khong_chac") or []: C("LUU_Y", f"AI đọc không chắc: {x}")
    if kq.get("loai") == "HSTT": C("CHAN", "NẠP SAI CHỖ — đây là HỒ SƠ THANH TOÁN: kéo file vào mục ① Nạp & duyệt (Excel để soát; PDF scan bấm '🤖 AI đọc bản scan')")
    return co, dict(so_dong=len(la), tong_dong=tong)

PCT = ("vat_pct", "pct_tam_ung", "pct_tt_dot", "pct_tt_quyet_toan", "pct_giu_lai")
def tu_khoa_cty(cty):
    """'VELA (Công ty CP Kỹ thuật Xây dựng VELA)' ⇒ ['VELA'] — từ khoá nhận ra công ty người dùng trong HĐ."""
    goc = str(cty or "").split("(")[0].strip()
    return [khong_dau(goc).upper()] if goc and not goc.startswith("(chưa") else []
def la_cty_minh(ten, kw): t = khong_dau(ten).upper(); return any(k and re.search(r"\b" + re.escape(k) + r"\b", t) for k in kw)
TIEN_TO_DOI = re.compile(r"^\s*(?:tổ\s+đội(?:\s+thi\s+công)?|đội\s+thi\s+công|đội|tổ)\s+", re.I)
def ten_chuan(s):
    """Tên đối tác CHUẨN (anh chốt 28/09): bỏ tiền tố 'Tổ đội / Đội thi công…' (chỉ giữ tên người) · VIẾT HOA toàn bộ cho đồng bộ."""
    t = re.sub(r"\s{2,}", " ", str(s or "")).strip(" -–,;")
    for _ in range(2): t = TIEN_TO_DOI.sub("", t)
    return t.upper()

def chuan_hoa(kq, cty=None):
    """KHÔNG tin AI về định dạng / phân loại: % dạng 90 ⇒ 0.9 (ngoài 0..1 ⇒ CHẶN); công ty mình là BÊN NHẬN ⇒ HĐ phía CĐT, đối tác = bên giao;
    đối tác trùng công ty mình ⇒ CHẶN; tổ đội / giao khoán ⇒ DTC. Trả danh sách cờ."""
    co = []; kw = tu_khoa_cty(cty if cty is not None else _CFG.get("cty"))
    if kw and kq.get("loai") in ("HD_CDT", "HD_DOI_TAC", "PLHD_CDT", "PLHD_DOI_TAC"):
        chu = khong_dau(" ".join(str(kq.get(x) or "") for x in ("noi_dung", "so_hd", "doi_tac_ten"))).lower()
        if kq.get("ben_tra_tien") or kq.get("ben_nhan_tien"):                        # ƯU TIÊN dòng tiền: mình TRẢ ⇒ chi phí, mình ĐƯỢC TRẢ ⇒ doanh thu
            nhan, giao = la_cty_minh(kq.get("ben_nhan_tien"), kw), la_cty_minh(kq.get("ben_tra_tien"), kw)
            ben_kia = kq.get("ben_tra_tien") if nhan else kq.get("ben_nhan_tien")
        elif kq.get("loai_doi_tac") == "NCC" or re.search(r"mua ban|cung cap|cung ung|hang hoa|vat tu|hdmb|hdnt", chu):
            m1, m2 = la_cty_minh(kq.get("ben_nhan"), kw), la_cty_minh(kq.get("ben_giao"), kw)   # mua bán / NCC: mình là BÊN MUA ⇒ trả tiền
            nhan, giao = False, (m1 or m2); ben_kia = kq.get("ben_giao") if m1 else kq.get("ben_nhan")
        else:
            nhan, giao = la_cty_minh(kq.get("ben_nhan"), kw), la_cty_minh(kq.get("ben_giao"), kw)
            ben_kia = kq.get("ben_giao") if nhan else kq.get("ben_nhan")
        if nhan and not giao and kq["loai"] in ("HD_DOI_TAC", "PLHD_DOI_TAC"):
            kq["loai"] = kq["loai"].replace("DOI_TAC", "CDT"); co.append(dict(muc="LUU_Y", mo_ta="Công ty mình là BÊN NHẬN thầu ⇒ app đổi thành HĐ phía CĐT (doanh thu)"))
        if nhan and not giao: kq["doi_tac_ten"], kq["loai_doi_tac"] = ben_kia, "CDT"
        if giao and not nhan and kq["loai"] in ("HD_CDT", "PLHD_CDT"):
            kq["loai"] = kq["loai"].replace("CDT", "DOI_TAC"); co.append(dict(muc="LUU_Y", mo_ta="Công ty mình là BÊN GIAO việc ⇒ app đổi thành HĐ đối tác (chi phí)"))
        if giao and not nhan:
            kq["doi_tac_ten"] = ben_kia
            if kq.get("loai_doi_tac") == "CDT": kq["loai_doi_tac"] = "NCC" if re.search(r"mua ban|cung cap|cung ung|hang hoa|vat tu", chu) else None
    if kq.get("doi_tac_ten"):                                         # chỉ giữ TÊN: bỏ mọi (…), CCCD/CMND/MST
        kq["doi_tac_ten"] = re.sub(r"\s{2,}", " ", re.sub(r"(?i)\b(CCCD|CMND|MST|Mã số thuế)\b\s*[:.]?\s*[\d\s.-]*", " ", re.sub(r"\([^)]*\)", " ", kq["doi_tac_ten"]))).strip(" -–,;")
    if kq.get("doi_tac_ten"):                                         # bỏ phần người đại diện: "… - Ông Chu Quang Huân, P.TGĐ (ủy quyền…)"
        kq["doi_tac_ten"] = re.split(r"\s*[-,–]\s*(?:Ông|Bà|Ong|Ba|Giám đốc|Giam doc|Tổng giám đốc|Đại diện|Dai dien|Người đại diện)\b|\s*[\(\[]\s*(?:đại diện|dai dien|ông|bà)|\s*[-,–(]?\s*(?:MST|Mã số thuế)\b|\s*[-,–(:]?\s*(?:S[ốo]\s*)?(?:STK|TK|T[àa]i kho[ảa]n)\b\s*[:.]?\s*\d", kq["doi_tac_ten"], flags=re.I)[0].strip(" -,–")
        kq["doi_tac_ten"] = ten_chuan(kq["doi_tac_ten"])
        if la_cty_minh(kq.get("doi_tac_ten"), kw): co.append(dict(muc="CHAN", mo_ta="Đối tác trùng tên công ty mình — không xác định được bên nào là đối tác"))
    for k in PCT:
        v = kq.get(k)
        if isinstance(v, (int, float)) and 1.0001 < v <= 100: kq[k] = round(v / 100, 6); co.append(dict(muc="LUU_Y", mo_ta=f"{k}: AI trả {v} ⇒ app đổi thành {kq[k]:.2%}"))
        elif isinstance(v, (int, float)) and not (0 <= kq[k] <= 1): co.append(dict(muc="CHAN", mo_ta=f"{k} = {v} không hợp lệ (phải 0–100%)"))
    chu = khong_dau(" ".join(str(kq.get(x) or "") for x in ("doi_tac_ten", "ben_nhan", "so_hd", "noi_dung"))).lower()
    cong_ty = bool(re.search(r"cong ty|ctcp|tnhh|co phan", khong_dau(kq.get("doi_tac_ten") or kq.get("ben_nhan") or "").lower()))
    if cong_ty and kq.get("loai_doi_tac") == "DTC": kq["loai_doi_tac"] = "NTP"; co.append(dict(muc="LUU_Y", mo_ta="Đối tác là công ty ⇒ Thầu phụ (NTP), không phải tổ đội"))
    if not cong_ty and kq.get("loai_doi_tac") in ("NTP", None) and re.search(r"\bto doi\b|doi thi cong|giao khoan|khoan nhan cong|hdgk", chu):
        co.append(dict(muc="LUU_Y", mo_ta=f"AI xếp loại đối tác {kq.get('loai_doi_tac')} nhưng HĐ là giao khoán / tổ đội ⇒ app đổi thành Đội thi công (DTC)")); kq["loai_doi_tac"] = "DTC"
    return co
def xu_ly_ai(data, da, i, khung, cty):
    import ai_doc
    s = doc_so(data, da); rec = s[i]; f = os.path.join(data, rec["duong_dan"])
    import ns_r00
    if f.lower().endswith((".xlsx", ".xlsm")) and ns_r00.la_mau(f):     # mẫu ngân sách nội bộ ⇒ bộ đọc TẤT ĐỊNH, không AI
        try: kq, _ = ns_r00.ket_qua_ai_gia(f); ap_ket_qua(data, da, i, khung, kq, dict(giay=0, luot=0, bo_doc="NS mẫu nội bộ — không dùng AI")); return
        except Exception as e: sua_rec(data, da, i, trang_thai="LOI_AI", loi=f"Bộ đọc R00: {e}"[:400]); return
    try: kq, meta = ai_doc.doc(f, cty=cty)
    except Exception as e:
        sua_rec(data, da, i, trang_thai="LOI_AI", loi=str(e)[:400]); return
    if kq.get("doc_duoc") is False:                                          # AI không mở được file (lỗi Read/Drive chưa đồng bộ) — KHÔNG được coi là đã phân loại xong
        sua_rec(data, da, i, trang_thai="LOI_AI", loi=("AI không đọc được nội dung file: " + (kq.get("ly_do_loai") or "không rõ lý do"))[:400]); return
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
    kw = tu_khoa_cty(_CFG.get("cty"))
    if loai == "BAO_GIA" and (la_cty_minh(kq.get("doi_tac_ten"), kw) or la_cty_minh(kq.get("ben_giao"), kw)):   # báo giá CÔNG TY MÌNH lập ⇒ gửi CĐT (doanh thu)
        loai, ma_dt, loai_dt = "BAO_GIA_CDT", None, None
        co0 = [c for c in co0 if "trùng tên công ty mình" not in c["mo_ta"]] + [dict(muc="LUU_Y", mo_ta="Báo giá do công ty mình lập ⇒ xếp là báo giá / đề xuất giá gửi CĐT")]
    if loai == "BAO_GIA" and not (ma_dt and LOAI_DT.get(loai_dt or "")): loai = "CHON_THAU"
    co, tk = danh_gia(kq, loai, rec.get("loai_chon"), dm); co = co0 + co
    if loai not in ("HD_CDT", "HD_DOI_TAC"): co = [c for c in co if not re.match(r"^(vat_pct|pct_\w+): AI trả", c["mo_ta"])]   # % chỉ có nghĩa với HĐ
    moi_dt = bool(ma_dt) and not any(d["ma"] == ma_dt for d in dm["doi_tac"])
    if moi_dt and loai in ("HD_DOI_TAC", "BAO_GIA", "QUYET_TOAN"): co.append(dict(muc="LUU_Y", mo_ta=f"Đối tác mới (chưa có trong khung) — app đề xuất mã {ma_dt}, loại {loai_dt}"))
    la_pl = str(kq.get("loai") or "").startswith("PLHD") or bool(re.search(r"phu luc|plhd|^pl[ _.-]?\d", khong_dau(rec["ten"]).lower()))
    if la_pl and loai in ("HD_CDT", "HD_DOI_TAC"):                  # PHỤ LỤC chỉ duyệt SAU HĐ gốc — không bao giờ ghi thành HĐ gốc
        import phu_luc as PL                                             # PL = điều chỉnh / bổ sung bảng giá HĐ theo thứ tự (anh 30/09)
        co = [c for c in co if not c["mo_ta"].startswith(PL_BO)]
        co = PL.co_phu_luc(khung, dict(rec, ai=kq, loai=loai, ma_doi_tac=ma_dt)) + co
    rel = thu_muc_dich(loai, ma_dt, loai_dt, rec.get("goi")) or "_HE_THONG/cho_phan_loai"
    with KHOA:
        s = doc_so(data, da); rec = s[i]; cu = os.path.join(data, rec["duong_dan"])
        if not rel.startswith("_HE_THONG"):
            dich = luu_file(data, da, rel, rec["ten"], open(cu, "rb").read())
            if os.path.normcase(dich) != os.path.normcase(cu) and "cho_phan_loai" in cu:
                os.chmod(cu, stat.S_IWRITE); os.remove(cu)
            rec["duong_dan"] = os.path.relpath(dich, data)
        rec.update(ai=kq, ai_meta=meta, co=co, la_phu_luc=la_pl, thong_ke=tk, loai=loai, loai_ten=LOAI[loai][0], ma_doi_tac=ma_dt, loai_doi_tac=loai_dt, doi_tac_moi=moi_dt,
                   can_nhap=loai in CAN_NHAP, trang_thai="DA_NHAP" if rec.get("da_nhap_khung") else ("SAI_NOI" if kq.get("loai") == "HSTT" else ("CHO_DUYET" if loai in CAN_NHAP else "DA_LUU")), loi=None, luc_ai=dt.datetime.now().isoformat(timespec="seconds"))
        s[i] = rec; ghi_so_nen(data, da, s)

PL_BO = ("Không đọc được giá trị HĐ", "Không đọc được số hợp đồng", "Σ bảng đơn giá", "Phụ lục HĐ — chưa có HĐ gốc", "Phụ lục của ")   # cờ HĐ gốc không áp cho PL
def soat_lai_pl(data, da, khung):
    """Phụ lục đang chờ duyệt ⇒ đánh giá lại theo bảng giá đang hiệu lực mỗi khi khung đổi (duyệt PL01 xong thì PL02 tự hết chặn, không phải AI đọc lại)."""
    import phu_luc as PL
    mt = os.path.getmtime(khung); s = doc_so(data, da)
    can = [i for i, r in s.items() if r.get("la_phu_luc") and r.get("trang_thai") == "CHO_DUYET" and r.get("loai") in ("HD_CDT", "HD_DOI_TAC") and r.get("pl_mtime") != mt]
    for i in can:
        try: moi = PL.co_phu_luc(khung, s[i])
        except Exception as e: moi = [dict(muc="CHAN", mo_ta=f"Không đối chiếu được phụ lục với khung: {e}"[:300], pl=True)]
        with KHOA:
            s2 = doc_so(data, da); r = s2[i]
            r["co"] = moi + [c for c in r.get("co", []) if not c.get("pl") and not c["mo_ta"].startswith(PL_BO)]; r["pl_mtime"] = mt
            ghi_so_nen(data, da, s2)
    return len(can)

# ── DUYỆT TỪNG CỜ (anh chốt 30/09): anh xác nhận ✓ OK từng cờ, hoặc hỏi Agent kiểm lại ĐÚNG cờ đó — không phải đọc lại cả file ──
def khoa_co(c): return f"{c['muc']}|{c['mo_ta']}"
def chan_con(rec):
    """Cờ CHẶN chưa được anh xác nhận ⇒ chưa ghi khung được. Cờ CHẶN số học của phụ lục (pl) chỉ hết bằng dữ liệu, không xác nhận tay."""
    ok = rec.get("co_ok") or {}
    return [c for c in rec.get("co", []) if c["muc"] == "CHAN" and (c.get("pl") or khoa_co(c) not in ok)]
def xac_nhan_co(data, da, i, khoa, ly_do="", bo=False):
    with KHOA:
        s = doc_so(data, da); r = s[i]; co = next((c for c in r.get("co", []) if khoa_co(c) == khoa), None)
        if co is None: raise ValueError("Cờ này không còn (hồ sơ vừa được soát lại) — anh tải lại trang")
        ok = r.setdefault("co_ok", {})
        if bo: ok.pop(khoa, None)
        else:
            if co["muc"] == "CHAN" and co.get("pl"): raise ValueError("Cờ CHẶN này là phép đối chiếu số của phụ lục — chỉ hết khi số liệu khớp (duyệt PL trước theo thứ tự, hoặc hỏi Agent sửa bảng)")
            if co["muc"] == "CHAN" and not (ly_do or "").strip(): raise ValueError("Cờ CHẶN cần anh ghi lý do chấp nhận")
            ok[khoa] = dict(ly_do=(ly_do or "").strip()[:500], luc=dt.datetime.now().isoformat(timespec="seconds"))
        ghi_so_nen(data, da, s); return r

SCHEMA_HOI = {"type": "object", "properties": {
    "tra_loi": {"type": "string", "description": "Kết luận ngắn, tiếng Việt, có dẫn trang/dòng làm bằng chứng"},
    "ket_luan": {"type": "string", "enum": ["SO_LIEU_DUNG", "DA_SUA", "KHONG_XAC_DINH"]},
    "sua_truong": {"type": "object", "description": "Trường đầu hồ sơ cần sửa: {tên trường: giá trị đúng} — chỉ các trường đã có trong kết quả đọc"},
    "sua_dong": {"type": "array", "items": {"type": "object", "properties": {
        "stt": {"type": "string"}, "noi_dung": {"type": "string"}, "dvt": {"type": "string"}, "kl": {"type": ["number", "null"]},
        "don_gia": {"type": ["number", "null"]}, "thanh_tien": {"type": ["number", "null"]}, "xoa": {"type": "boolean"}}, "required": ["stt"]},
        "description": "Dòng bảng cần sửa (khớp theo STT) — chỉ nêu trường cần đổi; xoa=true để bỏ dòng đọc nhầm"},
    "them_dong": {"type": "array", "items": {"type": "object", "properties": {
        "stt": {"type": "string"}, "noi_dung": {"type": "string"}, "dvt": {"type": "string"}, "kl": {"type": ["number", "null"]},
        "don_gia": {"type": ["number", "null"]}, "thanh_tien": {"type": ["number", "null"]}}, "required": ["stt", "noi_dung"]}}},
    "required": ["tra_loi", "ket_luan"]}
HUONG_DAN_HOI = ("Bạn là trợ lý soát hồ sơ hợp đồng / thanh toán xây dựng của {cty}. Nhiệm vụ HẸP: kiểm lại ĐÚNG MỘT cờ đối chiếu trên file hs.pdf / hs.txt "
    "theo câu hỏi của người phụ trách (ghi trong ngu_canh.txt, kèm kết quả đọc lần trước). Chỉ đọc các trang liên quan tới cờ (dùng tham số pages), "
    "không đọc lại toàn bộ nếu không cần. Chỉ SỬA khi thấy RÕ trên file — không đoán, không làm tròn lại số in trên file. "
    "Trả lời tiếng Việt, ngắn gọn, luôn dẫn trang / dòng làm bằng chứng. Không có gì cần sửa ⇒ ket_luan SO_LIEU_DUNG và giải thích vì sao cờ hiện lên.")
TRUONG_SUA = ("so_hd", "so_hd_goc", "ngay_ky", "doi_tac_ten", "ben_giao", "ben_nhan", "noi_dung", "dang_hd", "gia_tri_truoc_vat", "vat_pct", "gia_tri_sau_vat",
              "pct_tam_ung", "pct_tt_dot", "pct_tt_quyet_toan", "pct_giu_lai", "han_tt_ngay", "don_vi_han", "han_qt_ngay", "han_tra_gl_ngay", "bao_hanh_thang",
              "co_chu_ky", "co_dong_dau", "tong_ghi_tren_file")
def hoi_co(data, da, i, khoa, ghi_chu):
    """Xếp 1 câu hỏi Agent cho đúng 1 cờ — chạy nền, kết quả ghi vào rec['co_hoi'][khoa]."""
    with KHOA:
        s = doc_so(data, da); r = s[i]
        if not any(khoa_co(c) == khoa for c in r.get("co", [])): raise ValueError("Cờ này không còn — anh tải lại trang")
        if any(h.get("trang_thai") == "DANG" for h in (r.get("co_hoi") or {}).values()): raise ValueError("Agent đang kiểm 1 cờ khác của hồ sơ này — chờ xong rồi hỏi tiếp")
        r.setdefault("co_hoi", {})[khoa] = dict(trang_thai="DANG", ghi_chu=(ghi_chu or "").strip()[:1000], luc=dt.datetime.now().isoformat(timespec="seconds"))
        ghi_so_nen(data, da, s)
    threading.Thread(target=_chay_hoi, args=(data, da, i, khoa, ghi_chu), daemon=True).start()
    return dict(ok=True)
def _chay_hoi(data, da, i, khoa, ghi_chu):
    import ai_doc
    try:
        r = doc_so(data, da)[i]; f = os.path.join(data, r["duong_dan"]); a = r.get("ai") or {}
        dau = {k: v for k, v in a.items() if k not in ("bang", "nguon", "khong_chac")}
        ngu = (f"CỜ CẦN KIỂM: [{khoa.split('|', 1)[0]}] {khoa.split('|', 1)[1]}\n\nCÂU HỎI / GHI CHÚ CỦA NGƯỜI PHỤ TRÁCH: {ghi_chu or '(không ghi — kiểm xem cờ đúng hay sai)'}\n\n"
               f"KẾT QUẢ ĐỌC LẦN TRƯỚC — đầu hồ sơ:\n{json.dumps(dau, ensure_ascii=False, default=str)}\n\nBẢNG (mỗi dòng: stt | nội dung | ĐVT | KL | ĐG | thành tiền):\n"
               + "\n".join(f"{x.get('stt')} | {x.get('noi_dung')} | {x.get('dvt')} | {x.get('kl')} | {x.get('don_gia')} | {x.get('thanh_tien')}" for x in a.get("bang") or []))
        kq, meta = ai_doc.doc(f, cty=_CFG.get("cty", "(chưa khai báo)"), schema=SCHEMA_HOI, huong_dan=HUONG_DAN_HOI, timeout=900, ngu_canh=ngu)
        a2 = json.loads(json.dumps(a)); sua = []
        for k, v in (kq.get("sua_truong") or {}).items():
            if k in TRUONG_SUA and a2.get(k) != v: sua.append(f"{k}: {a2.get(k)} → {v}"); a2[k] = v
        bang = a2.setdefault("bang", [])
        for d in kq.get("sua_dong") or []:
            x = next((x for x in bang if str(x.get("stt")) == str(d["stt"]) and (x.get("dvt") or x.get("kl") is not None)), None)
            if x is None: continue
            if d.get("xoa"): bang.remove(x); sua.append(f"bỏ dòng {d['stt']}"); continue
            for k in ("noi_dung", "dvt", "kl", "don_gia", "thanh_tien"):
                if k in d and d[k] != x.get(k): sua.append(f"dòng {d['stt']} {k}: {x.get(k)} → {d[k]}"); x[k] = d[k]
        for d in kq.get("them_dong") or []:
            bang.append({k: d.get(k) for k in ("stt", "noi_dung", "dvt", "kl", "don_gia", "thanh_tien")}); sua.append(f"thêm dòng {d['stt']}")
        mt = khoa.split("|", 1)[1]
        if sua and mt.startswith("AI đọc không chắc: "):                     # Agent đã sửa đúng ý 'không chắc' được hỏi ⇒ ý đó hết
            y = mt[len("AI đọc không chắc: "):].strip(); kc = [x for x in a2.get("khong_chac") or [] if str(x).strip() != y]
            if len(kc) != len(a2.get("khong_chac") or []): a2["khong_chac"] = kc; sua.append("đã xử lý ý 'không chắc' này")
        h = dict(trang_thai="XONG", ghi_chu=ghi_chu, tra_loi=kq.get("tra_loi"), ket_luan=kq.get("ket_luan"), sua=sua, giay=meta.get("giay"), luc=dt.datetime.now().isoformat(timespec="seconds"))
        if sua:                                                              # có sửa ⇒ áp lại + soát lại toàn bộ cờ (không đọc lại file)
            ap_ket_qua(data, da, i, _CFG["du_an"]()[da]["khung"], a2, r.get("ai_meta") or {})
        with KHOA:
            s = doc_so(data, da); s[i].setdefault("co_hoi", {})[khoa] = h
            if sua: s[i].setdefault("lich_su_sua", []).append(dict(luc=h["luc"], co=khoa, sua=sua))
            ghi_so_nen(data, da, s)
    except Exception as e:
        traceback.print_exc()
        with KHOA:
            s = doc_so(data, da); s[i].setdefault("co_hoi", {})[khoa] = dict(trang_thai="LOI", ghi_chu=ghi_chu, tra_loi=str(e)[:400], luc=dt.datetime.now().isoformat(timespec="seconds"))
            ghi_so_nen(data, da, s)
def don_hoi_treo(data, da, phut=20):
    """Engine khởi động lại giữa chừng ⇒ câu hỏi DANG quá hạn chuyển LOI để anh hỏi lại."""
    han = dt.datetime.now() - dt.timedelta(minutes=phut)
    treo = lambda s: [h for r in s.values() for h in (r.get("co_hoi") or {}).values() if h.get("trang_thai") == "DANG" and dt.datetime.fromisoformat(h["luc"]) < han]
    if not treo(doc_so(data, da)): return
    with KHOA:
        s = doc_so(data, da)
        for h in treo(s): h.update(trang_thai="LOI", tra_loi="Bị ngắt (engine khởi động lại / quá giờ) — anh hỏi lại")
        ghi_so_nen(data, da, s)

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
    # BẮT BUỘC trùng TÊN RIÊNG đơn vị (≥ nửa) — chữ chung ('đợt', 'ứng', 'thanh toán'…) và số đợt KHÔNG đủ để khớp
    # (sự cố 26/09: PDF 'Nguyễn Đức Thống đợt 2 tạm ứng' bị gợi ý vào 'Hoàn ứng BCH đợt 2'). PDF ghi đợt khác ⇒ liệt kê nhưng KHÔNG chọn sẵn.
    t = _chuan(os.path.splitext(ten_pdf)[0]); dot_pdf = re.findall(r"(?:đợt|dot|đ)\s*0?(\d{1,2})", t)
    kd = khong_dau(t).lower(); w_pdf = set(re.split(r"[^a-z0-9]+", kd))
    kq = []
    for r in ds_hs:
        tt = r.get("tom_tat") or {}; ma = (r.get("phan_loai") or {}).get("ma_hd") or ""
        if ma == "HD-BCH": trung = 1 if ("hoan ung" in kd or "bch" in w_pdf) else 0
        else:
            if ma == "HD-CDT": rieng = {w for w in re.split(r"[^a-z0-9]+", khong_dau(os.path.splitext(r["ten"])[0]).lower()) if len(w) >= 3 and not w.isdigit()} - CHUNG_PDF
            else: rieng = set(loi_ten(tt.get("don_vi") or "")) - CHUNG_PDF
            n = len(rieng & w_pdf); trung = n if rieng and n >= len(rieng) - len(rieng) // 3 else 0
        if not trung: continue
        dung_dot = not dot_pdf or str(tt.get("dot")) in dot_pdf
        kq.append(((trung * 2 + 5) if dung_dot else 0, r["id"]))
    return sorted(kq, reverse=True)
CHUNG_PDF = {"dot", "hstt", "ho", "so", "thanh", "toan", "thanhtoan", "gia", "tri", "giatri", "tam", "ung", "hoan", "ky", "ban", "pdf", "scan",
             "cdt", "gui", "vela", "hang", "hai", "doi", "thi", "cong", "nha", "cung", "cap", "de", "nghi", "bang", "khoi", "luong", "value"}

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
