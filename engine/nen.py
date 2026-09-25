"""HỒ SƠ NỀN — nạp BoQ, ngân sách, gói thầu, báo giá, hợp đồng, quyết toán (Excel/PDF) + HSTT PDF (bản ký kèm HSTT Excel).
Kho PDF thực tế là BẢN SCAN (0 ký tự) ⇒ KHÔNG tự đọc số: file được LƯU đúng folder làm chứng từ; số liệu HĐ nhập qua form (xem hd_form).
Sổ đăng ký: <DATA>\\<dự án>\\_HE_THONG\\ho_so_nen.json (đường dẫn tương đối). Trùng nội dung (SHA-256) ⇒ không lưu lần 2."""
import os, re, json, hashlib, datetime as dt, openpyxl, warnings
warnings.filterwarnings("ignore")
DUOI = (".pdf", ".xlsx", ".xlsm", ".xls", ".doc", ".docx", ".jpg", ".jpeg", ".png")
LOAI = {   # mã: (tên hiển thị, thư mục tương đối — {dt} = <LOAI>_<mã đối tác>, {goi} = mã gói)
    "BOQ_CDT":   ("BoQ / dự toán HĐ với CĐT", "01_THIET_LAP/01_BOQ_CDT"),
    "NGAN_SACH": ("Ngân sách (R00 / điều chỉnh)", "01_THIET_LAP/02_NGAN_SACH"),
    "GOI_THAU":  ("Phân chia gói thầu", "01_THIET_LAP/03_GOI_THAU"),
    "CHON_THAU": ("Báo giá chọn thầu (chưa có HĐ)", "01_THIET_LAP/03_GOI_THAU/CHON_THAU/{goi}"),
    "HD_CDT":    ("Hợp đồng / PLHĐ với CĐT", "10_THU_CDT/HOP_DONG"),
    "HD_DOI_TAC": ("Hợp đồng / PLHĐ đối tác", "20_CHI_DOI_TAC/{dt}/HOP_DONG"),
    "BAO_GIA":   ("Báo giá / bảng giá đối tác", "20_CHI_DOI_TAC/{dt}/BAO_GIA"),
    "QUYET_TOAN": ("Hồ sơ quyết toán đối tác", "20_CHI_DOI_TAC/{dt}/QUYET_TOAN"),
    "HSTT_KY":   ("HSTT bản ký (PDF) kèm HSTT Excel", "")}
LOAI_DT = {"ĐTC": "DTC", "DTC": "DTC", "NTP": "NTP", "NCC": "NCC", "DVK": "DVK"}

def _sach(s): return re.sub(r"[^A-Za-z0-9_.-]", "", str(s or "")) or "KHAC"
def van_tay(raw): return hashlib.sha256(raw).hexdigest()
def so_path(data, da): return os.path.join(data, da, "_HE_THONG", "ho_so_nen.json")
def doc_so(data, da):
    p = so_path(data, da); return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}
def ghi_so_nen(data, da, s):
    os.makedirs(os.path.dirname(so_path(data, da)), exist_ok=True)
    json.dump(s, open(so_path(data, da), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

def danh_muc(khung):
    """Danh sách chọn cho form: đối tác (N1), HĐ (N4/N6), gói (N3) — đọc từ file khung."""
    wb = openpyxl.load_workbook(khung, read_only=True, data_only=True)
    dt_ = [dict(ma=r[6], ten=r[7] or "", loai=r[8] or "") for r in wb["N1_DanhMuc"].iter_rows(min_row=3, max_col=9, values_only=True) if r[6]]
    hd = [dict(ma_hd=r[0], ben="CĐT", ma_doi_tac=r[2], so_hd=r[3]) for r in wb["N4_HD_CDT"].iter_rows(min_row=2, max_col=4, values_only=True) if r[0]]
    hd += [dict(ma_hd=r[0], ben="ĐỐI TÁC", ma_doi_tac=r[2], so_hd=r[4]) for r in wb["N6_HD_DoiTac"].iter_rows(min_row=2, max_col=5, values_only=True) if r[0]]
    goi = [dict(ma=r[0], ten=r[1] or "") for r in wb["N3_GoiThau"].iter_rows(min_row=2, max_col=2, values_only=True) if r[0]]
    wb.close()
    return dict(doi_tac=dt_, hop_dong=hd, goi=goi, loai=[dict(ma=k, ten=v[0]) for k, v in LOAI.items() if k != "HSTT_KY"])

def _thu_muc_dt(khung, ma, loai_nhap):
    """Folder đối tác: ưu tiên loại trong N1; đối tác chưa có trong khung ⇒ dùng loại anh chọn (DTC/NTP/NCC/DVK)."""
    loai = {d["ma"]: d["loai"] for d in danh_muc(khung)["doi_tac"]}.get(ma, loai_nhap)
    l = LOAI_DT.get(loai or "")
    if not l: raise ValueError("Chọn loại đối tác (Đội / Thầu phụ / NCC / Dịch vụ khác) để biết xếp vào folder nào")
    return f"{l}_{_sach(ma)}"

def luu_file(data, da, rel_dir, ten, raw):
    """Ghi file vào <DATA>/<da>/<rel_dir>/ (chỉ đọc). Trùng tên khác nội dung ⇒ thêm _v2, _v3…"""
    d = os.path.normpath(os.path.join(data, da, rel_dir)); os.makedirs(d, exist_ok=True)
    goc, duoi = os.path.splitext(ten); dich = os.path.join(d, ten); k = 1
    while os.path.exists(dich):
        if van_tay(open(dich, "rb").read()) == van_tay(raw): return dich
        k += 1; dich = os.path.join(d, f"{goc}_v{k}{duoi}")
    open(dich, "wb").write(raw); os.chmod(dich, 0o444); return dich

def nap_nen(data, da, khung, ten, raw, loai, ma_dt=None, loai_dt=None, goi=None, ghi_chu=""):
    ten = os.path.basename(ten)
    if not ten.lower().endswith(DUOI): raise ValueError(f"'{ten}': chỉ nhận {', '.join(DUOI)}")
    if loai not in LOAI or loai == "HSTT_KY": raise ValueError("Chọn loại hồ sơ nền")
    vt = van_tay(raw); s = doc_so(data, da)
    cu = next((v for v in s.values() if v["van_tay"] == vt), None)
    if cu: return dict(cu, trung=True)
    mau = LOAI[loai][1]
    if "{dt}" in mau:
        if not ma_dt: raise ValueError("Chọn đối tác (hoặc nhập mã đối tác mới)")
        mau = mau.replace("{dt}", _thu_muc_dt(khung, ma_dt, loai_dt))
    if "{goi}" in mau: mau = mau.replace("{goi}", _sach(goi) if goi else "CHUA_GAN_GOI")
    dich = luu_file(data, da, mau, ten, raw)
    rec = dict(id=vt[:12], van_tay=vt, ten=ten, loai=loai, loai_ten=LOAI[loai][0], ma_doi_tac=ma_dt, goi=goi, ghi_chu=ghi_chu,
               duong_dan=os.path.relpath(dich, data), luc=dt.datetime.now().isoformat(timespec="seconds"),
               da_nhap_khung=False, can_nhap=loai in ("HD_CDT", "HD_DOI_TAC", "BOQ_CDT", "NGAN_SACH", "GOI_THAU"))
    s[rec["id"]] = rec; ghi_so_nen(data, da, s); return rec

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
    import shutil
    if not rec.get("luu_tru"): return
    d = os.path.dirname(rec["luu_tru"])
    for x in rec.get("dinh_kem", []):
        if x.get("da_xep") or not os.path.exists(x["file"]): continue
        dich = os.path.join(d, x["ten"])
        if not os.path.exists(dich): shutil.copy2(x["file"], dich); os.chmod(dich, 0o444)
        x["file"], x["da_xep"] = dich, True
