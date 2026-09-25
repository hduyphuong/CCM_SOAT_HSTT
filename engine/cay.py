"""CÂY THƯ MỤC DỮ LIỆU của mỗi dự án (phương án B: tách THU / CHI, mỗi đối tác = 1 bộ hồ sơ) + tự xếp HSTT đã ghi sổ.
<DATA>\\
  _CAU_HINH\\            du_an.json · trang.txt
  _DUNG_CHUNG\\DOI_TAC\\  HĐ nguyên tắc / bảng giá dùng cho NHIỀU dự án
  <MA_DU_AN>\\
    00_KHUNG\\ (+_backup)  file khung Excel duy nhất của dự án
    01_THIET_LAP\\         01_BOQ_CDT · 02_NGAN_SACH · 03_GOI_THAU\\CHON_THAU
    10_THU_CDT\\           HOP_DONG · HSTT\\Dxx_YYYYMMDD
    20_CHI_DOI_TAC\\<LOAI>_<MA_DT>\\  HOP_DONG · BAO_GIA · HSTT\\Dxx_YYYYMMDD · QUYET_TOAN
    30_BAO_CAO\\<YYYY-MM-DD>
    _HE_THONG\\            nap\\ (bản gốc đã nạp, chỉ đọc) · so_nap.json — engine quản lý
Chỉ tạo folder đối tác khi đã có HĐ trong file khung (N6) — đúng quy tắc "chưa có HĐ không nhập"."""
import os, re, shutil, openpyxl, warnings
warnings.filterwarnings("ignore")
KHUNG_DU_AN = ["00_KHUNG/_backup", "01_THIET_LAP/01_BOQ_CDT", "01_THIET_LAP/02_NGAN_SACH", "01_THIET_LAP/03_GOI_THAU/CHON_THAU",
               "10_THU_CDT/HOP_DONG", "10_THU_CDT/HSTT", "20_CHI_DOI_TAC", "30_BAO_CAO", "_HE_THONG/nap"]
BO_DOI_TAC = ["HOP_DONG", "BAO_GIA", "HSTT", "QUYET_TOAN"]
LOAI = {"ĐTC": "DTC", "DTC": "DTC", "NTP": "NTP", "NCC": "NCC", "DVK": "DVK"}
_CACHE = {}

def _sach(s): return re.sub(r"[^A-Za-z0-9_.-]", "", str(s)) or "KHAC"

def ban_do_hd(khung):
    """ma_hd → (bên, thư mục tương đối). CĐT ⇒ 10_THU_CDT; đối tác ⇒ 20_CHI_DOI_TAC\\<LOAI>_<MA_DT>. Cache theo mtime."""
    m = os.path.getmtime(khung)
    if _CACHE.get(khung, (None,))[0] == m: return _CACHE[khung][1]
    wb = openpyxl.load_workbook(khung, read_only=True, data_only=True)
    loai = {r[6]: r[8] for r in wb["N1_DanhMuc"].iter_rows(min_row=3, max_col=9, values_only=True) if r[6]}
    bd = {r[0]: ("CĐT", "10_THU_CDT") for r in wb["N4_HD_CDT"].iter_rows(min_row=2, max_col=1, values_only=True) if r[0]}
    for r in wb["N6_HD_DoiTac"].iter_rows(min_row=2, max_col=3, values_only=True):
        if not r[0] or r[0] in bd: continue
        l = LOAI.get(loai.get(r[2], ""))
        if l: bd[r[0]] = ("ĐỐI TÁC", os.path.join("20_CHI_DOI_TAC", f"{l}_{_sach(r[2])}"))    # CĐT khấu trừ (loại CĐT) không lập folder đối tác
    wb.close(); _CACHE[khung] = (m, bd); return bd

def tao_cay(data, da, khung=None):
    """Tạo (bổ sung) cây của 1 dự án — chạy lại bao nhiêu lần cũng được, không xoá gì."""
    goc = os.path.join(data, da)
    for d in KHUNG_DU_AN: os.makedirs(os.path.normpath(os.path.join(goc, d)), exist_ok=True)
    if khung and os.path.exists(khung):
        for ben, rel in set(ban_do_hd(khung).values()):
            if ben == "ĐỐI TÁC":
                for d in BO_DOI_TAC: os.makedirs(os.path.join(goc, rel, d), exist_ok=True)
    os.makedirs(os.path.join(data, "_CAU_HINH"), exist_ok=True); os.makedirs(os.path.join(data, "_DUNG_CHUNG", "DOI_TAC"), exist_ok=True)
    return goc

def luu_tru(data, da, rec, khung):
    """HSTT vừa ghi sổ ⇒ chép bản gốc vào …\\HSTT\\Dxx_YYYYMMDD\\ của đúng HĐ (chỉ đọc). Trả đường dẫn, hoặc None nếu chưa xác định được HĐ."""
    ma = (rec.get("phan_loai") or {}).get("ma_hd"); t = rec.get("tom_tat") or {}
    rel = ban_do_hd(khung).get(ma, (None, None))[1]
    if not rel or not rec.get("file") or not os.path.exists(rec["file"]): return None
    dot = t.get("dot"); ngay = str(t.get("ngay") or "")[:10].replace("-", "")
    thu_muc = os.path.join(data, da, rel, "HSTT", f"D{int(dot):02d}_{ngay}" if isinstance(dot, (int, float)) else f"Dxx_{ngay}")
    if rel.startswith("20_"):
        for d in BO_DOI_TAC: os.makedirs(os.path.join(data, da, rel, d), exist_ok=True)
    os.makedirs(thu_muc, exist_ok=True)
    dich = os.path.join(thu_muc, rec["ten"])
    if not os.path.exists(dich): shutil.copy2(rec["file"], dich); os.chmod(dich, 0o444)
    return dich
