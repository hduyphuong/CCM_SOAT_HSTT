"""KHAI BÁO THỦ CÔNG — chi phí KHÔNG có hồ sơ (lương, thưởng, BHXH, công tác phí…) — anh Phương yêu cầu 29/09/2026.
Ghi vào HĐ nội bộ HD-NOIBO (tự tạo lần đầu) · mỗi mã NS = 1 dòng HĐ (nhóm NB_<mã NS>, ĐVT 'đ', ĐG 1) ⇒ chi phí lên đúng mã NS trên báo cáo / EAC.
Mỗi lần khai báo = 1 dòng N9 THUC_HIEN (KL = số tiền trước VAT) · có VAT ⇒ vat_rieng (chi phí trước VAT, tiền chi gồm VAT) · tự kiểm lũy kế tăng đúng số tiền."""
import os, datetime as dt, openpyxl, warnings
warnings.filterwarnings("ignore")
MA_HD, MA_DT, TEN_DT = "HD-NOIBO", "NOIBO", "CHI PHÍ NỘI BỘ – KHAI BÁO TAY"

def ds_ma_ns(khung):
    """[(mã NS, tên)] từ N1_DanhMuc cột L:M."""
    wb = openpyxl.load_workbook(khung, read_only=True, data_only=True)       # PHẢI đóng: read_only giữ file ⇒ Excel COM không lưu được (treo / sập 29/09)
    try: return [dict(ma=str(r[11]), ten=str(r[12] or "")) for r in wb["N1_DanhMuc"].iter_rows(min_row=3, max_row=200, values_only=True) if r[11]]
    finally: wb.close()

def _dam_bao_hd(khung, backup, da=""):
    """Chưa có HD-NOIBO ⇒ tạo (N6 + đối tác N1) qua nhap_khung.ghi_hd — không kèm bảng (dòng theo mã NS thêm khi ghi)."""
    import kiem as K, nhap_khung as NK
    if MA_HD in K.doc_khung(khung)["hd"]: return None
    a = dict(loai="HD_DOI_TAC", so_hd="HĐ TẠM – KHAI BÁO TAY", ngay_ky=dt.date.today().isoformat(), doi_tac_ten=TEN_DT, loai_doi_tac="DVK",
             noi_dung="HĐ TẠM tự khai báo — chi phí nội bộ không có hồ sơ (lương, thưởng, BHXH, công tác phí…) — khai báo tay trên webapp", dang_hd="NGUYEN_TAC",
             gia_tri_truoc_vat=None, vat_pct=0, pct_tam_ung=0, pct_tt_dot=1, pct_tt_quyet_toan=1, han_tt_ngay=0, don_vi_han="LICH", bang=[],
             nguon=f"webapp · tự tạo khi khai báo tay · {dt.date.today():%d/%m/%Y}", ghi_chu="HĐ nội bộ: mỗi mã NS 1 dòng, mỗi lần khai báo 1 dòng N9 (nội dung ở cột Nguồn)")
    kq = NK.ghi_hd(khung, da, dict(loai="HD_DOI_TAC", ma_doi_tac=MA_DT, loai_doi_tac="DVK", ten="khai báo tay", goi=None, ai=a), {}, backup)
    if not kq["ok"]: raise ValueError("Không tạo được HĐ nội bộ: " + kq["ly_do"])
    return kq

def ghi(khung, backup, ma_ns, noi_dung, ngay, so_tien, vat=0.0, ghi_chu="", nguoi="anh", da=""):
    """Ghi 1 khoản khai báo tay. Trả dict kết quả (ok, dot, backup…)."""
    import kiem as K, ghi_so as G
    so_tien = float(so_tien); vat = float(vat or 0)
    if so_tien == 0: raise ValueError("Số tiền phải khác 0")
    if not str(noi_dung or "").strip(): raise ValueError("Thiếu nội dung khoản chi")
    ns = {x["ma"]: x["ten"] for x in ds_ma_ns(khung)}
    if ma_ns not in ns: raise ValueError(f"Mã NS '{ma_ns}' không có trong ngân sách khung")
    try: ngay_d = dt.date.fromisoformat(str(ngay)[:10])
    except ValueError: raise ValueError("Ngày không hợp lệ (YYYY-MM-DD)")
    tao = _dam_bao_hd(khung, backup, da)
    k = K.doc_khung(khung); co = {str(d["stt"]) for d in k["dong"].get(MA_HD, [])}
    dot = k["dot_cuoi"].get(MA_HD, 0) + 1
    nh = f"NB_{ma_ns}"
    moi = [] if ma_ns in co else [dict(stt=ma_ns, noi_dung=f"Khai báo tay → {ns[ma_ns]}", dvt="đ", don_gia=1, pham_vi="TRONG_HD", nhom=nh,
                                       nhom_moi=(nh, f"Khai báo tay → {ns[ma_ns]}", "đ", ma_ns), nguon="webapp · khai báo tay")]
    x = dict(loai="THUC_HIEN", stt=ma_ns, kl=so_tien, dg=1.0, so_tien=None, ma_hd=MA_HD, dot=dot, ngay=ngay_d,
             ghi=f"{str(noi_dung).strip()[:120]}" + (f" · {ghi_chu.strip()[:80]}" if ghi_chu and ghi_chu.strip() else "") + f" · khai báo tay bởi {nguoi} {dt.datetime.now():%d/%m/%Y %H:%M}")
    if vat: x["vat_rieng"] = vat
    ph = dict(tt="N9_TT_DoiTac", ct="N7_HD_DoiTac_ChiTiet", ma_hd=MA_HD, dong_moi=moi, dong_tt=[x], lk_hstt=None, tang_hstt=so_tien)
    kq = G.ghi(khung, dict(phan=[ph]), f"khai báo tay · {str(noi_dung)[:40]}", backup)
    if not kq["ok"]: raise ValueError(kq["ly_do"])
    return dict(ok=True, ma_hd=MA_HD, dot=dot, ma_ns=ma_ns, ten_ns=ns[ma_ns], so_tien=so_tien, vat=vat, tao_hd=bool(tao), backup=kq["backup"],
                thong_bao=f"Đã ghi {so_tien:,.0f} (trước VAT) vào {ma_ns} — {ns[ma_ns]} · ngày {ngay_d:%d/%m/%Y}" + (" · đã tạo HĐ nội bộ HD-NOIBO" if tao else "") + " · đã backup")
