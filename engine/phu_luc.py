"""PHỤ LỤC HĐ = ĐIỀU CHỈNH / BỔ SUNG bảng giá HĐ, hiệu lực theo thứ tự PL01, PL02… (anh chốt 30/09).
Bảng giá đang hiệu lực của HĐ = bảng HĐ gốc đã áp LẦN LƯỢT mọi phụ lục đã ghi.

Hai dạng phụ lục (app tự nhận theo SỐ TIỀN, không đoán theo tên):
  DAY_DU  — PL in lại TOÀN BỘ BOQ sau điều chỉnh: Σ bảng PL = giá trị hiện hành (HĐ gốc + các PL đã ghi) + giá trị PL (±)
            ⇒ cập nhật TẠI CHỖ từng dòng N5/N7 khớp (mã hiệu HT…, không có mã thì STT + tên) theo PL mới nhất;
              dòng PL chưa có ⇒ thêm dòng; dòng HĐ không còn trong PL ⇒ KL 0
  BO_SUNG — PL chỉ liệt kê hạng mục THÊM MỚI: Σ bảng PL = giá trị PL ⇒ thêm dòng mới
STT giữ nguyên ⇒ HSTT (N8/N9 khoá mã HĐ + STT) vẫn nối đúng dòng; giá trị cũ → mới ghi vào cột nguồn của dòng (lịch sử).
Không khớp dạng nào ⇒ CHẶN (thường do phụ lục TRƯỚC chưa duyệt, hoặc AI đọc sai bảng)."""
import re, unicodedata, openpyxl, warnings
import cong_thuc as T
warnings.filterwarnings("ignore")

def _na(s): return unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().upper().replace("Đ", "D")
def _stt(s):
    s = str(s if s is not None else "").strip(); return s[:-2] if s.endswith(".0") else s
def ma_hieu(noi_dung):
    """Mã hiệu hạng mục: 'HT1 - Xây…' hoặc '… (HT1)'. Không bắt mác vữa/bê tông (M75, B20…) nằm giữa câu."""
    t = _na(noi_dung).strip()
    m = re.match(r"^([A-Z]{1,4}\d{1,4}[A-Z]?)\s*[-–:.)]", t) or re.search(r"\(([A-Z]{1,4}\d{1,4}[A-Z]?)\)\s*$", t)
    return m.group(1) if m else None
def _tu(s): return {w for w in re.sub(r"[^A-Z0-9 ]", " ", _na(s)).split() if len(w) > 1 and not re.fullmatch(r"[A-Z]{1,4}\d{1,4}[A-Z]?", w)}
def giong(a, b):
    x, y = _tu(a), _tu(b); return len(x & y) / max(1, min(len(x), len(y)))
def tien(x): return x["kl"] * x["don_gia"] if isinstance(x.get("kl"), (int, float)) and isinstance(x.get("don_gia"), (int, float)) else 0
def tol(v): return max(1000, abs(v) * 0.001)

def khop(cu, pl):
    """cu: dòng đang hiệu lực [dict(row, stt, noi_dung, dvt, kl, dg, kl_lk)]; pl: dòng PL (nhap_khung.dong_hop_le) ⇒ [(dòng PL, dòng cũ | None)]."""
    theo_ma, theo_stt, dung, out = {}, {}, set(), []
    for d in cu:
        m = ma_hieu(d["noi_dung"])
        if m: theo_ma.setdefault(m, []).append(d)
        theo_stt.setdefault(_stt(d["stt"]), []).append(d)
    for x in pl:
        m, c = ma_hieu(x["noi_dung"]), None
        if m and m in theo_ma: c = next((d for d in theo_ma[m] if d["row"] not in dung), None)
        if c is None and not (m and theo_ma):                                # không có mã hiệu ⇒ cùng STT + tên giống
            c = next((d for d in theo_stt.get(_stt(x["stt"]), []) if d["row"] not in dung and giong(d["noi_dung"], x["noi_dung"]) >= 0.5), None)
        if c: dung.add(c["row"])
        out.append((x, c))
    return out

def _so(v, n=4):
    if isinstance(v, float) and v and abs(v) < 1: return f"{v:.6g}"                  # KL lot rất nhỏ (tiện ích 2%…) — đủ chữ số để thấy đổi
    return f"{v:,.{n}f}".rstrip("0").rstrip(".") if isinstance(v, float) else f"{v:,}" if isinstance(v, int) else str(v)
def thay_doi(x, c):
    """Các thay đổi dòng PL so với dòng đang hiệu lực: [(trường, cũ, mới)]."""
    t = []
    if isinstance(x["kl"], (int, float)) and abs((c["kl"] or 0) - x["kl"]) > max(1e-9, abs(x["kl"]) * 1e-9): t.append(("kl", c["kl"], x["kl"]))
    if isinstance(x["don_gia"], (int, float)) and abs((c["dg"] or 0) - x["don_gia"]) > 0.5: t.append(("dg", c["dg"], x["don_gia"]))
    if x.get("dvt") and _na(x["dvt"]).replace(" ", "") != _na(c["dvt"]).replace(" ", ""): t.append(("dvt", c["dvt"], x["dvt"]))
    return t

def danh_gia(cu, pl, gt_hien_hanh, g_pl, so_pl_da_ghi):
    """⇒ dict(che_do DAY_DU | BO_SUNG | None, g_ghi (giá trị PL ghi N4/N6), tong, ky_vong, cap_nhat [(x, c, thay)], them [x], mat [c], co [(muc, mô tả)])."""
    tong, g = sum(tien(x) for x in pl), g_pl if isinstance(g_pl, (int, float)) else None
    ky = gt_hien_hanh + (g or 0); kp = khop(cu, pl); n_khop = sum(1 for _, c in kp if c)
    kq = dict(che_do=None, g_ghi=g, tong=tong, ky_vong=ky, cap_nhat=[], them=[], mat=[], co=[]); co = kq["co"]
    da_ghi = ", ".join(so_pl_da_ghi) or "chưa có"
    if g is not None and cu and abs(tong - ky) <= tol(ky): kq["che_do"] = "DAY_DU"
    elif g is not None and abs(tong - g) <= tol(g): kq["che_do"] = "BO_SUNG"
    elif g is None and cu and n_khop >= 0.8 * len(cu):
        kq.update(che_do="DAY_DU", g_ghi=tong - gt_hien_hanh); co.append(("LUU_Y", f"PL không ghi giá trị điều chỉnh ⇒ app lấy Σ bảng PL − giá trị hiện hành = {tong - gt_hien_hanh:,.0f}"))
    elif g is None and n_khop == 0:
        kq.update(che_do="BO_SUNG", g_ghi=tong); co.append(("LUU_Y", f"PL không ghi giá trị ⇒ app lấy Σ bảng PL = {tong:,.0f} (hạng mục bổ sung)"))
    else:
        co.append(("CHAN", f"Σ bảng PL {tong:,.0f} không khớp dạng nào: ① BOQ đầy đủ sau điều chỉnh = giá trị hiện hành {gt_hien_hanh:,.0f} "
                           f"{'+' if (g or 0) >= 0 else '−'} giá trị PL {abs(g or 0):,.0f} = {ky:,.0f} (lệch {tong - ky:,.0f}) · ② chỉ hạng mục bổ sung = giá trị PL. "
                           f"Phụ lục đã ghi: {da_ghi} — còn phụ lục TRƯỚC chưa duyệt thì duyệt theo thứ tự PL01, PL02…; không thì AI đọc sai bảng, anh kiểm rồi bấm Đọc lại"))
        return kq
    if kq["che_do"] == "DAY_DU":
        for x, c in kp:
            if c is None: kq["them"].append(x)
            elif thay_doi(x, c): kq["cap_nhat"].append((x, c, thay_doi(x, c)))
        dung = {c["row"] for _, c in kp if c}
        kq["mat"] = [d for d in cu if d["row"] not in dung and (d["kl"] or 0)]
        co.append(("LUU_Y", f"PL dạng BOQ ĐẦY ĐỦ sau điều chỉnh (Σ {tong:,.0f} = hiện hành {gt_hien_hanh:,.0f} {'+' if (kq['g_ghi'] or 0) >= 0 else '−'} {abs(kq['g_ghi'] or 0):,.0f}) — "
                           f"duyệt sẽ CẬP NHẬT bảng giá HĐ theo PL này: {len(kq['cap_nhat'])} dòng đổi KL/ĐG/ĐVT · {len(kq['them'])} dòng mới · {len(kq['mat'])} dòng không còn (KL → 0); lịch sử cũ → mới ghi ở cột nguồn"))
        if abs(tong - ky) > 1000: co.append(("LUU_Y", f"Σ bảng PL lệch giá trị kỳ vọng {tong - ky:,.0f} (làm tròn)"))
        if kq["cap_nhat"]:
            ds = [f"{x['stt']} {ma_hieu(x['noi_dung']) or ''} " + ", ".join(f"{'KL' if k == 'kl' else 'ĐG' if k == 'dg' else 'ĐVT'} {_so(a)}→{_so(b)}" for k, a, b in t) for x, _, t in kq["cap_nhat"]]
            co.append(("LUU_Y", "Dòng thay đổi: " + " · ".join(ds[:12]) + (f" … (+{len(ds) - 12} dòng)" if len(ds) > 12 else "")))
        if kq["mat"]: co.append(("LUU_Y", "Dòng HĐ không còn trong PL ⇒ KL 0: " + ", ".join(f"{d['stt']} {str(d['noi_dung'])[:30]}" for d in kq["mat"][:8])))
        vuot = [(x, c) for x, c, t in kq["cap_nhat"] if isinstance(x["kl"], (int, float)) and (c.get("kl_lk") or 0) > x["kl"] * 1.0001 + 1e-6]
        vuot += [(None, d) for d in kq["mat"] if (d.get("kl_lk") or 0) > 1e-6]
        if vuot: co.append(("LUU_Y", "KL mới THẤP HƠN KL đã thanh toán: " + ", ".join(f"{c['stt']} (đã TT {_so(c['kl_lk'])} > {_so(x['kl'] if x else 0)})" for x, c in vuot[:8])))
    else:
        trung = [(x, c) for x, c in kp if c]
        if trung:
            co.append(("CHAN", f"PL dạng BỔ SUNG (Σ bảng = giá trị PL) nhưng {len(trung)} dòng trùng hạng mục đã có trong HĐ ("
                               + ", ".join(f"{x['stt']} {str(x['noi_dung'])[:25]}" for x, _ in trung[:5]) + ") — không xác định được là KL TĂNG THÊM hay KL MỚI; anh báo em cách hiểu"))
            return kq
        kq["them"] = list(pl)
        co.append(("LUU_Y", f"PL dạng BỔ SUNG hạng mục: {len(pl)} dòng mới, Σ {tong:,.0f} = giá trị PL — duyệt sẽ THÊM vào bảng giá HĐ"))
    return kq

def so_hd_chuan(s): return re.sub(r"[^A-Z0-9]", "", _na(s))
def doc_hien_hanh(khung, cdt, ma_dt, so_hd_goc=None):
    """Đọc (chỉ đọc) bảng giá ĐANG HIỆU LỰC của HĐ gốc từ khung ⇒ dict(ma_hd, gt, so_pl, cu) — ma_hd None nếu chưa có HĐ gốc."""
    sh_hd, sh_ct = ("N4_HD_CDT", "N5_BOQ_CDT") if cdt else ("N6_HD_DoiTac", "N7_HD_DoiTac_ChiTiet")
    ch = T.cot(T.COT_HD[sh_hd]); cc = T.cot(T.CT_COLS); ix = lambda L: openpyxl.utils.column_index_from_string(L) - 1
    wb = openpyxl.load_workbook(khung, read_only=True, data_only=True)
    try:
        hd = [r for r in wb[sh_hd].iter_rows(min_row=2, values_only=True) if r and r[0]]
        goc = "HD-CDT" if cdt else f"HD-{ma_dt}"
        cua = [r for r in hd if str(r[0]) == goc or str(r[0]).startswith(goc + "-")]
        ma_hd = goc if any(str(r[0]) == goc for r in cua) else None
        if not cdt and so_hd_goc:                                            # đối tác nhiều HĐ ⇒ theo số HĐ gốc ghi trên PL
            ma_hd = next((str(r[0]) for r in cua if r[ix(ch["loai_ghi"])] != "PHU_LUC" and so_hd_chuan(r[ix(ch["so_hd"])]) == so_hd_chuan(so_hd_goc)), ma_hd)
        if not ma_hd: return dict(ma_hd=None, gt=0, so_pl=[], cu=[])
        gt = sum(r[ix(ch["gia_tri_truoc_vat"])] or 0 for r in hd if str(r[0]) == ma_hd and isinstance(r[ix(ch["gia_tri_truoc_vat"])], (int, float)))
        so_pl = [str(r[ix(ch["so_hd"])] or "").strip() for r in hd if str(r[0]) == ma_hd and r[ix(ch["loai_ghi"])] == "PHU_LUC"]
        cu = [dict(row=n, stt=r[ix(cc["stt"])], noi_dung=r[ix(cc["noi_dung"])], dvt=r[ix(cc["dvt"])], kl=r[ix(cc["kl_hd"])], dg=r[ix(cc["don_gia"])],
                   kl_lk=r[ix(cc["kl_luy_ke"])] if isinstance(r[ix(cc["kl_luy_ke"])], (int, float)) else 0)
              for n, r in enumerate(wb[sh_ct].iter_rows(min_row=2, values_only=True), 2) if r and r[0] == ma_hd]
        return dict(ma_hd=ma_hd, gt=gt, so_pl=so_pl, cu=cu)
    finally: wb.close()

def co_phu_luc(khung, rec):
    """Cờ đối chiếu của 1 phụ lục (hồ sơ nền) với bảng giá đang hiệu lực trong khung — gọi lúc AI đọc xong và mỗi khi khung đổi."""
    import nhap_khung as NK
    a = rec.get("ai") or {}; cdt = rec.get("loai") == "HD_CDT"
    h = doc_hien_hanh(khung, cdt, rec.get("ma_doi_tac") or "", a.get("so_hd_goc"))
    goc = "HD-CDT" if cdt else f"HD-{rec.get('ma_doi_tac') or ''}"
    if not h["ma_hd"]: return [dict(muc="CHAN", mo_ta=f"Phụ lục HĐ — chưa có HĐ gốc {goc} trong khung: anh duyệt HĐ gốc trước", pl=True)]
    so_pl = str(a.get("so_hd") or "PL").strip()
    if so_pl in h["so_pl"]: return [dict(muc="CHAN", mo_ta=f"Phụ lục {so_pl} của {h['ma_hd']} đã ghi trước đó — không ghi lần 2", pl=True)]
    kq = danh_gia(h["cu"], NK.dong_hop_le(a.get("bang")), h["gt"], a.get("gia_tri_truoc_vat"), h["so_pl"])
    return [dict(muc="LUU_Y", mo_ta=f"Phụ lục của {h['ma_hd']} (sau {len(h['so_pl'])} PL đã ghi: {', '.join(h['so_pl']) or 'chưa có'})", pl=True)] + [dict(muc=m, mo_ta=t, pl=True) for m, t in kq["co"]]
