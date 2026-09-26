"""HSTT BẢN SCAN / PDF KÝ — AI đọc (Claude Code, gói Pro của người dùng) ⇒ dựng hồ sơ CHUẨN như bộ đọc Excel ⇒ cùng 4 lớp kiểm của engine.
Chống đọc nhầm: số tiền AI 'không chắc' ⇒ CHẶN · bất biến số học (Σ dòng = tổng · KL×ĐG · KT+KN=LK) do kiem.py + ở đây soát ·
luôn có lưu ý 'số do AI đọc từ scan'. Kết quả AI lưu trong sổ nạp ⇒ soát lại / ghi sổ KHÔNG gọi AI lần 2."""
import datetime as dt, ai_doc
from doc_hstt import na
S, I, N, B = ai_doc.S, ai_doc.I, ai_doc.N, ai_doc.B
DONG = {"type": "object", "properties": {"stt": S, "noi_dung": S, "dvt": S, "don_gia": N, "kl_ky_truoc": N, "kl_ky_nay": N, "kl_luy_ke": N,
                                         "tt_ky_truoc": N, "tt_ky_nay": N, "tt_luy_ke": N, "trang": I}, "required": ["noi_dung"]}
SCHEMA = {"type": "object", "properties": {
    "loai_ho_so": {"type": "string", "enum": ["THANH_TOAN", "TAM_UNG", "QUYET_TOAN", "KHAC"]}, "ly_do_loai": S, "so_trang": I,
    "don_vi": S, "so_hd": S, "du_an": S, "dot": I, "ngay": S, "vat_pct": N,
    "tong_ky_truoc": N, "tong_ky_nay": N, "tong_luy_ke": N,
    "tam_ung_de_nghi": N, "thu_hoi_tam_ung_ky_nay": N, "giu_lai_ky_nay": N, "de_nghi_thanh_toan": N,
    "bang": {"type": "array", "items": DONG}, "co_chu_ky": B,
    "nguon": {"type": "object", "additionalProperties": {"type": "string"}}, "khong_chac": {"type": "array", "items": {"type": "string"}}, "ghi_chu": S},
    "required": ["loai_ho_so", "bang", "khong_chac"]}
HUONG_DAN = """Bạn là QS/CCM người Việt, đọc BẢN SCAN hồ sơ thanh toán (HSTT) do đội thi công / thầu phụ / nhà cung cấp gửi công ty {cty}, để nhập sổ kiểm soát chi phí. Chính xác tuyệt đối về số.
'don_vi' = đơn vị ĐỀ NGHỊ được thanh toán (đội / thầu phụ / NCC), KHÔNG phải {cty}. 'du_an' = tên dự án/công trình ghi trên hồ sơ — hồ sơ KHÔNG ghi thì để null.
'loai_ho_so': TAM_UNG nếu chỉ là giấy đề nghị tạm ứng (không có khối lượng thực hiện kỳ này); THANH_TOAN nếu có bảng khối lượng / giá trị thực hiện; QUYET_TOAN nếu là hồ sơ quyết toán.
Tiền = số VND (không dấu phân cách); % = thập phân (8% → 0.08); ngày = YYYY-MM-DD (ngày lập / ký hồ sơ). Tổng kỳ trước / kỳ này / lũy kế = giá trị TRƯỚC VAT.
'bang' = từng dòng công việc của bảng khối lượng / giá trị (KHÔNG đưa dòng tiêu đề nhóm, dòng cộng, dòng tổng); KL và thành tiền kỳ trước / kỳ này / lũy kế đọc ĐÚNG CỘT; mỗi dòng ghi 'trang'.
'nguon' = trang lấy từng số tổng (khoá = tên trường). Số nào mờ / bị che / không đọc rõ ⇒ để null VÀ thêm TÊN TRƯỜNG (vd 'ngay', 'tong_ky_nay') vào 'khong_chac'; giải thích để ở 'ghi_chu'. TUYỆT ĐỐI KHÔNG ĐOÁN, KHÔNG TỰ TÍNH BÙ số."""
TIEN = ("tong_ky_truoc", "tong_ky_nay", "tong_luy_ke", "tam_ung_de_nghi", "thu_hoi_tam_ung_ky_nay", "giu_lai_ky_nay", "de_nghi_thanh_toan", "dot", "vat_pct")
NGUONG = 10

def doc(path, cty):
    """Gọi AI (chạy lâu 1–3 phút) — trả (kết quả thô, meta)."""
    return ai_doc.doc(path, cty, schema=SCHEMA, huong_dan=HUONG_DAN, timeout=1800)

def _n(v): return float(v) if isinstance(v, (int, float)) else 0.0

def dung_hs(ai, path):
    """Kết quả AI ⇒ hồ sơ cùng cấu trúc bộ đọc Excel (cover · lines · tong · tu/hu/gl/du_tru) + cờ riêng của bản scan."""
    kr = []; add = lambda *x: kr.append(x); ng = ai.get("nguon") or {}; o = lambda k: f"PDF trang {ng.get(k, '?')}"
    loai = ai.get("loai_ho_so") or "KHAC"; kc = [str(x) for x in (ai.get("khong_chac") or [])]
    lines = []
    for i, x in enumerate(ai.get("bang") or [], 1):
        if loai == "TAM_UNG": break                                      # tạm ứng: bảng KL đi kèm chỉ là SƠ BỘ ⇒ không ghi thực hiện
        nd_ = na(x.get("noi_dung")).strip()
        if nd_.startswith(("TONG", "CONG")) or (not x.get("dvt") and not _n(x.get("kl_ky_nay")) and not _n(x.get("kl_luy_ke"))): continue   # dòng nhóm / dòng cộng
        kt, kn = _n(x.get("kl_ky_truoc")), _n(x.get("kl_ky_nay"))
        lk = _n(x["kl_luy_ke"]) if isinstance(x.get("kl_luy_ke"), (int, float)) else kt + kn
        tk, tn = _n(x.get("tt_ky_truoc")), _n(x.get("tt_ky_nay"))
        tl = _n(x["tt_luy_ke"]) if isinstance(x.get("tt_luy_ke"), (int, float)) else tk + tn
        lines.append(dict(dong=f"tr.{x.get('trang') or '?'} #{i}", khung="", stt=str(x.get("stt") or i), ds=str(x.get("noi_dung") or "").strip(),
                          dvt=str(x.get("dvt") or "").strip(), kl_hd=None, dg=_n(x.get("don_gia")), kl_kt=kt, kl_kn=kn, kl_lk=lk,
                          tt_kt=tk, tt_kn=tn, tt_lk=tl, ngoai=False))
    sm = lambda k: sum(l[k] for l in lines)
    t = tuple(_n(ai[f]) if isinstance(ai.get(f), (int, float)) else sm(k) for f, k in (("tong_ky_truoc", "tt_kt"), ("tong_ky_nay", "tt_kn"), ("tong_luy_ke", "tt_lk")))
    if loai == "TAM_UNG": t = (0.0, 0.0, 0.0)                              # tạm ứng không phát sinh thực hiện ⇒ lũy kế lấy từ khung ở chuan_bi()
    try: ngay = dt.date.fromisoformat(str(ai.get("ngay"))[:10])
    except Exception: ngay = None
    # ── cờ riêng của bản scan (thêm vào 4 lớp kiểm chung) ──
    add("LUU_Y", "Hồ sơ", "PDF", ai.get("so_trang") or "—", "AI", "Số liệu do AI ĐỌC TỪ BẢN SCAN — anh đối chiếu các số trọng yếu với bản ký trước khi ghi sổ")
    if loai in ("QUYET_TOAN", "KHAC"): add("CHAN", "Hồ sơ", "PDF", loai, "—", f"AI nhận là hồ sơ '{loai}' — chế độ AI đọc scan mới hỗ trợ THANH TOÁN / TẠM ỨNG")
    tien_kc = [x for x in kc if x in TIEN or x.startswith("bang")]
    if tien_kc: add("CHAN", "Hồ sơ", "PDF", ", ".join(tien_kc)[:60], "không chắc", "AI KHÔNG CHẮC các số này — anh xem bản scan, bấm đọc lại hoặc xin bản Excel")
    khac = [x for x in kc if x not in tien_kc]
    if khac: add("LUU_Y", "Hồ sơ", "PDF", f"{len(khac)} điểm", "—", "AI không chắc: " + " · ".join(x[:70] for x in khac[:3]) + (" …" if len(khac) > 3 else ""))
    if loai == "TAM_UNG" and ai.get("bang"): add("LUU_Y", "Hồ sơ", "PDF", f"{len(ai['bang'])} dòng", "—", "Bảng khối lượng kèm hồ sơ tạm ứng chỉ là SƠ BỘ — KHÔNG ghi thực hiện, chỉ ghi tạm ứng")
    if ai.get("dot") is None: add("CHAN", "Hồ sơ", o("dot"), "—", "—", "AI không đọc được SỐ ĐỢT — không chống ghi trùng được")
    if not ai.get("don_vi"): add("CHAN", "Hồ sơ", o("don_vi"), "—", "—", "AI không đọc được tên đơn vị")
    if loai == "TAM_UNG" and not _n(ai.get("tam_ung_de_nghi")): add("CHAN", "Số học", o("tam_ung_de_nghi"), "—", "—", "Hồ sơ tạm ứng nhưng AI không đọc được số tiền đề nghị tạm ứng")
    if loai == "THANH_TOAN" and not lines: add("CHAN", "Số học", "PDF", "0 dòng", "—", "AI không đọc được bảng khối lượng / giá trị")
    if loai == "THANH_TOAN" and abs(t[0] + t[1] - t[2]) > NGUONG: add("CHAN", "Số học", o("tong_luy_ke"), round(t[2]), round(t[0] + t[1]), "Tổng lũy kế ≠ kỳ trước + kỳ này (AI đọc)")
    if ai.get("co_chu_ky") is False: add("LUU_Y", "Hồ sơ", "PDF", "—", "—", "AI không thấy chữ ký trên bản scan")
    thu_hoi, giu_lai = _n(ai.get("thu_hoi_tam_ung_ky_nay")), _n(ai.get("giu_lai_ky_nay"))
    cv = dict(ten_don_vi=ai.get("don_vi"), so_hd=ai.get("so_hd"), dot=ai.get("dot"), ngay=ngay, o={"so_hd": o("so_hd"), "dot": o("dot"), "ten_don_vi": o("don_vi")})
    return dict(file=path, sheet="PDF", cover=cv, lines=lines, tong=t, vat=ai.get("vat_pct"), tu=0.0, hu=0.0, tu_k=0.0, hu_k=-thu_hoi,
                gl=-giu_lai if giu_lai else None, du_tru=0.0, du_an_text="" if not ai.get("du_an") or any(x in na(ai["du_an"]) for x in ("KHONG GHI", "KHONG NEU", "KHONG CO")) else na(ai["du_an"]),
                mau="SCAN_AI", kiem_rieng=kr, tam_ung_moi=_n(ai.get("tam_ung_de_nghi")), thu_hoi=thu_hoi, loai_ai=loai)

def chuan_bi(hs, k):
    """Tạm ứng còn treo SAU hồ sơ này = treo trong khung + tạm ứng mới − thu hồi ⇒ để kiem/ke_hoach nhận đúng loại (TAM_UNG / THANH_TOAN)."""
    import kiem as K
    ma = K.phan_loai(hs, k)[0].get("ma_hd")
    if ma: hs["du_tru"] = k["tu_treo"][ma] + hs.get("tam_ung_moi", 0.0) - hs.get("thu_hoi", 0.0)
    if ma and hs.get("loai_ai") == "TAM_UNG": hs["tong"] = (k["lk_tien"][ma], 0.0, k["lk_tien"][ma])   # tạm ứng: lũy kế thực hiện KHÔNG đổi
