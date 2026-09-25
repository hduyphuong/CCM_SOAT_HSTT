"""AI ĐỌC HỒ SƠ bằng Claude Code trên máy (gói Pro/Max của người dùng — không cần API key).
Gọi `claude -p` chế độ sạch: không hook / không cấu hình cá nhân (--setting-sources project trong thư mục tạm), chỉ công cụ Read,
kết quả ép theo JSON Schema (--json-schema). PDF scan đọc theo trang; Excel được trích thành văn bản rồi đưa vào prompt.
Nguyên tắc: không đọc rõ ⇒ null + ghi 'khong_chac' — KHÔNG ĐOÁN số. Mọi kết quả phải được người dùng duyệt trước khi ghi khung."""
import subprocess, json, os, shutil, tempfile, time, openpyxl, warnings
warnings.filterwarnings("ignore")
S, I, N, B = {"type": ["string", "null"]}, {"type": ["integer", "null"]}, {"type": ["number", "null"]}, {"type": ["boolean", "null"]}
DONG = {"type": "object", "properties": {"stt": S, "noi_dung": S, "dvt": S, "kl": N, "don_gia": N, "thanh_tien": N, "ghi_chu": S}, "required": ["noi_dung"]}
SCHEMA_NEN = {"type": "object", "properties": {
    "loai": {"type": "string", "enum": ["HD_CDT", "HD_DOI_TAC", "PLHD_CDT", "PLHD_DOI_TAC", "BOQ_CDT", "NGAN_SACH", "GOI_THAU", "BAO_GIA",
                                         "QUYET_TOAN", "BIEN_BAN", "TO_TRINH", "HSTT", "KHAC"]},
    "ly_do_loai": S, "so_trang": I, "so_hd": S, "so_hd_goc": S, "ngay_ky": S, "ben_giao": S, "ben_nhan": S, "doi_tac_ten": S, "doi_tac_mst": S,
    "loai_doi_tac": {"type": ["string", "null"], "enum": ["CDT", "DTC", "NTP", "NCC", "DVK", None]}, "du_an": S, "noi_dung": S,
    "dang_hd": {"type": ["string", "null"], "enum": ["DON_GIA", "TRON_GOI", "NGUYEN_TAC", None]},
    "gia_tri_truoc_vat": N, "vat_pct": N, "gia_tri_sau_vat": N, "pct_tam_ung": N, "pct_tt_dot": N, "pct_tt_quyet_toan": N, "pct_giu_lai": N,
    "han_tt_ngay": I, "don_vi_han": {"type": ["string", "null"], "enum": ["LV", "LICH", None]}, "han_qt_ngay": I, "han_tra_gl_ngay": I, "bao_hanh_thang": I,
    "bang": {"type": "array", "items": DONG}, "tong_ghi_tren_file": N, "co_chu_ky": B, "co_dong_dau": B,
    "nguon": {"type": "object", "additionalProperties": {"type": "string"}}, "khong_chac": {"type": "array", "items": {"type": "string"}}, "ghi_chu": S},
    "required": ["loai", "ly_do_loai", "bang", "khong_chac"]}
HUONG_DAN = """Bạn là chuyên viên QS/CCM người Việt, đọc hồ sơ xây dựng để nhập vào sổ kiểm soát chi phí. Chính xác tuyệt đối về số.
Công ty người dùng (nhà thầu thi công): {cty}. Công ty người dùng NHẬN thầu từ chủ đầu tư/thầu chính ⇒ phía CĐT (HD_CDT/PLHD_CDT, loai_doi_tac=CDT).
Công ty người dùng GIAO việc cho đội/thầu phụ/nhà cung cấp ⇒ phía đối tác (HD_DOI_TAC/PLHD_DOI_TAC; DTC=đội thi công/tổ đội khoán nhân công,
NTP=thầu phụ pháp nhân, NCC=cung cấp vật tư/hàng hoá, DVK=dịch vụ khác).
Quy tắc: tiền = số VND (không dấu phân cách); phần trăm = thập phân (10% → 0.1); ngày = YYYY-MM-DD; 'bang' = bảng khối lượng/đơn giá/ngân sách
(mỗi dòng 1 hạng mục, bỏ dòng tiêu đề nhóm và dòng cộng); 'tong_ghi_tren_file' = số tổng in trên file để đối chiếu; 'nguon' ghi trang/ô lấy từng số tiền.
Không có hoặc đọc không rõ ⇒ để null VÀ thêm tên trường vào 'khong_chac'. TUYỆT ĐỐI KHÔNG ĐOÁN SỐ."""

def trich_excel(path, toi_da=1500):
    """Excel ⇒ văn bản 'Sheet | ô: giá trị' (giá trị đã tính) để AI đọc; cắt bớt khi quá dài."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True); out = []
    for ws in wb.worksheets:
        n = 0; out.append(f"=== SHEET: {ws.title}")
        for r in ws.iter_rows(values_only=True):
            v = [str(x).strip() for x in r if x not in (None, "")]
            if v: out.append(" | ".join(v)[:400]); n += 1
            if n >= toi_da: out.append("…(cắt bớt)"); break
    wb.close(); return "\n".join(out)

def doc(path, cty="(chưa khai báo)", schema=None, huong_dan=None, timeout=1500, model=None):
    """Trả (ket_qua_dict, meta). Lỗi/hết hạn mức ⇒ raise RuntimeError có thông điệp dễ hiểu."""
    d = tempfile.mkdtemp(prefix="ccm_ai_"); duoi = os.path.splitext(path)[1].lower()
    try:
        if duoi == ".pdf":
            shutil.copy2(path, os.path.join(d, "hs.pdf"))
            yeu_cau = "Đọc file hs.pdf trong thư mục hiện tại bằng công cụ Read, lần lượt theo tham số pages (tối đa 20 trang/lần) tới HẾT file, rồi trả kết quả."
        elif duoi in (".xlsx", ".xlsm"):
            open(os.path.join(d, "hs.txt"), "w", encoding="utf-8").write(trich_excel(path))      # KHÔNG nhét vào dòng lệnh: Windows giới hạn ~32k ký tự (WinError 206)
            yeu_cau = ("Đọc file hs.txt trong thư mục hiện tại bằng công cụ Read (nội dung file Excel đã trích thành văn bản: '=== SHEET: tên' rồi mỗi dòng "
                       "là các ô cách nhau bởi ' | '). File có thể dài — đọc tiếp bằng tham số offset/limit tới HẾT file, rồi trả kết quả.")
        else: raise RuntimeError(f"Chưa hỗ trợ AI đọc đuôi {duoi}")
        cmd = ["claude", "-p", yeu_cau, "--output-format", "json", "--json-schema", json.dumps(schema or SCHEMA_NEN, ensure_ascii=False),
               "--system-prompt", (huong_dan or HUONG_DAN).format(cty=cty), "--setting-sources", "project", "--tools", "Read",
               "--allowedTools", "Read", "--max-turns", "14", "--no-session-persistence"] + (["--model", model] if model else [])
        t0 = time.time()
        p = subprocess.run(cmd, cwd=d, capture_output=True, text=True, encoding="utf-8", timeout=timeout, shell=(os.name == "nt"))
        try: o = json.loads(p.stdout)
        except Exception: raise RuntimeError("Claude không trả kết quả (chưa đăng nhập Claude Code / hết hạn mức gói?): " + (p.stderr or p.stdout)[-300:])
        if o.get("is_error"): raise RuntimeError("Claude báo lỗi: " + str(o.get("result"))[:300])
        kq = o.get("structured_output")
        if kq is None:
            r = o.get("result") or ""; kq = json.loads(r[r.find("{"): r.rfind("}") + 1])
        return kq, dict(giay=round(time.time() - t0), luot=o.get("num_turns"), usd_quy_doi=o.get("total_cost_usd"))
    finally: shutil.rmtree(d, ignore_errors=True)
