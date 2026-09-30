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
    "ly_do_loai": S, "so_trang": I, "so_hd": S, "so_hd_goc": S, "ngay_ky": S, "ben_giao": S, "ben_nhan": S, "ben_tra_tien": S, "ben_nhan_tien": S, "doi_tac_ten": S, "doi_tac_mst": S,
    "loai_doi_tac": {"type": ["string", "null"], "enum": ["CDT", "DTC", "NTP", "NCC", "DVK", None]}, "du_an": S, "noi_dung": S,
    "dang_hd": {"type": ["string", "null"], "enum": ["DON_GIA", "TRON_GOI", "NGUYEN_TAC", None]},
    "gia_tri_truoc_vat": N, "vat_pct": N, "gia_tri_sau_vat": N, "pct_tam_ung": N, "pct_tt_dot": N, "pct_tt_quyet_toan": N, "pct_giu_lai": N,
    "han_tt_ngay": I, "don_vi_han": {"type": ["string", "null"], "enum": ["LV", "LICH", None]}, "han_qt_ngay": I, "han_tra_gl_ngay": I, "bao_hanh_thang": I,
    "bang": {"type": "array", "items": DONG}, "tong_ghi_tren_file": N, "co_chu_ky": B, "co_dong_dau": B,
    "nguon": {"type": "object", "additionalProperties": {"type": "string"}}, "khong_chac": {"type": "array", "items": {"type": "string"}}, "ghi_chu": S,
    "doc_duoc": {"type": "boolean"}},
    "required": ["loai", "ly_do_loai", "bang", "khong_chac", "doc_duoc"]}
HUONG_DAN = """Bạn là chuyên viên QS/CCM người Việt, đọc hồ sơ xây dựng để nhập vào sổ kiểm soát chi phí. Chính xác tuyệt đối về số.
Công ty người dùng (nhà thầu thi công): {cty}. Công ty người dùng NHẬN thầu từ chủ đầu tư/thầu chính ⇒ phía CĐT (HD_CDT/PLHD_CDT, loai_doi_tac=CDT).
Công ty người dùng GIAO việc cho đội/thầu phụ/nhà cung cấp ⇒ phía đối tác (HD_DOI_TAC/PLHD_DOI_TAC; DTC=đội thi công/tổ đội khoán nhân công,
NTP=thầu phụ pháp nhân, NCC=cung cấp vật tư/hàng hoá, DVK=dịch vụ khác).
'ben_tra_tien' = bên THANH TOÁN tiền theo HĐ, 'ben_nhan_tien' = bên ĐƯỢC thanh toán (HĐ mua bán vật tư: bên MUA trả tiền; HĐ giao thầu/giao khoán: bên giao việc trả tiền).
Quy tắc: tiền = số VND (không dấu phân cách); phần trăm = thập phân (10% → 0.1); ngày = YYYY-MM-DD; 'bang' = bảng khối lượng/đơn giá/ngân sách
(mỗi dòng 1 hạng mục, bỏ dòng tiêu đề nhóm và dòng cộng); 'tong_ghi_tren_file' = số tổng in trên file để đối chiếu; 'nguon' ghi trang/ô lấy từng số tiền.
Số hợp đồng chép NGUYÊN VĂN từng ký tự, giữ đủ chữ 'Đ' có gạch (HĐGK, HĐTC, HĐTP, HĐNT… — KHÔNG viết thành HGK/HDGK). Không có hoặc đọc không rõ ⇒ để null VÀ thêm tên trường vào 'khong_chac'. TUYỆT ĐỐI KHÔNG ĐOÁN SỐ.
'doc_duoc': true CHỈ KHI công cụ Read đã thực sự mở và trả về nội dung file (dù nội dung khó đọc/thiếu vài trường thì vẫn true, các trường đó ghi null + khong_chac).
'doc_duoc': false khi KHÔNG mở được file — file không tồn tại, rỗng, hỏng, hoặc công cụ Read báo lỗi bất kỳ lúc nào trong lúc đọc — dù chỉ 1 lần cũng phải false, TUYỆT ĐỐI KHÔNG được tự đoán/suy diễn 'loai' hay bất kỳ trường nào khác khi doc_duoc=false — mọi trường khác để null, 'bang' để [], 'ly_do_loai' ghi rõ lỗi Read gặp phải."""

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

MODEL_MAC_DINH, EFFORT_MAC_DINH = "claude-sonnet-5", "high"    # anh chốt 30/09 — trải nghiệm rồi đổi khi cần (Opus dùng khi thấy Sonnet sai)

def doc(path, cty="(chưa khai báo)", schema=None, huong_dan=None, timeout=1500, model=None, effort=None, ngu_canh=None):
    """Trả (ket_qua_dict, meta). Lỗi/hết hạn mức ⇒ raise RuntimeError có thông điệp dễ hiểu."""
    d = tempfile.mkdtemp(prefix="ccm_ai_"); duoi = os.path.splitext(path)[1].lower()
    try:
        if duoi == ".pdf":
            dich = os.path.join(d, "hs.pdf")
            for lan in range(5):                                            # nạp hàng loạt ⇒ Drive đôi khi chưa đồng bộ kịp; copy rỗng thì đợi rồi thử lại
                shutil.copy2(path, dich)
                if os.path.getsize(dich) > 0: break
                time.sleep(2)
            else: raise RuntimeError(f"File rỗng sau 5 lần thử copy (Drive có thể chưa đồng bộ xong): {os.path.basename(path)} — thử Đọc lại sau ít phút")
            yeu_cau = "Đọc file hs.pdf trong thư mục hiện tại bằng công cụ Read, lần lượt theo tham số pages (tối đa 20 trang/lần) tới HẾT file, rồi trả kết quả."
        elif duoi in (".xlsx", ".xlsm"):
            open(os.path.join(d, "hs.txt"), "w", encoding="utf-8").write(trich_excel(path))      # KHÔNG nhét vào dòng lệnh: Windows giới hạn ~32k ký tự (WinError 206)
            yeu_cau = ("Đọc file hs.txt trong thư mục hiện tại bằng công cụ Read (nội dung file Excel đã trích thành văn bản: '=== SHEET: tên' rồi mỗi dòng "
                       "là các ô cách nhau bởi ' | '). File có thể dài — đọc tiếp bằng tham số offset/limit tới HẾT file, rồi trả kết quả.")
        else: raise RuntimeError(f"Chưa hỗ trợ AI đọc đuôi {duoi}")
        if ngu_canh:                                                       # sổ sách hiện có ⇒ AI ghép đúng dòng HĐ, không đoán
            open(os.path.join(d, "ngu_canh.txt"), "w", encoding="utf-8").write(ngu_canh)
            yeu_cau += " TRƯỚC KHI đọc hồ sơ, đọc file ngu_canh.txt (danh sách hợp đồng + dòng HĐ đang có trong sổ, và lỗi của lần đọc trước nếu có) để ghép đúng."
        cmd = ["claude", "-p", yeu_cau, "--output-format", "json", "--json-schema", json.dumps(schema or SCHEMA_NEN, ensure_ascii=False),
               "--system-prompt", (huong_dan or HUONG_DAN).format(cty=cty), "--setting-sources", "project", "--tools", "Read",
               "--allowedTools", "Read", "--max-turns", "14", "--no-session-persistence",
               "--model", model or MODEL_MAC_DINH, "--effort", effort or EFFORT_MAC_DINH]
        t0 = time.time()
        p = subprocess.run(cmd, cwd=d, capture_output=True, text=True, encoding="utf-8", timeout=timeout, shell=(os.name == "nt"),
                           creationflags=0x08000000 if os.name == "nt" else 0)          # CREATE_NO_WINDOW: engine chạy ẩn thì không bật cửa sổ đen
        try: o = json.loads(p.stdout)
        except Exception: raise RuntimeError("Claude không trả kết quả (chưa đăng nhập Claude Code / hết hạn mức gói?): " + (p.stderr or p.stdout)[-300:])
        if o.get("is_error"): raise RuntimeError("Claude báo lỗi: " + str(o.get("result"))[:300])
        kq = o.get("structured_output")
        if kq is None:
            r = o.get("result") or ""; kq = json.loads(r[r.find("{"): r.rfind("}") + 1])
        return kq, dict(giay=round(time.time() - t0), luot=o.get("num_turns"), usd_quy_doi=o.get("total_cost_usd"))
    finally: shutil.rmtree(d, ignore_errors=True)
