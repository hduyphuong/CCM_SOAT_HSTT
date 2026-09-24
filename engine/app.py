"""ENGINE SOÁT HSTT — chạy trên máy anh:  python app.py      (mặc định http://127.0.0.1:8765)
Trang (GitHub Pages / file local) gọi vào đây. Dữ liệu CHỈ nằm trên máy: D:\\QLCP_HD\\WEBAPP_SOAT_HSTT_DATA\\<dự án>\\
  GET  /ping                      · GET /du-an                 · GET /ho-so?du_an=X
  POST /nap    {du_an, ten, b64}  → lưu file, phân loại, tự kiểm 4 lớp, kế hoạch ghi sổ
  POST /duyet  {du_an, id, hanh_dong: DONG_Y | YEU_CAU_SUA | TRA_DOI, ly_do}
  GET  /bao-cao?du_an=X           → R0 tổng quan + 90_Check (đọc từ file khung)"""
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json, os, sys, base64, datetime as dt, threading, traceback
from urllib.parse import urlparse, parse_qs
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import doc_hstt as D, kiem as K, ghi_so as G

PORT = 8765
DATA = r"D:\QLCP_HD\WEBAPP_SOAT_HSTT_DATA"
CAU_HINH = os.path.join(DATA, "du_an.json")          # {"DU_AN_A": {"ten": "...", "khung": "<đường dẫn file khung .xlsx>"}}
TRANG_TXT = os.path.join(DATA, "trang.txt")          # 1 dòng: https://<tài-khoản>.github.io — trang khác KHÔNG gọi được engine
TRANG = open(TRANG_TXT, encoding="utf-8").read().strip().rstrip("/") if os.path.exists(TRANG_TXT) else ""
KHOA = threading.Lock()                               # 1 lần ghi sổ tại 1 thời điểm

def _json(o):
    if isinstance(o, (dt.date, dt.datetime)): return o.isoformat()
    raise TypeError(type(o))
def du_an(): return json.load(open(CAU_HINH, encoding="utf-8")) if os.path.exists(CAU_HINH) else {}
def so_nap_path(da): return os.path.join(DATA, da, "so_nap.json")
def so_nap(da): p = so_nap_path(da); return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}
def luu_so(da, s):
    os.makedirs(os.path.join(DATA, da), exist_ok=True)
    json.dump(s, open(so_nap_path(da), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=_json)

def xu_ly_nap(b):
    da, ten = b["du_an"], os.path.basename(b["ten"])
    cfg = du_an()[da]; thu_muc = os.path.join(DATA, da, "nap"); os.makedirs(thu_muc, exist_ok=True)
    raw = base64.b64decode(b["b64"]); tam = os.path.join(thu_muc, "_tam_" + ten)
    open(tam, "wb").write(raw); vt = D.van_tay(tam)
    s = so_nap(da); i = vt[:12]
    cu = {v["van_tay"] for k_, v in s.items() if k_ != i and v["trang_thai"] != "TRA_DOI"}     # trùng nội dung với HỒ SƠ KHÁC
    dich = os.path.join(thu_muc, f"{i}_{ten}")
    if os.path.exists(dich): os.remove(tam)                        # nạp lại đúng file cũ: giữ bản gốc chỉ đọc
    else: os.replace(tam, dich); os.chmod(dich, 0o444)
    hs = D.doc_file(dich); hs["van_tay"] = vt
    k = K.doc_khung(cfg["khung"]); k["tu_khoa"] = cfg.get("tu_khoa", []); kq = K.kiem(hs, k, cu)
    kh = G.ke_hoach(hs, kq, k) if kq["phan_loai"]["ma_hd"] else None
    cu_rec = s.get(i)
    if cu_rec and cu_rec["trang_thai"] == "DA_GHI_SO":           # đã ghi sổ ⇒ không cho ghi lần 2
        cu_rec["co"] = [dict(muc="CHAN", lop="Hồ sơ", vi_tri="file", hstt=i, doi_chieu="đã ghi sổ", mo_ta="Hồ sơ này ĐÃ GHI SỔ lúc " + str(cu_rec.get("luc_duyet")))]
        return cu_rec
    rec = dict(id=i, van_tay=vt, ten=ten, file=dich, luc=dt.datetime.now().isoformat(timespec="seconds"),
               trang_thai=cu_rec["trang_thai"] if cu_rec else "CHO_DUYET",
               phan_loai=kq["phan_loai"], tom_tat=kq["tom_tat"], co=kq["co"], ke_hoach=kh, ly_do=cu_rec.get("ly_do") if cu_rec else None, ket_qua=None)
    s[i] = rec; luu_so(da, s)
    return rec

def xu_ly_duyet(b):
    da, i, hd = b["du_an"], b["id"], b["hanh_dong"]
    s = so_nap(da); rec = s[i]
    if rec["trang_thai"] == "DA_GHI_SO": return dict(ok=False, ly_do="Hồ sơ đã ghi sổ trước đó")
    if hd == "DONG_Y":
        if any(c["muc"] == "CHAN" for c in rec["co"]): return dict(ok=False, ly_do="Còn cờ CHẶN — nút Đồng ý bị khoá")
        with KHOA:
            cfg = du_an()[da]
            hs = D.doc_file(rec["file"]); hs["van_tay"] = rec["van_tay"]
            k = K.doc_khung(cfg["khung"]); k["tu_khoa"] = cfg.get("tu_khoa", []); kq = K.kiem(hs, k)                  # kiểm lại ngay trước khi ghi (khung có thể đã đổi)
            if any(c["muc"] == "CHAN" for c in kq["co"]): return dict(ok=False, ly_do="Kiểm lại trước khi ghi phát sinh cờ CHẶN", co=kq["co"])
            kh = G.ke_hoach(hs, kq, k)
            kq_ghi = G.ghi(cfg["khung"], kh, rec["ten"], os.path.join(DATA, da, "backup"))
        rec["ket_qua"] = kq_ghi
        if kq_ghi["ok"]: rec["trang_thai"] = "DA_GHI_SO"
    else:
        rec["trang_thai"] = hd; rec["ly_do"] = b.get("ly_do") or ""
    rec["luc_duyet"] = dt.datetime.now().isoformat(timespec="seconds"); s[i] = rec; luu_so(da, s)
    return dict(ok=rec["trang_thai"] != "CHO_DUYET", ho_so=rec)

def bao_cao(da):
    import openpyxl
    wb = openpyxl.load_workbook(du_an()[da]["khung"], data_only=True, read_only=True)
    r0 = [(r[1], r[2]) for r in wb["R0_TongQuan"].iter_rows(min_row=4, max_row=30, max_col=4, values_only=True) if r[1]]
    ck = [dict(so=r[0], ten=r[1], trai=r[2], phai=r[3], ket_luan=r[5]) for r in wb["90_Check"].iter_rows(min_row=2, max_row=60, max_col=7, values_only=True) if r[1]]
    wb.close()
    return dict(tong_quan=r0, kiem=ck, ghi_chu="Số đọc từ lần Excel tính gần nhất (sau mỗi lần ghi sổ).")

class H(BaseHTTPRequestHandler):
    def _cors(self):
        o = self.headers.get("Origin", "")
        if (TRANG and o == TRANG) or o.startswith("http://localhost:") or o.startswith("http://127.0.0.1:"):   # CHỈ trang của anh + chạy thử tại máy
            self.send_header("Access-Control-Allow-Origin", o or "*"); self.send_header("Vary", "Origin")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Private-Network", "true")
    def _tra(self, code, obj):
        b = json.dumps(obj, ensure_ascii=False, default=_json).encode("utf-8")
        self.send_response(code); self._cors(); self.send_header("Content-Type", "application/json; charset=utf-8"); self.end_headers(); self.wfile.write(b)
    def do_OPTIONS(self): self.send_response(204); self._cors(); self.end_headers()
    def do_GET(self):
        u = urlparse(self.path); q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path == "/ping": return self._tra(200, dict(ok=True, engine="soat-hstt", phien_ban="1.0"))
            if u.path == "/du-an": return self._tra(200, {k: v.get("ten", k) for k, v in du_an().items()})
            if u.path == "/ho-so": return self._tra(200, sorted(so_nap(q["du_an"]).values(), key=lambda r: r["luc"], reverse=True))
            if u.path == "/bao-cao": return self._tra(200, bao_cao(q["du_an"]))
            self._tra(404, dict(loi="không có đường dẫn này"))
        except Exception as e: traceback.print_exc(); self._tra(500, dict(loi=str(e)))
    def do_POST(self):
        try:
            b = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
            if self.path == "/nap": return self._tra(200, xu_ly_nap(b))
            if self.path == "/duyet": return self._tra(200, xu_ly_duyet(b))
            self._tra(404, dict(loi="không có đường dẫn này"))
        except Exception as e: traceback.print_exc(); self._tra(500, dict(loi=str(e)))
    def log_message(self, fmt, *a): sys.stderr.write(f"[{dt.datetime.now():%H:%M:%S}] {fmt % a}\n")

if __name__ == "__main__":
    for st in (sys.stdout, sys.stderr): st.reconfigure(encoding="utf-8", errors="replace")   # console Windows cp1252 không in được tiếng Việt
    os.makedirs(DATA, exist_ok=True)
    port = int(sys.argv[1]) if len(sys.argv) > 1 else PORT
    print(f"Engine soát HSTT chạy tại http://127.0.0.1:{port}  ·  dữ liệu: {DATA}")
    ThreadingHTTPServer(("127.0.0.1", port), H).serve_forever()
