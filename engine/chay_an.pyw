"""Chạy engine ẨN (không cửa sổ) — dùng cho Task Scheduler 'Engine soat HSTT' (tự bật khi đăng nhập Windows).
Đã có engine chạy ở cổng 8765 ⇒ thoát luôn (không chạy 2 bản). Log: <DATA>\_CAU_HINH\log\engine.log."""
import os, sys, json, socket, runpy
D = os.path.dirname(os.path.abspath(__file__))
try:
    socket.create_connection(("127.0.0.1", 8765), timeout=1).close(); sys.exit(0)
except OSError: pass
p = os.path.join(D, "cau_hinh_may.json"); cfg = json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}
log = os.path.join(cfg.get("DATA", D), "_CAU_HINH", "log"); os.makedirs(log, exist_ok=True)
sys.stdout = sys.stderr = open(os.path.join(log, "engine.log"), "a", encoding="utf-8", buffering=1)
os.chdir(D); sys.path.insert(0, D); sys.argv = [os.path.join(D, "app.py"), "8765"]
runpy.run_path(os.path.join(D, "app.py"), run_name="__main__")
