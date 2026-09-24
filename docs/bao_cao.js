// ĐẦU RA (chỉ đọc) — Tổng quan · Hợp đồng · Bill thanh toán · Báo cáo tài chính · Đối tác.
// Mọi con số lấy từ GET /du-lieu = các sheet báo cáo của file khung Excel (R1…R7, 90_Check). Trang KHÔNG tự tính lại số tài chính.
let DL = null;
const tien = v => typeof v === "number" && isFinite(v) ? Math.round(v).toLocaleString("vi-VN") : "—";
const ty = v => { if (typeof v !== "number" || !isFinite(v)) return "—"; const a = Math.abs(v);
  return a >= 1e9 ? (v / 1e9).toLocaleString("vi-VN", {maximumFractionDigits: 2}) + " tỷ" : a >= 1e6 ? (v / 1e6).toLocaleString("vi-VN", {maximumFractionDigits: 1}) + " tr" : tien(v) };
const pc = v => typeof v === "number" && isFinite(v) ? (v * 100).toLocaleString("vi-VN", {maximumFractionDigits: 1}) + "%" : "—";
const dd = v => v ? String(v).slice(0, 10).split("-").reverse().join("/") : "—";
const ng = v => typeof v === "number" && v < -0.5 ? "neg" : "";
const bar = (v, max, cls = "") => `<div class="bar"><i class="${cls}" style="width:${Math.max(0, Math.min(100, max ? v / max * 100 : 0)).toFixed(1)}%"></i></div>`;
const CV = k => (DL.cvr[k] || {}).gia_tri;
const NHOM = {"B.1": "Nhân công & thầu phụ", "B.2": "Vật tư — nhà cung cấp", "B.3": "Chi phí gián tiếp công trường", "B.4": "Dự phòng"};
const MUC_CK = k => k.startsWith("ĐẠT") ? ["ok", "Đạt"] : k.startsWith("LỆCH") ? ["er", "Lệch số"] : k.startsWith("CÓ") ? ["er", "Rủi ro chi phí"] : ["wa", k.startsWith("CHƯA") ? "Chưa tin được" : "Thiếu hồ sơ"];

async function taiDL(force) {
  const da = $("#da").value;
  if (DL && !force && DL._da === da) return DL;
  DL = await api("/du-lieu?du_an=" + encodeURIComponent(da)); DL._da = da; return DL;
}
function chan(d) {                                                  // dòng nguồn dưới mỗi màn — minh bạch số lấy từ đâu
  return `<p class="note src">Nguồn: file khung Excel của dự án (sheet R1…R7, N1–N9, 90_Check) · cut-off ${dd(d.du_an.moc)} · file cập nhật ${esc(d.cap_nhat)}.
    Trang chỉ đọc, không tự tính lại số — muốn sửa số thì sửa ở Excel hoặc nạp HSTT.</p>`;
}
function xuatCSV(ten, cot, ds) {
  const q = v => { v = v == null ? "" : typeof v === "number" ? String(Math.round(v * 100) / 100) : String(v).replace(/<[^>]+>/g, ""); return /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v };
  const s = "sep=,\n" + [cot.map(c => q(c.t)), ...ds.map(x => cot.map(c => q(c.csv ? c.csv(x) : x[c.k])))].map(r => r.join(",")).join("\n");
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob(["﻿" + s], {type: "text/csv"}));
  a.download = `${ten}_${(DL.du_an.ma || "")}_${new Date().toISOString().slice(0, 10)}.csv`; a.click(); URL.revokeObjectURL(a.href);
}
// Bảng dùng chung: tìm kiếm · lọc nhanh · phân trang · xuất CSV · bấm dòng xem chi tiết
function taoBang(id, cfg) {
  const st = {q: "", loc: "", tr: 0}, N = cfg.moi_trang || 20, el = $("#" + id);
  el.innerHTML = `<div class="card"><div class="hd"><div><h3>${cfg.tieu}</h3><div class="note">${cfg.phu}</div></div><span class="sp"></span>
    <input class="tim" placeholder="🔎 ${cfg.goi_y}">${cfg.loc ? `<div class="seg">${cfg.loc.ds.map(([v, t]) => `<button data-v="${v}" class="${v === "" ? "on" : ""}">${t}</button>`).join("")}</div>` : ""}
    <button class="btn2" data-x="csv">⭳ Xuất Excel (CSV)</button></div>
    <div class="right"><table class="bang"></table></div><div class="pager"></div></div><div class="ct" style="margin-top:14px"></div>${chan(DL)}`;
  const loc = () => cfg.ds().filter(x => (!st.loc || cfg.loc.f(x, st.loc)) && (!st.q || cfg.tim.some(k => String(x[k] ?? "").toLowerCase().includes(st.q))));
  const ve = () => {
    const ds = loc(), so_tr = Math.max(1, Math.ceil(ds.length / N)); st.tr = Math.min(st.tr, so_tr - 1);
    const trang = ds.slice(st.tr * N, st.tr * N + N);
    el.querySelector(".bang").innerHTML = `<thead><tr>${cfg.cot.filter(c => !c.an).map(c => `<th class="${c.n ? "n" : ""}">${c.t}</th>`).join("")}</tr></thead><tbody>` +
      (trang.map((x, i) => `<tr class="click" data-i="${i}">${cfg.cot.filter(c => !c.an).map(c => `<td class="${c.n ? "n" : ""}">${c.v ? c.v(x) : esc(x[c.k])}</td>`).join("")}</tr>`).join("")
       || `<tr><td colspan="${cfg.cot.filter(c => !c.an).length}" class="empty">Không có dòng nào khớp</td></tr>`) + "</tbody>" + (cfg.tong ? `<tfoot>${cfg.tong(ds)}</tfoot>` : "");
    el.querySelector(".pager").innerHTML = `<span class="note">${ds.length ? `${st.tr * N + 1}–${Math.min(ds.length, st.tr * N + N)} / ${ds.length} dòng` : "0 dòng"}</span><span class="sp"></span>
      <button class="btn2" data-p="-1" ${st.tr ? "" : "disabled"}>‹ Trước</button><span class="note">Trang ${st.tr + 1}/${so_tr}</span><button class="btn2" data-p="1" ${st.tr < so_tr - 1 ? "" : "disabled"}>Sau ›</button>`;
    el.querySelectorAll(".bang tbody tr.click").forEach(tr => tr.onclick = () => {
      el.querySelectorAll(".bang tr").forEach(x => x.classList.toggle("sel", x === tr));
      const box = el.querySelector(".ct"); box.innerHTML = cfg.click(trang[+tr.dataset.i]); box.scrollIntoView({behavior: "smooth", block: "start"});
    });
    el.querySelectorAll(".pager [data-p]").forEach(b => b.onclick = () => { st.tr += +b.dataset.p; ve() });
  };
  el.querySelector(".tim").oninput = e => { st.q = e.target.value.trim().toLowerCase(); st.tr = 0; ve() };
  el.querySelectorAll(".seg button").forEach(b => b.onclick = () => { el.querySelectorAll(".seg button").forEach(x => x.classList.toggle("on", x === b)); st.loc = b.dataset.v; st.tr = 0; ve() });
  el.querySelector("[data-x=csv]").onclick = () => xuatCSV(cfg.file, cfg.cot, loc());
  ve();
}
const kv = (k, v, cls = "") => `<div class="kv"><span>${k}</span><b class="${cls}">${v}</b></div>`;
const gtHD = h => h.don_gia && !h.gia_tri ? `<span class="note">HĐ đơn giá</span>` : tien(h.gia_tri);

// ───────────── HỢP ĐỒNG (thẻ chi tiết dùng chung cho Bill / Đối tác) ─────────────
function theHD(h) {
  const bills = DL.bill.filter(b => b.ma_hd === h.ma_hd), dh = h.den_han || {};
  return `<div class="card"><div class="hd"><div><h3>${esc(h.so_hd || h.ma_hd)}</h3><div class="note">${esc(h.doi_tac)} · ${esc(h.noi_dung || "")}</div></div><span class="sp"></span>
    <span class="chip ${h.ben === "CĐT" ? "ok" : ""}">${h.ben === "CĐT" ? "Doanh thu (CĐT)" : esc(h.loai)}</span><span class="chip">${esc(h.trang_thai)}</span>
    ${h.ho_so_thieu ? '<span class="chip wa">Thiếu hồ sơ</span>' : ""}${h.canh_bao ? `<span class="chip er">${esc(h.canh_bao)}</span>` : ""}</div>
    <div class="body"><div class="left">
      ${kv("Ngày ký", dd(h.ngay_ky))}${kv("Dạng HĐ", esc(h.dang || "—"))}${kv("Gói thầu", esc(h.ma_goi || "—"))}${h.phu_luc ? kv("Phụ lục", h.phu_luc) : ""}
      ${kv("VAT", pc(h.vat))}${kv("Tạm ứng · TT mỗi đợt · quyết toán", `${pc(h.pct_tu)} · ${pc(h.pct_tt)} · ${pc(h.pct_qt)}`)}<hr>
      ${kv("Giá trị HĐ (trước VAT)", gtHD(h))}${kv("Thực hiện lũy kế", tien(h.thuc_hien))}${h.pct != null ? bar(h.thuc_hien, h.gia_tri, h.pct > 1 ? "er" : "") + kv("% hoàn thành", pc(h.pct), h.pct > 1 ? "neg" : "") : ""}
      ${kv("Tiền thanh toán lũy kế", tien(h.tt_lk))}${kv("Giữ lại đang giữ", tien(h.giu_lai))}${kv("Tạm ứng chưa hoàn", tien(h.tu_con))}${h.khau_tru ? kv("Khấu trừ lũy kế", tien(h.khau_tru)) : ""}
      ${h.gia_tri ? kv("Còn lại theo HĐ", tien(h.con_lai), ng(h.con_lai)) : ""}${dh.han_gl ? kv("Hạn trả giữ lại", esc(String(dh.han_gl).includes("T00") ? dd(dh.han_gl) : dh.han_gl)) : ""}
      ${h.ho_so_thieu ? `<div class="warn"><b>Hồ sơ còn thiếu</b><br>${esc(h.ho_so_thieu)}</div>` : ""}</div>
    <div class="right"><h4>Các đợt thanh toán</h4><table><tr><th>Đợt</th><th>Ngày HSTT</th><th class="n">Sản lượng kỳ</th><th class="n">Lũy kế</th><th class="n">Tiền TT đợt</th><th>Hạn TT</th></tr>${bills.map(b =>
      `<tr><td>Đợt ${b.dot}</td><td>${dd(b.ngay)}</td><td class="n">${tien(b.san_luong_ky)}</td><td class="n">${tien(b.san_luong_lk)}</td><td class="n"><b>${tien(b.tien_tt)}</b></td><td>${dd(b.han)}</td></tr>`).join("")
      || `<tr><td colspan="6" class="empty">Chưa có đợt thanh toán</td></tr>`}</table>
      <h4 style="margin-top:14px">Dòng hợp đồng (${h.dong.length})</h4><table><tr><th>STT</th><th>Nội dung</th><th>ĐVT</th><th class="n">KL HĐ</th><th class="n">Đơn giá</th><th class="n">KL lũy kế</th><th>Tiến độ KL</th></tr>${h.dong.map(x => {
      const co = typeof x.kl_hd === "number" && x.kl_hd > 0, vuot = co && x.kl_lk > x.kl_hd * 1.0001;
      return `<tr><td>${esc(x.stt)}</td><td>${esc(x.noi_dung)}${x.ngoai ? ' <span class="chip wa">ngoài HĐ</span>' : ""}</td><td>${esc(x.dvt)}</td><td class="n">${fmt(x.kl_hd)}</td>
        <td class="n">${tien(x.don_gia)}</td><td class="n ${vuot ? "neg" : ""}">${fmt(Math.round(x.kl_lk * 100) / 100)}</td><td style="min-width:90px">${co ? bar(x.kl_lk, x.kl_hd, vuot ? "er" : "") : ""}</td></tr>` }).join("")}</table></div></div></div>`;
}
async function tabHopDong() {
  const d = await taiDL();
  taoBang("t-hd", {tieu: "Hợp đồng", phu: `${d.hop_dong.length} hợp đồng — CĐT (doanh thu) và đội / thầu phụ / NCC (chi phí)`, goi_y: "Tìm đối tác, số HĐ, nội dung…", file: "HopDong",
    tim: ["doi_tac", "so_hd", "noi_dung", "loai", "ma_hd"], ds: () => d.hop_dong,
    loc: {ds: [["", "Tất cả"], ["CĐT", "CĐT"], ["Đội thi công", "Đội"], ["Thầu phụ", "Thầu phụ"], ["Nhà cung cấp", "NCC"], ["thieu", "Thiếu hồ sơ"]],
          f: (h, v) => v === "CĐT" ? h.ben === "CĐT" : v === "thieu" ? !!h.ho_so_thieu : h.loai === v && h.ben !== "CĐT"},
    cot: [{t: "Số hợp đồng", v: h => `<b>${esc(h.so_hd || h.ma_hd)}</b><div class="note">ký ${dd(h.ngay_ky)}</div>`, csv: h => h.so_hd},
          {t: "Đối tác", k: "doi_tac"}, {t: "Nội dung", k: "noi_dung"},
          {t: "Loại", v: h => `<span class="chip ${h.ben === "CĐT" ? "ok" : ""}">${h.ben === "CĐT" ? "Doanh thu" : esc(h.loai)}</span>`, csv: h => h.ben === "CĐT" ? "CĐT" : h.loai},
          {t: "Giá trị HĐ", n: 1, v: gtHD, csv: h => h.gia_tri}, {t: "Thực hiện LK", n: 1, v: h => tien(h.thuc_hien), csv: h => h.thuc_hien},
          {t: "Hoàn thành", v: h => h.pct == null ? '<span class="note">—</span>' : `<div class="pp">${bar(h.thuc_hien, h.gia_tri, h.pct > 1 ? "er" : "")}<span class="${h.pct > 1 ? "neg" : ""}">${pc(h.pct)}</span></div>`, csv: h => h.pct},
          {t: "Trạng thái", v: h => `<span class="chip ok">${esc(h.trang_thai)}</span>${h.ho_so_thieu ? ' <span class="chip wa">thiếu hồ sơ</span>' : ""}${h.canh_bao ? ` <span class="chip er">${esc(h.canh_bao)}</span>` : ""}`,
           csv: h => [h.trang_thai, h.ho_so_thieu && "Thiếu: " + h.ho_so_thieu, h.canh_bao].filter(Boolean).join(" | ")}],
    tong: ds => { const cp = ds.filter(h => h.ben !== "CĐT"); return `<tr><td colspan="4">Cộng chi phí (${cp.length} HĐ đối tác đang lọc)</td><td class="n">${tien(cp.reduce((a, h) => a + h.gia_tri, 0))}</td>
      <td class="n">${tien(cp.reduce((a, h) => a + h.thuc_hien, 0))}</td><td colspan="2"></td></tr>` },
    click: theHD});
}
// ───────────── BILL THANH TOÁN ─────────────
async function tabBill() {
  const d = await taiDL(), ds = d.bill.slice().sort((a, b) => String(b.ngay).localeCompare(String(a.ngay)) || a.doi_tac.localeCompare(b.doi_tac));
  const thang = [...new Set(ds.map(b => String(b.ngay || "").slice(0, 7)).filter(Boolean))].slice(0, 4);
  taoBang("t-bill", {tieu: "Bill thanh toán theo kỳ", phu: `${ds.length} đợt — sản lượng, tạm ứng/hoàn ứng, tiền thanh toán từng đợt`, goi_y: "Tìm đối tác, số HĐ…", file: "BillThanhToan",
    tim: ["doi_tac", "so_hd", "ma_hd"], ds: () => ds,
    loc: {ds: [["", "Tất cả"], ["CĐT", "Doanh thu"], ["DT", "Chi phí"], ...thang.map(t => [t, "T" + t.split("-").reverse().join("/")])],
          f: (b, v) => v === "CĐT" ? b.ben === "CĐT" : v === "DT" ? b.ben !== "CĐT" : String(b.ngay || "").startsWith(v)},
    cot: [{t: "Đối tác · số HĐ", v: b => `<b>${esc(b.doi_tac)}</b><div class="note">${esc(b.so_hd)}</div>`, csv: b => b.doi_tac}, {t: "Số HĐ", k: "so_hd", an: 1},
          {t: "Đợt", v: b => `<span class="chip ${b.ben === "CĐT" ? "ok" : ""}">Đợt ${b.dot}</span>`, csv: b => b.dot}, {t: "Ngày HSTT", v: b => dd(b.ngay), csv: b => dd(b.ngay)},
          {t: "Giá trị HĐ", n: 1, v: b => b.don_gia && !b.gia_tri_hd ? '<span class="note">HĐ đơn giá</span>' : tien(b.gia_tri_hd), csv: b => b.gia_tri_hd},
          {t: "Sản lượng kỳ", n: 1, v: b => tien(b.san_luong_ky), csv: b => b.san_luong_ky}, {t: "Tổng sản lượng", n: 1, v: b => tien(b.san_luong_lk), csv: b => b.san_luong_lk},
          {t: "Tạm ứng / hoàn ứng", n: 1, v: b => b.tam_ung || b.hoan_ung ? `${b.tam_ung ? "+" + tien(b.tam_ung) : ""}${b.tam_ung && b.hoan_ung ? "<br>" : ""}${b.hoan_ung ? tien(b.hoan_ung) : ""}` : '<span class="note">—</span>', csv: b => `${Math.round(b.tam_ung)}/${Math.round(b.hoan_ung)}`},
          {t: "Tiền TT đợt này", n: 1, v: b => `<b>${tien(b.tien_tt)}</b>`, csv: b => b.tien_tt},
          {t: "% HĐ", v: b => b.pct == null ? '<span class="note">—</span>' : `<div class="pp">${bar(b.san_luong_lk, b.gia_tri_hd, b.pct > 1 ? "er" : "")}<span>${pc(b.pct)}</span></div>`, csv: b => b.pct},
          {t: "Tình trạng", v: b => `<span class="chip ${b.tinh_trang === "Đã trả" ? "ok" : ""}">${b.tinh_trang}</span><div class="note">hạn ${dd(b.han)}</div>${b.qua_app ? '<span class="chip ok">ghi qua app</span>' : ""}`,
           csv: b => `${b.tinh_trang} · hạn ${dd(b.han)}`}],
    tong: ds => [["CĐT", "Doanh thu — CĐT"], ["DT", "Chi phí — đối tác"]].map(([k, t]) => { const r = ds.filter(b => (b.ben === "CĐT") === (k === "CĐT")); return r.length ?
      `<tr><td colspan="4">${t}: ${r.length} đợt đang lọc</td><td class="n">${tien(r.reduce((a, b) => a + b.san_luong_ky, 0))}</td><td></td><td></td>
      <td class="n">${tien(r.reduce((a, b) => a + b.tien_tt, 0))}</td><td colspan="2"></td></tr>` : "" }).join(""),
    click: b => theHD(d.hop_dong.find(h => h.ma_hd === b.ma_hd))});
}
// ───────────── ĐỐI TÁC ─────────────
async function tabDoiTac() {
  const d = await taiDL();
  taoBang("t-dt", {tieu: "Đối tác CĐT / NTP / NCC", phu: `${d.doi_tac.length} đối tác — tổng hợp theo hợp đồng trong file khung`, goi_y: "Tìm tên, mã đối tác…", file: "DoiTac",
    tim: ["ten", "ma", "loai_ten"], ds: () => d.doi_tac,
    loc: {ds: [["", "Tất cả"], ["CĐT", "CĐT"], ["ĐTC", "Đội thi công"], ["NTP", "Thầu phụ"], ["NCC", "NCC"]], f: (x, v) => x.loai === v},
    cot: [{t: "Đối tác", v: x => `<b>${esc(x.ten)}</b>`, csv: x => x.ten}, {t: "Mã viết tắt", v: x => `<span class="chip">${esc(x.ma)}</span>`, csv: x => x.ma},
          {t: "Loại", v: x => `<span class="chip ${x.loai === "CĐT" ? "ok" : ""}">${esc(x.loai_ten)}</span>`, csv: x => x.loai_ten}, {t: "Số HĐ", n: 1, k: "so_hd"},
          {t: "Giá trị HĐ", n: 1, v: x => x.gia_tri ? tien(x.gia_tri) : '<span class="note">HĐ đơn giá</span>', csv: x => x.gia_tri},
          {t: "Thực hiện LK", n: 1, v: x => tien(x.thuc_hien), csv: x => x.thuc_hien}, {t: "Tiền TT lũy kế", n: 1, v: x => tien(x.tt_lk), csv: x => x.tt_lk},
          {t: "Giữ lại", n: 1, v: x => tien(x.giu_lai), csv: x => x.giu_lai}, {t: "Hồ sơ", v: x => x.ho_so_thieu ? `<span class="chip wa">${x.ho_so_thieu} HĐ thiếu</span>` : '<span class="chip ok">đủ</span>', csv: x => x.ho_so_thieu}],
    click: x => x.ma_hd.map(m => theHD(d.hop_dong.find(h => h.ma_hd === m))).join('<div style="height:12px"></div>') || '<div class="card"><div class="empty">Chưa có hợp đồng</div></div>'});
}
const lienKet = ds => !ds.length ? '<span class="note">chưa giao thầu</span>' :
  `<span class="chip" title="${esc(ds.map(h => h.doi_tac + " — " + h.so_hd).join("\n"))}">${ds.length} HĐ</span> <span class="note">${esc(ds.slice(0, 2).map(h => h.doi_tac.replace(/^(Tổ đội|Công ty (CP|TNHH)( MTV| TM DV XD| SX-TM| TV TK XD)?)\s*/i, "")).join(", "))}${ds.length > 2 ? ` +${ds.length - 2}` : ""}</span>`;
// ───────────── BÁO CÁO TÀI CHÍNH (Hàng A doanh thu · B chi phí · C lợi nhuận) ─────────────
async function tabBCTC() {
  const d = await taiDL(), T = d.tong_lai_lo["TỔNG DỰ ÁN"] || {}, el = $("#t-bctc");
  const S = (ds, k) => ds.reduce((a, x) => a + (typeof x[k] === "number" ? x[k] : 0), 0);
  const cot = [{t: "Tên hạng mục", k: "ten"}, {t: "Mã NS", k: "ma_ns"}, {t: "Nhóm", k: "nhom"}, {t: "Đối tác liên kết", csv: x => x.hop_dong.map(h => h.doi_tac).join("; ")},
    {t: "Doanh thu phân bổ", k: "dt_pb"}, {t: "NS hiện hành", k: "ns"}, {t: "GT HĐ đã ký", k: "cam_ket"}, {t: "Đã thực hiện", k: "thuc_hien"},
    {t: "Còn lại NS", csv: x => x.ns - x.thuc_hien}, {t: "Chi phí còn lại", k: "con_lai"}, {t: "EAC", k: "eac"}, {t: "NS − EAC", k: "ns_tru_eac"}, {t: "LN dự kiến", k: "ln_dk"}, {t: "Cảnh báo", k: "canh_bao"}];
  const ve = () => {
    const q = (el.querySelector(".tim").value || "").toLowerCase();
    const ok = x => !q || [x.ten, x.ma_ns, ...x.hop_dong.map(h => h.doi_tac), ...x.hop_dong.map(h => h.so_hd)].some(s => String(s ?? "").toLowerCase().includes(q));
    const dong = (x, lv) => `<tr class="${lv}"><td>${lv === "l2" ? `${esc(x.ten)}<div class="note">${esc(x.ma_ns)}</div>` : x.ten}${x.canh_bao ? `<div><span class="chip er">${esc(x.canh_bao)}</span></div>` : ""}</td>
      <td>${x.hop_dong ? lienKet(x.hop_dong) : ""}</td>
      <td class="n">${tien(x.dt_pb)}</td><td class="n">${tien(x.ns)}</td><td class="n">${tien(x.cam_ket)}</td><td class="n">${tien(x.thuc_hien)}</td>
      <td class="n ${ng(x.ns - x.thuc_hien)}">${typeof x.ns === "number" ? tien(x.ns - x.thuc_hien) : "—"}</td><td class="n">${tien(x.eac)}${Math.abs(x.ns_tru_eac || 0) >= 1 ? `<div class="note ${ng(x.ns_tru_eac)}">NS−EAC ${tien(x.ns_tru_eac)}</div>` : ""}</td><td class="n ${ng(x.ln_dk)}">${tien(x.ln_dk)}</td></tr>`;
    let h = `<thead><tr><th>Tên hạng mục</th><th>Hợp đồng liên kết</th><th class="n">Doanh thu phân bổ</th><th class="n">Phân bổ dự trù (NS)</th><th class="n">GT HĐ đã ký</th>
      <th class="n">Đã thực hiện</th><th class="n">Còn lại NS</th><th class="n">EAC</th><th class="n">LN dự kiến</th></tr></thead><tbody>
      <tr class="l0 a"><td colspan="9">HÀNG A — DOANH THU (trước VAT)</td></tr>
      <tr class="l2"><td>Hợp đồng với CĐT — bản gốc<div class="note">R1 mục 1</div></td><td>${esc((d.hop_dong.find(x => x.ben === "CĐT") || {}).so_hd || "")}</td><td class="n">${tien(CV("1"))}</td><td colspan="6"></td></tr>
      <tr class="l2"><td>Phụ lục / phát sinh với CĐT<div class="note">R1 mục 2</div></td><td></td><td class="n">${tien(CV("2"))}</td><td colspan="6"></td></tr>
      <tr class="l1"><td>Cộng doanh thu điều chỉnh</td><td class="note">đã nghiệm thu ${tien(CV("3a"))} · ${pc(CV("3a") / CV("3"))}</td><td class="n">${tien(CV("3"))}</td><td colspan="6"></td></tr>
      <tr class="l0 b"><td colspan="9">HÀNG B — CHI PHÍ (theo mã ngân sách)</td></tr>`;
    for (const nh of Object.keys(NHOM)) {
      const all = d.ns.filter(x => x.nhom === nh), hs = all.filter(ok); if (!hs.length) continue;
      h += dong({ten: `${nh} — ${NHOM[nh]}`, dt_pb: S(all, "dt_pb"), ns: S(all, "ns"), cam_ket: S(all, "cam_ket"), thuc_hien: S(all, "thuc_hien"), eac: S(all, "eac"), ns_tru_eac: S(all, "ns_tru_eac"), ln_dk: S(all, "ln_dk")}, "l1");
      h += hs.map(x => dong(x, "l2")).join("");
    }
    const B = {ten: "Cộng chi phí (Hàng B)", dt_pb: S(d.ns, "dt_pb"), ns: S(d.ns, "ns"), cam_ket: S(d.ns, "cam_ket"), thuc_hien: S(d.ns, "thuc_hien"), eac: S(d.ns, "eac"), ns_tru_eac: S(d.ns, "ns_tru_eac"), ln_dk: S(d.ns, "ln_dk")};
    h += dong(B, "l1 tong") + "</tbody>";
    el.querySelector(".bang").innerHTML = h;
  };
  const A = CV("3"), Bns = T.ns, eac = CV("10"), lnk = T.ln_kh, lnd = CV("11"), tt6 = CV("6"), dp7 = CV("7");
  el.innerHTML = `<div class="card"><div class="hd"><div><h3>Báo cáo tài chính dự án</h3><div class="note">${esc(d.du_an.ten)} · CĐT ${esc(d.du_an.cdt)} · cut-off ${dd(d.du_an.moc)}</div></div><span class="sp"></span>
    <input class="tim" placeholder="🔎 Tìm hạng mục, đối tác, số HĐ…"><button class="btn2" data-x="csv">⭳ Xuất Excel (CSV)</button><button class="btn2" onclick="window.print()">🖨 In / PDF</button></div>
    <div class="right"><table class="bang bctc"></table></div>
    <div class="hangc"><div><span>Tổng Hàng A — Doanh thu</span><b>${tien(A)}</b></div><div><span>Tổng Hàng B — Ngân sách chi phí</span><b>${tien(Bns)}</b></div>
      <div><span>Hàng C — Lợi nhuận kế hoạch (A − B)</span><b class="${ng(lnk)}">${tien(lnk)}</b><em>${pc(A ? lnk / A : null)}</em></div>
      <div><span>EAC — chi phí cuối dự kiến</span><b>${tien(eac)}</b><em>NS − EAC ${tien(Bns - eac)}</em></div>
      <div class="hl"><span>Lợi nhuận dự kiến (A − EAC)</span><b class="${ng(lnd)}">${tien(lnd)}</b><em>${pc(CV("12"))}</em></div></div>
    ${!tt6 && !dp7 ? `<div class="warn" style="margin:0 16px 14px">Trích trước (R1 mục 6) và dự phòng rủi ro (mục 7) đang = 0 ⇒ lợi nhuận dự kiến là <b>trước</b> trích trước/dự phòng; 90_Check #27 báo CVR “chưa tin được”.</div>` : ""}</div>${chan(d)}`;
  el.querySelector(".tim").oninput = ve;
  el.querySelector("[data-x=csv]").onclick = () => xuatCSV("BaoCaoTaiChinh", cot, d.ns);
  ve();
}
// ───────────── TỔNG QUAN (dashboard) ─────────────
function bieuDoTien(ds) {
  let n = ds.length; while (n > 1 && !ds[n - 1].thu && !ds[n - 1].chi) n--; ds = ds.slice(0, Math.min(ds.length, n + 1));
  const W = 920, H = 260, L = 70, B = 30, T = 18, cw = (W - L - 10) / ds.length;
  const hi = Math.max(1, ...ds.map(x => Math.max(x.thu, x.chi, x.lk))), lo = Math.min(0, ...ds.map(x => x.lk));
  const y = v => T + (H - T - B) * (hi - v) / (hi - lo), y0 = y(0);
  const moc = String(DL.du_an.moc || "").slice(0, 7).split("-").reverse().join("/");
  let s = `<svg viewBox="0 0 ${W} ${H}" class="chart" role="img" aria-label="Dòng tiền theo tháng">`;
  for (let i = 0; i <= 5; i++) { const v = lo + (hi - lo) * i / 5, yy = y(v); s += `<line x1="${L}" x2="${W - 10}" y1="${yy}" y2="${yy}" class="gl"/><text x="${L - 6}" y="${yy + 4}" class="ax" text-anchor="end">${ty(v)}</text>` }
  ds.forEach((x, i) => { const cx = L + i * cw, bw = Math.min(20, cw / 2 - 3);
    if (x.thang === moc) s += `<rect x="${cx}" y="${T}" width="${cw}" height="${H - T - B}" class="moc"/><text x="${cx + cw / 2}" y="${T - 5}" class="ax" text-anchor="middle">cut-off</text>`;
    s += `<rect x="${cx + cw / 2 - bw - 1}" y="${y(x.thu)}" width="${bw}" height="${y0 - y(x.thu)}" class="thu"><title>${x.thang} · Thu ${tien(x.thu)}</title></rect>`;
    s += `<rect x="${cx + cw / 2 + 1}" y="${y(x.chi)}" width="${bw}" height="${y0 - y(x.chi)}" class="chi"><title>${x.thang} · Chi ${tien(x.chi)}</title></rect>`;
    s += `<text x="${cx + cw / 2}" y="${H - 10}" class="ax" text-anchor="middle">${x.thang}</text>` });
  s += `<line x1="${L}" x2="${W - 10}" y1="${y0}" y2="${y0}" class="zero"/><polyline class="lk" points="${ds.map((x, i) => `${L + i * cw + cw / 2},${y(x.lk)}`).join(" ")}"/>`;
  ds.forEach((x, i) => s += `<circle cx="${L + i * cw + cw / 2}" cy="${y(x.lk)}" r="3.5" class="lkd"><title>${x.thang} · Lũy kế ròng ${tien(x.lk)}</title></circle>`);
  const am = ds.reduce((a, x) => x.lk < a.lk ? x : a, ds[0]);
  return s + `</svg><div class="legend"><span><i class="thu"></i>Thu từ CĐT</span><span><i class="chi"></i>Chi cho đối tác</span><span><i class="lk"></i>Lũy kế ròng</span>
    ${am && am.lk < 0 ? `<span>Âm sâu nhất: <b class="neg">${tien(am.lk)}</b> (${am.thang})</span>` : ""}<span class="note">${esc((ds.find(x => x.ghi) || {}).ghi || "")}</span></div>`;
}
async function tongQuan(force) {
  const el = $("#t-bc"); el.innerHTML = `<div class="card"><div class="empty">Đang đọc file khung…</div></div>`;
  let d; try { d = await taiDL(force) } catch (e) { el.innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>`; return }
  const T = d.tong_lai_lo["TỔNG DỰ ÁN"] || {}, A = CV("3"), nt = CV("3a"), cp = CV("8"), eac = CV("10"), ln = CV("11"), cam = CV("5"), ns = T.ns;
  const ck = d.kiem.map(c => ({...c, m: MUC_CK(c.ket_luan)})), dem = t => ck.filter(c => c.m[1] === t).length;
  const doiTac = d.hop_dong.filter(h => h.ben !== "CĐT"), ttDT = doiTac.reduce((a, h) => a + h.tt_lk, 0);
  // Điểm cần chú ý — chỉ nêu SỰ KIỆN + SỐ + NGUỒN, xếp theo mức ảnh hưởng tiền
  const y = [];
  d.ns.filter(x => x.canh_bao).forEach(x => y.push({m: "er", loai: "Ngân sách", t: `${x.ten} (${x.ma_ns}) — ${x.canh_bao}`, so: /LỖ/.test(x.canh_bao) && typeof x.ln_dk === "number" ? x.ln_dk : x.ns_tru_eac, nguon: /LỖ/.test(x.canh_bao) ? "R2 · R7" : "R2"}));
  d.ns.filter(x => typeof x.chenh_ln === "number" && x.chenh_ln < -1 && !x.canh_bao).forEach(x => y.push({m: "wa", loai: "Lợi nhuận", t: `${x.ten} — LN dự kiến giảm so kế hoạch`, so: x.chenh_ln, nguon: "R7"}));
  d.hop_dong.filter(h => h.canh_bao).forEach(h => y.push({m: "er", loai: "Hợp đồng", t: `${h.doi_tac} — ${h.canh_bao}`, so: null, nguon: "R3"}));
  const wd = {}; d.workdone_canh_bao.forEach(x => (wd[x.canh_bao] = wd[x.canh_bao] || []).push(x.ten));
  Object.entries(wd).forEach(([k, v]) => y.push({m: "wa", loai: "Workdone", t: `${v.length} nhóm công việc: ${k}`, chi_tiet: v.join(" · "), so: null, nguon: "R6"}));
  const thieu = d.hop_dong.filter(h => h.ho_so_thieu); if (thieu.length) y.push({m: "wa", loai: "Hồ sơ", t: `${thieu.length} hợp đồng còn thiếu hồ sơ`, chi_tiet: thieu.map(h => h.doi_tac).join(" · "), so: null, nguon: "R3"});
  ck.filter(c => c.m[1] === "Lệch số").forEach(c => y.push({m: "er", loai: "Đối chiếu", t: `90_Check #${c.so}: ${c.ten}`, so: c.lech, nguon: "90_Check"}));
  y.sort((a, b) => (a.m === "er" ? 0 : 1) - (b.m === "er" ? 0 : 1) || Math.abs(b.so || 0) - Math.abs(a.so || 0));
  const nhom = Object.keys(NHOM).map(k => { const r = d.ns.filter(x => x.nhom === k), S = f => r.reduce((a, x) => a + (x[f] || 0), 0);
    return {k, ns: S("ns"), th: S("thuc_hien"), eac: S("eac")} });
  const mxN = Math.max(1, ...nhom.map(x => Math.max(x.ns, x.eac)));
  const goi = d.goi.filter(g => g["Mã gói"]);
  el.innerHTML = `
  <div class="card hero"><div><div class="note up">Báo cáo tổng quan dự án</div><h2>${esc(d.du_an.ten)}</h2>
      <div class="note">Chủ đầu tư <b>${esc(d.du_an.cdt)}</b> · Cut-off <b>${dd(d.du_an.moc)}</b> · ${d.hop_dong.length} hợp đồng · ${d.bill.length} đợt thanh toán · file cập nhật ${esc(d.cap_nhat)}</div></div>
    <span class="sp"></span><div class="tin"><div class="note">Độ tin cậy số liệu (90_Check)</div><div><span class="chip ok">${dem("Đạt")}/${ck.length} đạt</span>
      ${dem("Lệch số") ? `<span class="chip er">${dem("Lệch số")} lệch số</span>` : ""}${dem("Rủi ro chi phí") ? `<span class="chip er">${dem("Rủi ro chi phí")} rủi ro chi phí</span>` : ""}
      ${dem("Thiếu hồ sơ") ? `<span class="chip wa">${dem("Thiếu hồ sơ")} thiếu hồ sơ</span>` : ""}${dem("Chưa tin được") ? '<span class="chip wa">CVR chưa tin được</span>' : ""}</div></div>
    <div class="acts0"><button class="btn2" onclick="tongQuan(true)">⟳ Làm mới</button><button class="btn2" onclick="window.print()">🖨 In / PDF</button></div></div>
  <div class="kpis">
    <div class="kpi"><span>Doanh thu hợp đồng (trước VAT)</span><b>${ty(A)}</b><div class="note">HĐ gốc ${ty(CV("1"))} · phụ lục ${ty(CV("2"))}</div></div>
    <div class="kpi"><span>Đã nghiệm thu với CĐT</span><b>${ty(nt)}</b>${bar(nt, A, "ok")}<div class="note">${pc(nt / A)} doanh thu · còn ${ty(A - nt)}</div></div>
    <div class="kpi"><span>Chi phí đã thực hiện</span><b>${ty(cp)}</b>${bar(cp, eac, "er")}<div class="note">${pc(cp / eac)} EAC · HĐ đã ký chưa TH ${ty(cam)}</div></div>
    <div class="kpi ${ln < 0 ? "xau" : "tot"}"><span>Lợi nhuận dự kiến (DT − EAC)</span><b class="${ng(ln)}">${ty(ln)} <small>${pc(CV("12"))}</small></b>
      <div class="note">Kế hoạch (DT − NS) ${ty(T.ln_kh)} · ${pc(A ? T.ln_kh / A : null)} · chênh <span class="${ng(ln - T.ln_kh)}">${ty(ln - T.ln_kh)}</span></div></div></div>
  <div class="g2">
    <div class="card"><div class="hd"><h3>Giá trị — Chi phí (CVR)</h3><span class="sp"></span><span class="chip">R1_CVR</span></div><div class="pad">
      ${[["Doanh thu điều chỉnh", A, "ok"], ["Ngân sách chi phí hiện hành", ns, ""], ["EAC — chi phí cuối dự kiến", eac, eac > ns ? "er" : ""]].map(([t, v, c]) =>
        `<div class="cmp"><span>${t}</span>${bar(v, Math.max(A, ns, eac), c)}<b>${tien(v)}</b></div>`).join("")}
      <div class="cmp2"><div><span>Tiến độ doanh thu</span><b>${pc(nt / A)}</b><div class="note">nghiệm thu / doanh thu</div></div>
        <div><span>Tiến độ chi phí</span><b>${pc(cp / eac)}</b><div class="note">đã thực hiện / EAC</div></div>
        <div><span>Đã cam kết</span><b>${pc((cp + cam) / eac)}</b><div class="note">(thực hiện + HĐ chưa TH) / EAC</div></div></div>
      <table class="mini">${["3", "3a", "4", "5", "6", "7", "8", "9", "10", "11"].map(k => d.cvr[k] ? `<tr><td class="note">${k}</td><td>${esc(d.cvr[k].ten)}</td><td class="n ${ng(d.cvr[k].gia_tri)}">${tien(d.cvr[k].gia_tri)}</td></tr>` : "").join("")}</table></div></div>
    <div class="card"><div class="hd"><h3>Công nợ & tiền</h3><span class="sp"></span><span class="chip">R1 mục 15–19 · R3</span></div><div class="pad">
      <div class="tien2"><div><h4>Với Chủ đầu tư</h4>${kv("Được thanh toán lũy kế", tien(CV("15")))}${kv("CĐT đang giữ lại", tien(CV("16")))}${kv("Tạm ứng CĐT chưa thu hồi", tien(CV("17")))}</div>
        <div><h4>Với đối tác</h4>${kv("Đã thanh toán lũy kế", tien(ttDT))}${kv("Mình đang giữ lại", tien(CV("19")))}${kv("Tạm ứng chưa hoàn", tien(CV("18")))}</div></div>
      <div class="net"><span>Chênh tiền theo HSTT (CĐT trả − trả đối tác)</span><b class="${ng(CV("15") - ttDT)}">${tien(CV("15") - ttDT)}</b></div>
      <p class="note">Tính theo hồ sơ thanh toán đã duyệt — chưa phải tiền đã về / đã chi thật (khung chưa nhập ngày trả).</p></div></div></div>
  <div class="card"><div class="hd"><h3>Dòng tiền theo tháng</h3><span class="sp"></span><span class="chip">R5_DongTien</span></div><div class="pad">${bieuDoTien(d.dong_tien)}</div></div>
  <div class="g2">
    <div class="card"><div class="hd"><h3>Ngân sách theo nhóm chi phí</h3><span class="sp"></span><span class="chip">R2</span></div><div class="pad">
      ${nhom.map(x => `<div class="nh"><div class="t"><b>${x.k}</b>&nbsp;${NHOM[x.k]}<span class="sp"></span><span class="note">NS ${ty(x.ns)} · TH ${ty(x.th)} · EAC <b class="${x.eac > x.ns + 1 ? "neg" : ""}">${ty(x.eac)}</b></span></div>
        <div class="bar2"><i style="width:${(x.ns / mxN * 100).toFixed(1)}%" class="ns"></i><i style="width:${(x.th / mxN * 100).toFixed(1)}%" class="th"></i><em style="left:${(x.eac / mxN * 100).toFixed(1)}%"></em></div>
        <div class="note">${pc(x.ns ? x.th / x.ns : null)} ngân sách đã dùng · NS − EAC <span class="${ng(x.ns - x.eac)}">${tien(x.ns - x.eac)}</span></div></div>`).join("")}
      <div class="legend"><span><i class="ns"></i>Ngân sách</span><span><i class="th"></i>Đã thực hiện</span><span><i class="eac"></i>EAC</span></div></div></div>
    <div class="card"><div class="hd"><h3>Lãi lỗ theo gói thầu</h3><span class="sp"></span><span class="chip">R7_LaiLo_TheoGoi</span></div><div class="right"><table>
      <tr><th>Gói</th><th class="n">DT phân bổ</th><th class="n">EAC</th><th class="n">LN dự kiến</th><th class="n">%</th><th class="n">so KH</th></tr>
      ${goi.map(g => `<tr><td><b>${esc(g["Mã gói"])}</b> <span class="note">${esc(String(g["Tên gói thầu"] || "").slice(0, 44))}</span></td><td class="n">${ty(g["Doanh thu phân bổ"])}</td><td class="n">${ty(g["EAC"])}</td>
        <td class="n ${ng(g["LN dự kiến"])}">${ty(g["LN dự kiến"])}</td><td class="n ${ng(g["% LN dự kiến"])}">${pc(g["% LN dự kiến"])}</td><td class="n ${ng(g["LN dự kiến − kế hoạch"])}">${ty(g["LN dự kiến − kế hoạch"])}</td></tr>`).join("")}</table>
      <p class="note">Gói có doanh thu phân bổ = 0 (vận chuyển, gián tiếp…) là chi phí chung — lỗ trên gói do cách phân bổ; xem lợi nhuận ở tổng dự án.</p></div></div></div>
  <div class="card"><div class="hd"><h3>Điểm cần chú ý</h3><span class="note">${y.length} mục · nêu sự kiện, số tiền và sheet nguồn — xếp theo mức ảnh hưởng</span></div><div class="right"><table>
    <tr><th>Mức</th><th>Loại</th><th>Nội dung</th><th class="n">Số tiền ảnh hưởng</th><th>Nguồn</th></tr>
    ${y.map(x => `<tr><td><span class="chip ${x.m}">${x.m === "er" ? "Cần xử lý" : "Cần xem"}</span></td><td>${x.loai}</td><td>${esc(x.t)}${x.chi_tiet ? `<div class="note">${esc(x.chi_tiet)}</div>` : ""}</td>
      <td class="n ${ng(x.so)}">${x.so == null ? "" : tien(x.so)}</td><td><span class="chip">${x.nguon}</span></td></tr>`).join("")}</table></div></div>
  <div class="card"><details><summary class="pad">90_Check — ${ck.length} phép kiểm của file khung (bấm để xem)</summary><div class="right"><table>
    <tr><th>#</th><th>Phép kiểm</th><th class="n">Vế trái</th><th class="n">Vế phải</th><th>Kết luận</th></tr>
    ${ck.map(c => `<tr><td>${c.so}</td><td>${esc(c.ten)}</td><td class="n">${tien(c.trai)}</td><td class="n">${tien(c.phai)}</td><td><span class="chip ${c.m[0]}">${esc(c.ket_luan)}</span></td></tr>`).join("")}</table></div></details></div>
  ${chan(d)}`;
}
const TAB_DAU_RA = {bc: tongQuan, hd: tabHopDong, bill: tabBill, bctc: tabBCTC, dt: tabDoiTac};
