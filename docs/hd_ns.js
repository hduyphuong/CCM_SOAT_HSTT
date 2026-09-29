// ───────────── SO GIÁ — nghiệp vụ ĐỘC LẬP (anh Phương 29/09/2026): CHỈ ĐỌC khung, không ghi, chưa nối luồng HSTT ─────────────
// Chế độ 1 · Kiểm HĐ ↔ Ngân sách (1 HĐ, cộng dồn các HĐ đã ký)   ·   Chế độ 2 · PTLN (nhiều đơn vị nhận thầu ↔ NS ↔ BoQ CĐT)
document.head.insertAdjacentHTML("beforeend", `<style>
.sg-m{display:inline-block;font-size:11px;font-weight:700;padding:2px 8px;border-radius:999px;white-space:nowrap}
.sg-m.DO{background:#fde2e1;color:#b42318}.sg-m.VANG{background:#fff3cd;color:#8a5a00}.sg-m.XAM{background:#eef0f4;color:#5b6474}.sg-m.XANH{background:#d1e7dd;color:#0f5132}
tr.sg-DO td{background:#fff6f5}tr.sg-VANG td{background:#fffbeb}
.sg-ns{font-size:11.5px;color:var(--mu);max-width:260px}.sg-ld{font-size:11.5px;max-width:320px}
.sg-top{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}.sg-top label{display:flex;flex-direction:column;gap:4px;font-size:11.5px;font-weight:600;color:var(--mu);flex:1 1 380px}
.sg-top select{font:inherit;font-size:13px;padding:8px 10px;border:1px solid var(--ru);border-radius:8px;background:#fff}
.sg-up{color:#b42318;font-weight:700}.sg-dn{color:#0f5132}
</style>`);
const SG = {che_do: "hd", ds: null, kq: null, loc: ""};
const SG_MUC = {DO: "🔴 Vượt ĐG", VANG: "🟡 Cần xem", XAM: "⚪ Chưa so được", XANH: "🟢 Đạt"};
const sgM = m => `<span class="sg-m ${m}">${SG_MUC[m] || m}</span>`;
const sgT = v => typeof v === "number" && isFinite(v) ? Math.round(v).toLocaleString("vi-VN") : "—";
const sgK = v => typeof v === "number" && isFinite(v) ? v.toLocaleString("vi-VN", {maximumFractionDigits: 2}) : "—";
const sgP = v => typeof v === "number" && isFinite(v) ? `<span class="${v > 0.005 ? "sg-up" : v < -0.005 ? "sg-dn" : ""}">${v > 0 ? "+" : ""}${(v * 100).toFixed(1)}%</span>` : "—";
function sgCSV(ten, cot, ds) {
  const q = v => { v = v == null ? "" : typeof v === "number" ? String(Math.round(v * 100) / 100) : String(v).replace(/<[^>]+>/g, ""); return /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v };
  const s = "sep=,\n" + [cot.map(c => q(c[0])), ...ds.map(x => cot.map(c => q(c[1](x))))].map(r => r.join(",")).join("\n");
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob(["﻿" + s], {type: "text/csv"}));
  a.download = `${ten}_${$("#da").value}_${new Date().toISOString().slice(0, 10)}.csv`; a.click(); URL.revokeObjectURL(a.href);
}

async function tabSoGia() {
  const el = $("#t-sogia");
  el.innerHTML = `<div class="card"><div class="hd"><div><h3>So giá</h3><div class="note">Nghiệp vụ độc lập — <b>chỉ đọc</b> file khung, không ghi gì, không ảnh hưởng hồ sơ thanh toán.</div></div><span class="sp"></span>
    <div class="seg" id="sg-cd"><button data-v="hd" class="${SG.che_do === "hd" ? "on" : ""}">Kiểm HĐ ↔ Ngân sách</button><button data-v="ptln" class="${SG.che_do === "ptln" ? "on" : ""}">PTLN · so giá thầu</button></div></div>
    <div class="pad" id="sg-chon"></div></div><div id="sg-kq" class="dr"></div>`;
  el.querySelectorAll("#sg-cd button").forEach(b => b.onclick = () => { SG.che_do = b.dataset.v; SG.kq = null; tabSoGia() });
  if (SG.che_do === "ptln") return typeof tabPTLN === "function" ? tabPTLN() : ($("#sg-chon").innerHTML = `<div class="empty">Chế độ PTLN đang làm — dùng chung lõi ghép dòng với Kiểm HĐ ↔ NS.</div>`);
  $("#sg-chon").innerHTML = `<div class="sk" style="height:44px"></div>`;
  try { SG.ds = await api("/hd-ns-ds?du_an=" + encodeURIComponent($("#da").value)) } catch (e) { $("#sg-chon").innerHTML = `<div class="empty">${esc(e.message)}</div>`; return }
  const kh = SG.ds.filter(x => x.nguon === "khung"), nen = SG.ds.filter(x => x.nguon === "nen");
  const op = x => `<option value="${x.nguon}|${esc(x.ma)}">${esc(x.ten)}${x.so_hd ? " · " + esc(x.so_hd) : ""} · ${x.so_dong} dòng · ${sgT(x.gia_tri)}</option>`;
  $("#sg-chon").innerHTML = `<div class="sg-top"><label>Hợp đồng / báo giá cần kiểm<select id="sg-hd"><option value="">— chọn —</option>
      <optgroup label="HĐ đã có trong khung (${kh.length})">${kh.map(op).join("")}</optgroup>
      ${nen.length ? `<optgroup label="Hồ sơ nền AI đã đọc, CHƯA nhập khung — kiểm trước khi ký (${nen.length})">${nen.map(op).join("")}</optgroup>` : ""}</select></label>
    <button class="btn ok" id="sg-chay">🔎 Kiểm</button></div>
    <p class="note" style="margin-top:8px">HĐ / báo giá mới chưa có trong danh sách ⇒ thả file vào <b>Hồ sơ nền</b> (AI đọc bảng đơn giá) rồi quay lại đây — không cần duyệt nhập khung.</p>`;
  $("#sg-chay").onclick = sgChay; $("#sg-hd").onchange = sgChay;
  if (SG.kq) sgVe();
}
async function sgChay() {
  const v = $("#sg-hd").value; if (!v) return;
  const [ng, ma] = v.split("|"); $("#sg-kq").innerHTML = `<div class="sk" style="height:420px"></div>`;
  try { SG.kq = await api(`/hd-ns?du_an=${encodeURIComponent($("#da").value)}&${ng === "nen" ? "nen_id" : "ma_hd"}=${encodeURIComponent(ma)}`); SG.loc = ""; sgVe() }
  catch (e) { $("#sg-kq").innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>` }
}
function sgVe() {
  const r = SG.kq, h = r.hd, t = r.tong, d = t.dem || {};
  const dsLoc = () => r.dong.filter(x => !SG.loc || x.muc === SG.loc);
  const ve_dong = () => { $("#sg-dong").innerHTML = `<thead><tr><th>STT</th><th>Nội dung HĐ</th><th>ĐVT</th><th class="n">KL HĐ</th><th class="n">ĐG HĐ</th><th class="n">ĐG NS</th><th class="n">Chênh</th><th>Dòng NS so sánh</th><th>Mã NS</th><th>Kết quả</th></tr></thead><tbody>` +
    (dsLoc().map(x => `<tr class="sg-${x.muc}"><td>${esc(x.stt)}</td><td>${esc(x.noi_dung)}${x.pham_vi === "NGOAI_HD" ? ` <span class="note">(ngoài HĐ)</span>` : ""}</td><td>${esc(x.dvt)}</td>
      <td class="n">${sgK(x.kl)}</td><td class="n">${sgT(x.dg)}</td><td class="n">${sgT(x.dg_ns)}</td><td class="n">${sgP(x.chenh_dg)}</td>
      <td class="sg-ns">${x.ns_noi_dung ? esc(x.ns_noi_dung) + `<br><i>${esc(x.cach_so || "")}${x.do_giong != null && x.cach_so === "đúng dòng NS" ? " · giống " + Math.round(x.do_giong * 100) + "%" : ""}${x.ns_ma_cv ? " · " + esc(x.ns_ma_cv) : ""}</i>` : "—"}</td>
      <td>${esc(x.ma_ns || "—")}${x.nguon_ma === "luật (gợi ý)" ? `<br><i class="note">gợi ý</i>` : ""}</td><td class="sg-ld">${sgM(x.muc)}<br>${esc(x.ly_do || "")}${x.tien_vuot ? `<br><b class="sg-up">+${sgT(x.tien_vuot)} đ</b>` : ""}</td></tr>`).join("")
     || `<tr><td colspan="10" class="empty">Không có dòng nào</td></tr>`) + "</tbody>" };
  $("#sg-kq").innerHTML = `
  <div class="kpis">
    <div class="kpi"><span>${esc(h.ma_hd || h.loai || "Hồ sơ")} · ${esc(h.doi_tac || "")}</span><b>${sgT(h.gia_tri)}</b><div class="note">${esc(h.so_hd || "")} ${h.dang_hd ? "· " + esc(h.dang_hd) : ""} · ${t.so_dong} dòng · so được ${sgT(t.tien_co_ns)} đ</div></div>
    <div class="kpi"><span>🔴 Vượt đơn giá NS</span><b>${d.DO || 0} <small>dòng</small></b><div class="note">tiền vượt ${sgT(t.tien_vuot_dg)} đ (ĐG vượt × KL HĐ)</div></div>
    <div class="kpi"><span>🟡 Cần anh xem</span><b>${d.VANG || 0} <small>dòng</small></b><div class="note">khớp dòng chưa chắc / so bình quân mã</div></div>
    <div class="kpi"><span>Mã NS cộng dồn vượt</span><b>${t.ma_ns_vuot} <small>mã</small></b><div class="note">⚪ ${d.XAM || 0} dòng chưa so được · 🟢 ${d.XANH || 0} dòng đạt</div></div>
  </div>
  ${r.ghi_chu.map(g => `<div class="warn">${esc(g)}</div>`).join("")}
  <div class="card"><div class="hd"><div><h3>Từng dòng HĐ ↔ dòng ngân sách</h3><div class="note">Ghép theo mã NS + ĐVT + nội dung giống nhất. Đỏ = chắc chắn vượt đơn giá · Vàng = anh xem lại dòng so sánh.</div></div><span class="sp"></span>
    <div class="seg" id="sg-loc">${[["", "Tất cả"], ["DO", "Đỏ"], ["VANG", "Vàng"], ["XAM", "Xám"], ["XANH", "Xanh"]].map(([v, tx]) => `<button data-v="${v}" class="${SG.loc === v ? "on" : ""}">${tx}${v ? ` ${d[v] || 0}` : ""}</button>`).join("")}</div>
    <button class="btn2" id="sg-csv">⭳ Xuất CSV</button></div><div class="right"><table class="bang" id="sg-dong"></table></div></div>
  <div class="card"><div class="hd"><div><h3>Giá trị theo mã ngân sách</h3><div class="note">HĐ này + các HĐ chi phí đã ký khác trong khung, so với NS hiện hành.</div></div></div><div class="right"><table class="bang">
    <thead><tr><th>Mã NS</th><th>Tên</th><th class="n">Ngân sách</th><th class="n">HĐ này</th><th class="n">HĐ đã ký khác</th><th class="n">Cộng dồn</th><th class="n">% NS</th><th>Kết quả</th></tr></thead><tbody>
    ${r.gt.map(x => `<tr class="sg-${x.muc}"><td>${esc(x.ma_ns)}</td><td>${esc(x.ten_ns)}</td><td class="n">${sgT(x.gt_ns)}</td><td class="n">${sgT(x.gt_nay)}</td>
      <td class="n" title="${esc(x.hd_khac.join(", "))}">${sgT(x.gt_khac)} <span class="note">(${x.hd_khac.length} HĐ)</span></td><td class="n">${sgT(x.gt_tong)}</td>
      <td class="n">${x.gt_ns ? Math.round(x.gt_tong / x.gt_ns * 100) + "%" : "—"}</td><td class="sg-ld">${sgM(x.muc)} ${esc(x.ly_do)}</td></tr>`).join("")}</tbody></table></div></div>
  <div class="card"><div class="hd"><div><h3>Khối lượng theo mã NS · ĐVT</h3><div class="note">HĐ đơn giá: KL trong HĐ thường là tạm tính ⇒ cờ vàng để cân nhắc, không phải lỗi.</div></div></div><div class="right"><table class="bang">
    <thead><tr><th>Mã NS</th><th>ĐVT</th><th class="n">KL ngân sách</th><th class="n">KL HĐ này</th><th class="n">KL HĐ khác</th><th class="n">Cộng dồn</th><th>Kết quả</th></tr></thead><tbody>
    ${r.kl.map(x => `<tr class="sg-${x.muc}"><td>${esc(x.ma_ns)}<br><span class="note">${esc(x.ten_ns)}</span></td><td>${esc(x.dvt)}</td><td class="n">${sgK(x.kl_ns)}</td><td class="n">${sgK(x.kl_nay)}</td>
      <td class="n">${sgK(x.kl_khac)}</td><td class="n">${sgK(x.kl_tong)}</td><td class="sg-ld">${sgM(x.muc)} ${esc(x.ly_do)}</td></tr>`).join("")}</tbody></table></div></div>
  <p class="note src">Nguồn: file khung — N2 (ngân sách) · N6/N7 (HĐ đội/NCC) · N1 (nhóm → mã NS); hồ sơ nền: bảng AI đọc. Ngưỡng: ĐG vượt &gt; ${r.nguong.dg * 100}% · độ giống tối thiểu ${Math.round(r.nguong.khop * 100)}%. Trang chỉ đọc, không ghi khung.</p>`;
  ve_dong();
  $("#sg-kq").querySelectorAll("#sg-loc button").forEach(b => b.onclick = () => { SG.loc = b.dataset.v; $("#sg-kq").querySelectorAll("#sg-loc button").forEach(x => x.classList.toggle("on", x === b)); ve_dong() });
  $("#sg-csv").onclick = () => sgCSV(`KiemHD_NS_${h.ma_hd || "hoso"}`, [["STT", x => x.stt], ["Nội dung", x => x.noi_dung], ["Phạm vi", x => x.pham_vi], ["ĐVT", x => x.dvt], ["KL HĐ", x => x.kl], ["ĐG HĐ", x => x.dg],
    ["ĐG NS", x => x.dg_ns], ["Chênh %", x => x.chenh_dg == null ? "" : Math.round(x.chenh_dg * 1000) / 10], ["Dòng NS so sánh", x => x.ns_noi_dung], ["Cách so", x => x.cach_so], ["Mã NS", x => x.ma_ns],
    ["Kết quả", x => SG_MUC[x.muc]], ["Lý do", x => x.ly_do], ["Tiền vượt", x => x.tien_vuot]], dsLoc());
}

// ───────────── PTLN — so giá NHIỀU đơn vị nhận thầu ↔ Ngân sách ↔ BoQ CĐT (lõi ghép dòng dùng chung) ─────────────
document.head.insertAdjacentHTML("beforeend", `<style>
.pt-ds{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:6px 14px;max-height:240px;overflow:auto;padding:4px 2px}
.pt-ds label{display:flex;gap:8px;align-items:flex-start;font-size:12.5px;cursor:pointer}.pt-ds label small{color:var(--mu)}
td.pt-up{background:#fde2e1!important;color:#b42318;font-weight:700}td.pt-min{color:#0f5132;font-weight:800}
.pt-h{font-size:11px;line-height:1.25;white-space:normal;min-width:92px}
</style>`);
const PT = {chon: new Set(), kq: null};
async function tabPTLN() {
  const el = $("#sg-chon"); el.innerHTML = `<div class="sk" style="height:120px"></div>`;
  try { SG.ds = SG.ds || await api("/hd-ns-ds?du_an=" + encodeURIComponent($("#da").value)) } catch (e) { el.innerHTML = `<div class="empty">${esc(e.message)}</div>`; return }
  const nhom = (tieu, ds) => ds.length ? `<div class="note" style="margin:8px 0 4px"><b>${tieu}</b></div><div class="pt-ds">${ds.map(x => { const v = `${x.nguon}|${x.ma}`;
    return `<label><input type="checkbox" value="${esc(v)}" ${PT.chon.has(v) ? "checked" : ""}><span>${esc(x.ten)}<br><small>${esc(x.so_hd || "")} · ${x.so_dong} dòng · ${sgT(x.gia_tri)}</small></span></label>` }).join("")}</div>` : "";
  el.innerHTML = `${nhom("Báo giá / HĐ ở Hồ sơ nền — chưa nhập khung", SG.ds.filter(x => x.nguon === "nen"))}${nhom("HĐ đã ký trong khung", SG.ds.filter(x => x.nguon === "khung"))}
    <div class="sg-top" style="margin-top:10px"><span class="note" id="pt-dem">Đã chọn ${PT.chon.size} đơn vị</span><span class="sp" style="flex:1"></span><button class="btn2" id="pt-bo">Bỏ chọn</button><button class="btn ok" id="pt-chay">⚖️ So sánh</button></div>
    <p class="note" style="margin-top:8px">Báo giá mới ⇒ thả vào <b>Hồ sơ nền</b> (loại Báo giá, AI đọc bảng đơn giá) rồi tick ở đây. Chọn ≥ 2 đơn vị để xếp hạng.</p>`;
  el.querySelectorAll(".pt-ds input").forEach(c => c.onchange = () => { c.checked ? PT.chon.add(c.value) : PT.chon.delete(c.value); $("#pt-dem").textContent = `Đã chọn ${PT.chon.size} đơn vị` });
  $("#pt-bo").onclick = () => { PT.chon.clear(); tabPTLN() };
  $("#pt-chay").onclick = async () => {
    if (!PT.chon.size) return; $("#sg-kq").innerHTML = `<div class="sk" style="height:420px"></div>`;
    try { PT.kq = await api(`/ptln?du_an=${encodeURIComponent($("#da").value)}&ds=${encodeURIComponent([...PT.chon].join(","))}`); ptVe() }
    catch (e) { $("#sg-kq").innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>` }
  };
  if (PT.kq) ptVe();
}
function ptVe() {
  const r = PT.kq, dv = r.don_vi, n = dv.length;
  const o = (x, i) => { const v = x.dg_dv[i]; if (v == null) return `<td class="n note">—</td>`;
    const up = x.dg_ns && v > x.dg_ns * 1.005, mn = n > 1 && x.dv_thap.includes(i) && x.dv_thap.length < Object.keys(x.dg_dv).length;
    const g = x.gia[i], nhieu = g.length > 1 ? ` title="${esc(g.map(y => y.stt + ": " + y.noi_dung + " = " + sgT(y.dg)).join("\n"))}"` : "";
    return `<td class="n ${up ? "pt-up" : ""} ${mn ? "pt-min" : ""}"${nhieu}>${sgT(v)}${g.length > 1 ? "*" : ""}${up ? `<br><small>${sgP(v / x.dg_ns - 1)}</small>` : ""}</td>` };
  $("#sg-kq").innerHTML = `
  <div class="kpis" style="grid-template-columns:repeat(${Math.min(n, 4)},1fr)">${dv.map(d => `<div class="kpi"><span>${d.hang ? `#${d.hang} · ` : ""}${esc(d.ten)}</span><b>${sgT(d.gt_chung)}</b>
    <div class="note">trên ${r.so_chung} dòng chung${d.gt_ns_chung ? ` · ${sgP(d.gt_chung / d.gt_ns_chung - 1)} so NS` : ""}<br>báo ${d.so_hang} dòng NS · 🔴 ${d.vuot_ns} vượt NS · ${d.so_le} dòng không ghép được</div></div>`).join("")}</div>
  ${r.ghi_chu.map(g => `<div class="warn">${esc(g)}</div>`).join("")}
  <div class="card"><div class="hd"><div><h3>Bảng so giá theo dòng ngân sách</h3><div class="note">Đỏ = vượt đơn giá NS · <b style="color:#0f5132">xanh đậm</b> = thấp nhất · * = đơn vị có nhiều dòng cùng ghép 1 dòng NS (lấy giá cao nhất, rê chuột xem).</div></div><span class="sp"></span>
    <button class="btn2" id="pt-csv">⭳ Xuất CSV</button></div><div class="right"><table class="bang">
    <thead><tr><th>Mã NS</th><th>Công việc (ngân sách)</th><th>ĐVT</th><th class="n">KL NS</th><th class="n">ĐG BoQ CĐT</th><th class="n">ĐG NS</th>${dv.map(d => `<th class="n pt-h">${esc(d.ten)}</th>`).join("")}<th class="n">Thấp nhất / giá bán</th></tr></thead><tbody>
    ${r.hang.map(x => `<tr><td>${esc(x.ma_ns)}</td><td>${esc(x.noi_dung)}${x.boq_noi_dung ? `<br><span class="note">BoQ: ${esc(x.boq_noi_dung)}</span>` : ""}</td><td>${esc(x.dvt)}</td><td class="n">${sgK(x.kl_ns)}</td>
      <td class="n">${sgT(x.dg_boq)}</td><td class="n">${sgT(x.dg_ns)}</td>${dv.map(d => o(x, d.i)).join("")}<td class="n">${x.dg_boq && x.thap_nhat ? Math.round(x.thap_nhat / x.dg_boq * 100) + "%" : "—"}</td></tr>`).join("")}</tbody>
    <tfoot><tr><td colspan="6">Σ ĐG × KL NS — ${r.so_chung} dòng mọi đơn vị cùng báo</td>${dv.map(d => `<td class="n">${sgT(d.gt_chung)}</td>`).join("")}<td></td></tr></tfoot></table></div></div>
  ${r.le.length ? `<div class="card"><div class="hd"><div><h3>Dòng báo giá chưa ghép được dòng ngân sách (${r.le.length})</h3><div class="note">Không đặt cạnh nhau được — anh xem riêng (ĐVT lệch NS, chưa có mã NS, khoản trọn gói…).</div></div></div><div class="right"><table class="bang">
    <thead><tr><th>Đơn vị</th><th>STT</th><th>Nội dung</th><th>ĐVT</th><th class="n">ĐG</th><th>Mã NS</th><th>Lý do</th></tr></thead><tbody>
    ${r.le.map(l => `<tr><td>${esc(dv[l.i].ten)}</td><td>${esc(l.stt)}</td><td>${esc(l.noi_dung)}</td><td>${esc(l.dvt)}</td><td class="n">${sgT(l.dg)}</td><td>${esc(l.ma_ns || "—")}</td><td class="sg-ld">${esc(l.ly_do || "")}</td></tr>`).join("")}</tbody></table></div></div>` : ""}
  <p class="note src">Nguồn: N2 (ngân sách) · N5 (BoQ CĐT) · N6/N7 (HĐ trong khung) · bảng AI đọc ở Hồ sơ nền. Ghép dòng: cùng lõi với Kiểm HĐ ↔ NS. Trang chỉ đọc, không ghi khung.</p>`;
  $("#pt-csv").onclick = () => sgCSV("PTLN", [["Mã NS", x => x.ma_ns], ["Công việc NS", x => x.noi_dung], ["ĐVT", x => x.dvt], ["KL NS", x => x.kl_ns], ["ĐG BoQ CĐT", x => x.dg_boq], ["ĐG NS", x => x.dg_ns],
    ...dv.map(d => [d.ten, x => x.dg_dv[d.i]]), ["Thấp nhất", x => x.thap_nhat]], r.hang);
}
