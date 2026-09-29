// ───────────── SO GIÁ — GIAI ĐOẠN CHỌN ĐƠN VỊ THI CÔNG / NHÀ CUNG CẤP (anh Phương 29/09/2026) ─────────────
// Phiên so giá (1 gói) → thả báo giá ứng viên → AI đọc → PTLN (↔ NS ↔ BoQ CĐT ↔ giá đã ký) → chọn đơn vị → LƯU kết quả.
// KHÔNG ghi khung (khung chỉ ghi khi nạp HĐ / HSTT). Báo giá KHÔNG thuộc hồ sơ nền.
document.head.insertAdjacentHTML("beforeend", `<style>
.sg-top{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}.sg-top label{display:flex;flex-direction:column;gap:4px;font-size:11.5px;font-weight:600;color:var(--mu);flex:1 1 320px}
.sg-top select,.sg-top input,.sg-kl textarea{font:inherit;font-size:13px;padding:8px 10px;border:1px solid var(--ru);border-radius:8px;background:#fff}
.sg-dvs{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px;margin-top:14px}
.sg-dv{border:1px solid var(--ru);border-radius:12px;padding:12px 14px;background:#fff;display:flex;flex-direction:column;gap:6px;font-size:12.5px}
.sg-dv input.ten{font:inherit;font-weight:700;font-size:13px;border:1px solid transparent;border-radius:6px;padding:3px 5px;margin:-3px -5px}.sg-dv input.ten:hover,.sg-dv input.ten:focus{border-color:var(--ru)}
.sg-dv.tat{opacity:.55}.sg-tt{font-size:11px;font-weight:700;padding:2px 8px;border-radius:999px;align-self:flex-start}
.sg-tt.CHO_AI,.sg-tt.DANG_DOC{background:#e8eefc;color:#2f5bea}.sg-tt.XONG{background:#d1e7dd;color:#0f5132}.sg-tt.LOI_AI{background:#fde2e1;color:#b42318}
.sg-up{color:#b42318;font-weight:700}.sg-dn{color:#0f5132}
td.pt-up{background:#fde2e1!important;color:#b42318;font-weight:700}td.pt-min{color:#0f5132;font-weight:800}td.pt-ref{color:#6b5bb5}
.pt-h{font-size:11px;line-height:1.25;white-space:normal;min-width:92px}.kpi.chon{outline:3px solid #12a06d}
.sg-kl{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}.sg-kl textarea{flex:1 1 380px;min-height:60px}
</style>`);
const SG = {phien: null, ds: [], p: null, kq: null, chon: null, hen: null};
const sgT = v => typeof v === "number" && isFinite(v) ? Math.round(v).toLocaleString("vi-VN") : "—";
const sgK = v => typeof v === "number" && isFinite(v) ? v.toLocaleString("vi-VN", {maximumFractionDigits: 2}) : "—";
const sgP = v => typeof v === "number" && isFinite(v) ? `<span class="${v > 0.005 ? "sg-up" : v < -0.005 ? "sg-dn" : ""}">${v > 0 ? "+" : ""}${(v * 100).toFixed(1)}%</span>` : "—";
const SG_TT = {CHO_AI: "⏳ chờ AI đọc", DANG_DOC: "🤖 AI đang đọc…", XONG: "✅ đã đọc", LOI_AI: "❌ lỗi đọc"};
const sgDA = () => encodeURIComponent($("#da").value);
function sgCSV(ten, cot, ds) {
  const q = v => { v = v == null ? "" : typeof v === "number" ? String(Math.round(v * 100) / 100) : String(v).replace(/<[^>]+>/g, ""); return /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v };
  const s = "sep=,\n" + [cot.map(c => q(c[0])), ...ds.map(x => cot.map(c => q(c[1](x))))].map(r => r.join(",")).join("\n");
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob(["﻿" + s], {type: "text/csv"}));
  a.download = `${ten}_${$("#da").value}_${new Date().toISOString().slice(0, 10)}.csv`; a.click(); URL.revokeObjectURL(a.href);
}

async function tabSoGia() {
  clearTimeout(SG.hen); const el = $("#t-sogia");
  el.innerHTML = `<div class="card"><div class="hd"><div><h3>So giá — chọn đơn vị thi công / nhà cung cấp</h3><div class="note">Mỗi gói thầu 1 <b>phiên so giá</b>: thả báo giá các ứng viên → app so với ngân sách, BoQ CĐT và giá đã ký → anh chọn đơn vị → <b>lưu kết quả</b>. Không ghi vào file khung.</div></div></div>
    <div class="pad"><div class="sg-top"><label>Phiên so giá<select id="sg-ph"><option value="">— chọn phiên —</option></select></label>
      <label style="flex:1 1 260px">Hoặc tạo phiên mới — tên gói / hạng mục<input id="sg-ten" placeholder="VD: Nhân công xây tô Block B"></label><button class="btn ok" id="sg-tao">➕ Tạo phiên</button></div></div></div>
    <div id="sg-body" class="dr"></div><div id="sg-kq" class="dr"></div>`;
  try { SG.ds = await api("/so-gia?du_an=" + sgDA()) } catch (e) { $("#sg-body").innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>`; return }
  $("#sg-ph").innerHTML += SG.ds.map(p => `<option value="${esc(p.id)}">${esc(p.ten)} · ${esc(p.ngay.slice(0, 10))} · ${p.so_dv} báo giá${p.chon ? " · ✅ chọn " + esc(p.chon) : ""}</option>`).join("");
  $("#sg-ph").onchange = e => { SG.phien = e.target.value; SG.kq = null; sgPhien() };
  $("#sg-tao").onclick = async () => {
    const ten = $("#sg-ten").value.trim(); if (!ten) return $("#sg-ten").focus();
    try { const p = await api("/so-gia/tao", {du_an: $("#da").value, ten}); SG.phien = p.id; SG.kq = null; await tabSoGia() } catch (e) { $("#sg-body").innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>` }
  };
  if (SG.phien && SG.ds.some(p => p.id === SG.phien)) { $("#sg-ph").value = SG.phien; sgPhien() } else SG.phien = null;
}

async function sgPhien() {
  clearTimeout(SG.hen); if (!SG.phien) { $("#sg-body").innerHTML = ""; $("#sg-kq").innerHTML = ""; return }
  let p; try { p = SG.p = await api(`/so-gia/phien?du_an=${sgDA()}&phien=${encodeURIComponent(SG.phien)}`) } catch (e) { $("#sg-body").innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>`; return }
  const xong = p.don_vi.filter(d => d.trang_thai === "XONG" && d.dung).length, dang = p.don_vi.filter(d => ["CHO_AI", "DANG_DOC"].includes(d.trang_thai)).length;
  $("#sg-body").innerHTML = `<div class="card"><div class="hd"><div><h3>${esc(p.ten)}</h3><div class="note">Lưu tại <code>${esc(p.thu_muc)}</code>${p.ket_qua ? ` · ✅ đã lưu kết quả ${esc(p.ket_qua.luc.replace("T", " ").slice(0, 16))}: chọn <b>${esc(p.ket_qua.chon_ten || "—")}</b>` : ""}</div></div><span class="sp"></span>
      <button class="btn ok" id="sg-ss" ${xong ? "" : "disabled"}>⚖️ So sánh ${xong} báo giá</button></div>
    <div class="pad"><div class="drop" id="sg-drop" style="padding:22px"><b>Thả báo giá của các ứng viên vào đây</b> — PDF hoặc Excel, nhiều file một lần (mỗi file 1 đơn vị)<br><span class="note">AI đọc bảng khối lượng / đơn giá (vài phút mỗi file, chạy nền — trang tự cập nhật). Báo giá lưu trong folder phiên, không vào hồ sơ nền.</span>
      <input type="file" id="sg-f" multiple accept=".pdf,.xlsx,.xlsm" hidden></div>
      <div class="sg-dvs">${p.don_vi.map(d => `<div class="sg-dv ${d.dung ? "" : "tat"}" data-dv="${esc(d.id)}">
        <input class="ten" value="${esc(d.ten)}" title="Tên đơn vị (sửa được)"><span class="note">📄 ${esc(d.ten_file)}</span>
        <span class="sg-tt ${d.trang_thai}">${SG_TT[d.trang_thai] || d.trang_thai}</span>
        ${d.trang_thai === "XONG" ? `<span>${d.so_dong} dòng có đơn giá · Σ ${sgT(d.tong)}</span>` : ""}${d.loi ? `<span class="sg-up">${esc(d.loi)}</span>` : ""}
        ${(d.canh_bao || []).map(c => `<div class="warn" style="margin:0">${esc(c)}</div>`).join("")}
        <div style="display:flex;gap:8px;align-items:center;margin-top:2px"><label class="note"><input type="checkbox" class="dung" ${d.dung ? "checked" : ""}> đưa vào so sánh</label><span class="sp" style="flex:1"></span>
          ${["XONG", "LOI_AI"].includes(d.trang_thai) ? `<button class="btn2 dl">🤖 Đọc lại</button>` : ""}</div></div>`).join("") || `<div class="empty">Chưa có báo giá nào trong phiên</div>`}</div></div></div>`;
  if (SG.da_ve === SG.phien) [...$("#sg-body").children].forEach(c => c.style.animation = "none");   // tự làm mới 6 giây/lần ⇒ không chạy lại hiệu ứng hiện dần (nhấp nháy)
  SG.da_ve = SG.phien;
  const drop = $("#sg-drop"), inp = $("#sg-f");
  drop.onclick = () => inp.click(); inp.onchange = () => sgNap([...inp.files]);
  drop.ondragover = e => { e.preventDefault(); drop.classList.add("hot") }; drop.ondragleave = () => drop.classList.remove("hot");
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove("hot"); sgNap([...e.dataTransfer.files]) };
  $("#sg-body").querySelectorAll(".sg-dv").forEach(c => {
    const dv = c.dataset.dv, gui = b => api("/so-gia/sua", {du_an: $("#da").value, phien: SG.phien, dv, ...b});
    c.querySelector(".ten").onchange = e => gui({ten: e.target.value});
    c.querySelector(".dung").onchange = async e => { await gui({dung: e.target.checked}); sgPhien() };
    const dl = c.querySelector(".dl"); if (dl) dl.onclick = async () => { await api("/so-gia/doc-lai", {du_an: $("#da").value, phien: SG.phien, dv}); sgPhien() };
  });
  $("#sg-ss").onclick = sgSoSanh;
  if (dang) SG.hen = setTimeout(sgPhien, 6000);                        // AI đang đọc ⇒ tự làm mới
  if (SG.kq && SG.kq.phien.id === SG.phien) ptVe(); else $("#sg-kq").innerHTML = "";
}
async function sgNap(files) {
  files = files.filter(f => /\.(pdf|xlsx|xlsm)$/i.test(f.name)); if (!files.length) return;
  $("#sg-drop").querySelector("b").textContent = `Đang gửi ${files.length} file…`; const loi = [];
  for (const f of files) { try { await api("/so-gia/nap", {du_an: $("#da").value, phien: SG.phien, ten: f.name, b64: await b64(f)}) } catch (e) { loi.push(f.name + ": " + e.message) } }
  await sgPhien(); if (loi.length) $("#sg-body").querySelector(".pad").insertAdjacentHTML("afterbegin", `<div class="warn">${loi.map(esc).join("<br>")}</div>`);
}
async function sgSoSanh() {
  $("#sg-kq").innerHTML = `<div class="sk" style="height:420px"></div>`;
  try { SG.kq = await api(`/so-gia/ptln?du_an=${sgDA()}&phien=${encodeURIComponent(SG.phien)}`); SG.chon = SG.kq.ket_qua && SG.kq.ket_qua.chon; ptVe() }
  catch (e) { $("#sg-kq").innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>` }
}
function ptVe() {
  const r = SG.kq, dv = r.don_vi, n = dv.length, kq = r.ket_qua;
  const o = (x, i) => { const v = x.dg_dv[i]; if (v == null) return `<td class="n note">—</td>`;
    const up = x.dg_ns && v > x.dg_ns * 1.005, mn = n > 1 && x.dv_thap.includes(+i) && x.dv_thap.length < Object.keys(x.dg_dv).length;
    const g = x.gia[i], t = ` title="${esc(g.map(y => `${y.stt}: ${y.noi_dung} = ${sgT(y.dg)} (giống ${Math.round((y.giong || 0) * 100)}%)`).join("\n"))}"`;
    const dk = x.da_ky_min && v > x.da_ky_min * 1.005 ? `<br><small class="note">cao hơn giá đã ký</small>` : "";
    return `<td class="n ${up ? "pt-up" : ""} ${mn ? "pt-min" : ""}"${t}>${sgT(v)}${g.length > 1 ? "*" : ""}${up ? `<br><small>${sgP(v / x.dg_ns - 1)} NS</small>` : ""}${dk}</td>` };
  $("#sg-kq").innerHTML = `
  <div class="kpis" style="grid-template-columns:repeat(${Math.min(n, 4)},1fr)">${dv.map(d => `<label class="kpi ${SG.chon === d.id ? "chon" : ""}" style="cursor:pointer"><span><input type="radio" name="sg-chon" value="${esc(d.id)}" ${SG.chon === d.id ? "checked" : ""}> ${d.hang ? `#${d.hang} · ` : ""}${esc(d.ten)}</span>
    <b>${sgT(d.gt_chung)}</b><div class="note">Σ trên ${r.so_chung} dòng chung${d.gt_ns_chung ? ` · ${sgP(d.gt_chung / d.gt_ns_chung - 1)} so NS` : ""}<br>
    báo giá ${sgT(d.gt_bao_gia)} · ${d.so_hang} dòng khớp NS<br>🔴 ${d.vuot_ns} dòng vượt NS (<span class="sg-up">${sgT(d.vuot_ns_tien)} đ</span>) · ${d.vuot_da_ky} dòng cao hơn giá đã ký<br>
    chênh ròng so NS ${sgT(d.chenh_ns)} đ <i>(tham khảo)</i> · ${d.so_le} dòng chưa ghép</div></label>`).join("")}</div>
  ${r.ghi_chu.map(g => `<div class="warn">${esc(g)}</div>`).join("")}
  <div class="card"><div class="hd"><div><h3>Bảng so giá theo dòng ngân sách</h3><div class="note">Đỏ = vượt ĐG NS · <b style="color:#0f5132">xanh đậm</b> = thấp nhất · "Đã ký" = đơn giá thấp nhất đã ký cho đúng công việc này (HĐ trong khung) · * = nhiều dòng báo giá cùng ghép 1 dòng NS (lấy giá cao nhất) · rê chuột xem dòng gốc.</div></div><span class="sp"></span>
    <button class="btn2" id="pt-csv">⭳ Xuất CSV</button></div><div class="right"><table class="bang">
    <thead><tr><th>Mã NS</th><th>Công việc (ngân sách)</th><th>ĐVT</th><th class="n">KL NS</th><th class="n">ĐG BoQ CĐT</th><th class="n">ĐG NS</th><th class="n">Đã ký thấp nhất</th>${dv.map(d => `<th class="n pt-h">${esc(d.ten)}</th>`).join("")}<th class="n">Thấp nhất / giá bán</th></tr></thead><tbody>
    ${r.hang.map(x => `<tr><td>${esc(x.ma_ns)}</td><td>${esc(x.noi_dung)}${x.boq_noi_dung ? `<br><span class="note">BoQ: ${esc(x.boq_noi_dung)}</span>` : ""}</td><td>${esc(x.dvt)}</td><td class="n">${sgK(x.kl_ns)}</td>
      <td class="n">${sgT(x.dg_boq)}</td><td class="n">${sgT(x.dg_ns)}</td><td class="n pt-ref" title="${esc((x.da_ky_ai || []).join(", "))} · ${x.da_ky_so_hd} HĐ đã ký, ${sgT(x.da_ky_min)}–${sgT(x.da_ky_max)}">${sgT(x.da_ky_min)}${x.da_ky_so_hd ? `<br><small>${x.da_ky_so_hd} HĐ</small>` : ""}</td>
      ${dv.map(d => o(x, d.i)).join("")}<td class="n">${x.dg_boq && x.thap_nhat ? Math.round(x.thap_nhat / x.dg_boq * 100) + "%" : "—"}</td></tr>`).join("")}</tbody>
    <tfoot><tr><td colspan="7">Σ ĐG × KL NS — ${r.so_chung} dòng mọi đơn vị cùng báo</td>${dv.map(d => `<td class="n">${sgT(d.gt_chung)}</td>`).join("")}<td></td></tr></tfoot></table></div></div>
  <div class="card"><div class="hd"><div><h3>Theo mã ngân sách — gói này còn bao nhiêu NS</h3><div class="note">Đã giao = Σ HĐ chi phí đã ký trong khung. HĐ đơn giá: KL tạm tính nên "đã giao" có thể cao hơn thực tế.</div></div></div><div class="right"><table class="bang">
    <thead><tr><th>Mã NS</th><th>Tên</th><th class="n">Ngân sách</th><th class="n">Đã giao</th><th class="n">NS còn lại</th>${dv.map(d => `<th class="n pt-h">${esc(d.ten)}</th>`).join("")}</tr></thead><tbody>
    ${r.theo_ma.map(x => `<tr><td>${esc(x.ma_ns)}</td><td>${esc(x.ten_ns)}</td><td class="n">${sgT(x.ns)}</td><td class="n" title="${esc(x.hd_da_ky.join(", "))}">${sgT(x.da_giao)}</td>
      <td class="n ${x.con_lai != null && x.con_lai < 0 ? "sg-up" : ""}">${sgT(x.con_lai)}</td>${dv.map(d => { const v = x.gt[d.i]; return `<td class="n ${v != null && x.con_lai != null && v > x.con_lai ? "pt-up" : ""}">${sgT(v)}</td>` }).join("")}</tr>`).join("")}</tbody></table></div></div>
  ${r.le.length ? `<div class="card"><div class="hd"><div><h3>Dòng báo giá chưa ghép được ngân sách (${r.le.length})</h3><div class="note">Không đặt cạnh nhau được — anh xem riêng (ĐVT lệch NS, chưa có mã NS, khoản trọn gói…).</div></div></div><div class="right"><table class="bang">
    <thead><tr><th>Đơn vị</th><th>STT</th><th>Nội dung</th><th>ĐVT</th><th class="n">ĐG</th><th>Mã NS</th><th>Lý do</th></tr></thead><tbody>
    ${r.le.map(l => `<tr><td>${esc(dv[l.i].ten)}</td><td>${esc(l.stt)}</td><td>${esc(l.noi_dung)}</td><td>${esc(l.dvt)}</td><td class="n">${sgT(l.dg)}</td><td>${esc(l.ma_ns || "—")}</td><td class="note">${esc(l.ly_do || "")}</td></tr>`).join("")}</tbody></table></div></div>` : ""}
  <div class="card"><div class="hd"><div><h3>Kết luận chọn thầu</h3><div class="note">Chỉ LƯU kết quả (sổ phiên + file PTLN Excel trong folder phiên) — không ghi khung. Khung chỉ ghi khi nạp hợp đồng.</div></div></div>
    <div class="pad sg-kl"><textarea id="sg-gc" placeholder="Lý do chọn / điều kiện đàm phán lại (vd: chống nóng xuống 10.000 như giá đã ký)…">${esc(kq ? kq.ghi_chu : "")}</textarea>
    <button class="btn ok" id="sg-luu">💾 Lưu kết quả${SG.chon ? "" : " (chưa chọn đơn vị)"}</button><div id="sg-da-luu" class="note">${kq ? `Đã lưu ${esc(kq.luc.replace("T", " ").slice(0, 16))} · <code>${esc(kq.file)}</code>` : ""}</div></div></div>
  <p class="note src">Nguồn: báo giá AI đọc (phiên so giá) · N2 ngân sách · N5 BoQ CĐT · N6/N7 HĐ đã ký (giá tham chiếu). Ghép dòng theo mã NS + ĐVT + nội dung. Không ghi khung.</p>`;
  $("#sg-kq").querySelectorAll("input[name=sg-chon]").forEach(c => c.onchange = () => { SG.chon = c.value; ptVe() });
  $("#sg-luu").onclick = async () => {
    try { const k = await api("/so-gia/luu", {du_an: $("#da").value, phien: SG.phien, chon: SG.chon, ghi_chu: $("#sg-gc").value}); SG.kq.ket_qua = k;
          $("#sg-da-luu").innerHTML = `✅ Đã lưu ${esc(k.luc.replace("T", " ").slice(0, 16))} · <code>${esc(k.file)}</code>` } catch (e) { $("#sg-da-luu").innerHTML = `<span class="sg-up">${esc(e.message)}</span>` }
  };
  $("#pt-csv").onclick = () => sgCSV("PTLN_" + (r.phien.ten || ""), [["Mã NS", x => x.ma_ns], ["Công việc NS", x => x.noi_dung], ["ĐVT", x => x.dvt], ["KL NS", x => x.kl_ns], ["ĐG BoQ CĐT", x => x.dg_boq],
    ["ĐG NS", x => x.dg_ns], ["Đã ký thấp nhất", x => x.da_ky_min], ...dv.map(d => [d.ten, x => x.dg_dv[d.i]]), ["Thấp nhất", x => x.thap_nhat]], r.hang);
}
