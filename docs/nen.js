// HỒ SƠ NỀN — anh CHỈ NẠP FILE; app tự nhận loại, tự đọc (Claude trên máy anh, gói Pro), tự đối chiếu, tự xếp folder; anh duyệt 1 lần để ghi khung.
// HSTT PDF (bản ký) thả ở ô Nạp & duyệt ⇒ gắn vào HSTT Excel cùng HĐ + đợt.
let DM = null, NEN_DS = [], NEN_POLL = null, NEN_SEL = null, NEN_SIG = "";
const KCO = c => c.muc + "|" + c.mo_ta;                                  // khoá 1 cờ — trùng engine nen.khoa_co
const chanCon = r => (r.co || []).filter(c => c.muc === "CHAN" && (c.pl || !(r.co_ok || {})[KCO(c)]));   // CHẶN chưa xử lý (CHẶN số học phụ lục chỉ hết bằng dữ liệu)
const sigNen = r => JSON.stringify([r.trang_thai, r.co, r.co_ok, r.co_hoi]);
const CAN_DT = ["HD_DOI_TAC", "BAO_GIA", "QUYET_TOAN"], LOAI_DT_TEN = {DTC: "Đội thi công", NTP: "Thầu phụ", NCC: "Nhà cung cấp", DVK: "Dịch vụ khác", CDT: "Chủ đầu tư"};
const TT_NEN = {CHO_AI: ["🤖 app đang đọc…", ""], CHO_DUYET: ["chờ anh duyệt", "wa"], DA_LUU: ["đã lưu chứng từ", "ok"], SAI_NOI: ["⛔ nạp sai chỗ — HSTT, nạp ở ① Nạp & duyệt", "er"], DA_NHAP: ["đã nhập khung", "ok"],
                LOI_AI: ["lỗi đọc — bấm đọc lại", "er"], CHO_PHAN_LOAI: ["chưa rõ loại", "wa"]};
async function taiDM() { DM = await api("/danh-muc?du_an=" + encodeURIComponent($("#da").value)); return DM }
async function tabNen() {
  const el = $("#t-nen");
  el.innerHTML = `<div class="sk" style="height:300px"></div>`;
  try { await taiDM() } catch (e) { el.innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>`; return }
  el.innerHTML = `
  <div class="card"><div class="hd"><div><h3>Nạp hồ sơ nền</h3><div class="note">Anh chỉ việc thả file (Excel / PDF, nhiều file một lần) — app <b>tự nhận loại, tự đọc số liệu, tự xếp đúng folder</b>, rồi đưa kết quả để anh duyệt.</div></div></div>
    <div class="pad nen-form">
      <label style="flex:0 1 320px">Loại hồ sơ<select id="n-loai"><option value="">🤖 App tự nhận loại (khuyên dùng)</option>${DM.loai.map(l => `<option value="${l.ma}">${esc(l.ten)}</option>`).join("")}</select></label>
      <label class="n-dt">Đối tác<select id="n-dt"><option value="">🤖 app tự nhận đối tác</option>${DM.doi_tac.filter(d => d.loai !== "CĐT").map(d =>
        `<option value="${esc(d.ma)}">${esc(d.ten)} (${esc(d.ma)})</option>`).join("")}<option value="__moi">＋ Đối tác mới…</option></select></label>
      <label class="n-moi">Mã đối tác mới<input id="n-ma" placeholder="VD NVAn (không dấu, viết liền)"></label>
      <label class="n-moi">Loại đối tác<select id="n-ldt">${Object.entries(LOAI_DT_TEN).filter(([k]) => k !== "CDT").map(([k, v]) => `<option value="${k}">${v}</option>`).join("")}</select></label>
      <label class="n-goi">Gói thầu<select id="n-goi"><option value="">— chưa gán gói —</option>${DM.goi.map(g => `<option value="${esc(g.ma)}">${esc(g.ma)} · ${esc(g.ten)}</option>`).join("")}</select></label>
      <label style="flex:1 1 260px">Ghi chú (tuỳ chọn)<input id="n-gc" placeholder="VD: PL02 điều chỉnh đơn giá trát ngoài"></label>
    </div>
    <div class="drop" id="n-drop" style="margin:0 18px 16px"><b>Kéo thả hồ sơ nền vào đây</b> — hợp đồng, phụ lục, BoQ, ngân sách, gói thầu, báo giá, quyết toán · .pdf .xlsx<br>
      <span class="note">Claude trên máy anh đọc từng file (~1–2 phút/HĐ, tính vào hạn mức gói). File lưu chỉ đọc; trùng nội dung thì không lưu lần 2.</span>
      <input type="file" id="n-f" multiple hidden accept=".pdf,.xlsx,.xlsm,.xls,.doc,.docx,.jpg,.jpeg,.png"></div>
    <div id="n-kq" style="padding:0 18px"></div></div>
  <div class="card"><div class="hd"><div><h3>Sổ hồ sơ nền</h3><div class="note">Bấm một dòng để xem app đọc được gì, các cờ đối chiếu, và duyệt ghi vào khung</div></div></div>
    <div class="right"><table class="bang" id="n-tb"></table></div></div><div id="n-ct"></div>
  <p class="note src">Số liệu do AI đọc từ bản scan — app tự đối chiếu (Σ bảng đơn giá ↔ giá trị HĐ, trường thiếu, trường AI không chắc) nhưng <b>anh vẫn duyệt trước khi ghi vào khung</b>.</p>`;
  const hien = () => { const l = $("#n-loai").value, moi = $("#n-dt").value === "__moi";
    document.querySelectorAll(".n-dt").forEach(x => x.classList.toggle("hide", !CAN_DT.includes(l)));
    document.querySelectorAll(".n-moi").forEach(x => x.classList.toggle("hide", !(CAN_DT.includes(l) && moi)));
    document.querySelectorAll(".n-goi").forEach(x => x.classList.toggle("hide", l !== "CHON_THAU")) };
  $("#n-loai").onchange = hien; $("#n-dt").onchange = hien; hien();
  const dz = $("#n-drop"); dz.onclick = () => $("#n-f").click();
  $("#n-f").onchange = e => { napNen([...e.target.files]); e.target.value = "" };
  dz.ondragover = e => { e.preventDefault(); dz.classList.add("hot") }; dz.ondragleave = () => dz.classList.remove("hot");
  dz.ondrop = async e => { e.preventDefault(); dz.classList.remove("hot"); napNen((await layFile(e.dataTransfer, /\.(pdf|xlsx|xlsm|xls|doc|docx|jpe?g|png)$/i)).files) };   // nhận cả THƯ MỤC
  soNen();
}
async function napNen(files) {
  const l = $("#n-loai").value, moi = $("#n-dt").value === "__moi", dtv = $("#n-dt").value;
  const tham = {loai: l || null, ma_dt: CAN_DT.includes(l) ? (moi ? $("#n-ma").value.trim() : dtv) || null : null,
                loai_dt: moi ? $("#n-ldt").value : null, goi: l === "CHON_THAU" ? $("#n-goi").value : null, ghi_chu: $("#n-gc").value.trim()};
  for (const f of files) {
    const dong = document.createElement("div"); dong.className = "n-dong"; dong.innerHTML = `<span class="chip">đang gửi…</span> ${esc(f.name)}`; $("#n-kq").prepend(dong);
    try { const r = await api("/nap-nen", {du_an: $("#da").value, ten: f.name, b64: await b64(f), ...tham});
      dong.innerHTML = `<span class="chip ${r.trung ? "wa" : "ok"}">${r.trung ? "đã có — không lưu lần 2" : r.trang_thai === "CHO_AI" ? "đã nhận · app đang đọc" : "đã lưu"}</span> ${esc(f.name)}` }
    catch (e) { dong.innerHTML = `<span class="chip er">lỗi</span> ${esc(f.name)} <span class="note">${esc(e.message)}</span>` }
  }
  soNen();
}
async function soNen() {
  NEN_DS = await api("/ho-so-nen?du_an=" + encodeURIComponent($("#da").value));
  const tb = $("#n-tb"); if (!tb) return;
  tb.innerHTML = `<thead><tr><th>Ngày nạp</th><th>Loại (app nhận)</th><th>File</th><th>Đối tác</th><th class="n">Giá trị / tổng</th><th>Cờ</th><th>Trạng thái</th></tr></thead><tbody>` + (NEN_DS.map((r, i) => {
    const a = r.ai || {}, t = TT_NEN[r.trang_thai] || [r.trang_thai, ""], nc = chanCon(r).length, nk = Object.keys(r.co_ok || {}).filter(k => (r.co || []).some(c => KCO(c) === k)).length,
          nl = (r.co || []).filter(c => c.muc !== "CHAN").length - (r.co || []).filter(c => c.muc !== "CHAN" && (r.co_ok || {})[KCO(c)]).length;
    return `<tr class="click" data-i="${i}"><td>${esc(r.luc.replace("T", " ").slice(0, 16))}</td><td><span class="chip">${esc(r.loai_ten)}</span></td>
      <td>${esc(r.ten)}${r.ghi_chu ? `<div class="note">${esc(r.ghi_chu)}</div>` : ""}</td><td>${esc(a.doi_tac_ten || r.ma_doi_tac || "—")}${r.doi_tac_moi ? ' <span class="chip wa">mới</span>' : ""}</td>
      <td class="n">${fmt(a.gia_tri_truoc_vat ?? (r.thong_ke || {}).tong_dong ?? null)}</td>
      <td>${nc ? `<span class="chip er">${nc} chặn</span> ` : ""}${nl ? `<span class="chip wa">${nl} lưu ý</span> ` : ""}${nk ? `<span class="chip ok">✓ ${nk} đã kiểm</span>` : !nc && !nl && r.ai ? '<span class="chip ok">sạch</span>' : ""}</td>
      <td><span class="chip ${t[1]}">${t[0]}</span></td></tr>` }).join("") || `<tr><td colspan="7" class="empty">Chưa có hồ sơ nền nào</td></tr>`) + "</tbody>";
  tb.querySelectorAll("tr.click").forEach(tr => tr.onclick = () => { tb.querySelectorAll("tr").forEach(x => x.classList.toggle("sel", x === tr)); theNen(NEN_DS[+tr.dataset.i]) });
  const sel = NEN_DS.find(r => r.id === NEN_SEL);                       // thẻ đang mở: dữ liệu đổi (Agent trả lời, soát lại) ⇒ vẽ lại tại chỗ
  if (sel && $("#n-ct").innerHTML && sigNen(sel) !== NEN_SIG) theNen(sel, false);
  const dang = NEN_DS.some(r => r.trang_thai === "CHO_AI" || Object.values(r.co_hoi || {}).some(h => h.trang_thai === "DANG"));
  clearTimeout(NEN_POLL); if (dang && !$("#t-nen").classList.contains("hide")) NEN_POLL = setTimeout(soNen, 6000);
}
function bangCo(r) {
  const OK = r.co_ok || {}, HOI = r.co_hoi || {}, co = r.co || [];
  const hoi = h => h ? `<div class="co-hoi ${h.trang_thai}">🤖 ${h.trang_thai === "DANG" ? "Agent đang kiểm lại đúng cờ này…" : esc(h.tra_loi || "")}${h.giay ? ` <span class="note">(${h.giay}s)</span>` : ""}
      ${(h.sua || []).length ? `<div class="note">✏️ Đã sửa: ${esc(h.sua.join(" · "))} ⇒ app đã soát lại toàn bộ cờ</div>` : ""}${h.ghi_chu ? `<div class="note">Anh hỏi: ${esc(h.ghi_chu)}</div>` : ""}</div>` : "";
  const dong = co.map((c, j) => { const k = KCO(c), ok = OK[k], h = HOI[k], dang = h && h.trang_thai === "DANG";
    return `<tr data-j="${j}" class="${ok ? "co-daok" : ""}"><td><span class="b ${c.muc}">${c.muc === "CHAN" ? "CHẶN" : "LƯU Ý"}</span></td>
      <td>${esc(c.mo_ta)}${ok ? `<div class="note co-okt">✓ Anh đã kiểm ${esc(ok.luc.slice(0, 16).replace("T", " "))}${ok.ly_do ? " — " + esc(ok.ly_do) : ""}</div>` : ""}${hoi(h)}<div class="co-f"></div></td>
      <td class="co-act">${ok ? '<button class="btn2 sm co-bo" title="Bỏ xác nhận">↺</button>' : c.muc === "CHAN" && c.pl ? '<span class="note" title="Hết khi số liệu khớp">số liệu</span>' : '<button class="btn2 sm co-ok">✓ OK</button>'}
        <button class="btn2 sm co-hoi-b" ${dang ? "disabled" : ""}>💬 Hỏi</button></td></tr>` }).join("");
  const cu = Object.entries(HOI).filter(([k]) => !co.some(c => KCO(c) === k));   // cờ đã HẾT sau khi Agent sửa
  return (co.length ? `<table class="co-bang">${dong}</table>` : '<div class="note">Không có cờ nào — số liệu tự đối chiếu khớp.</div>')
    + (cu.length ? `<details style="margin-top:8px"><summary class="note">Cờ đã xử lý qua Agent (${cu.length})</summary>${cu.map(([k, h]) => `<div class="note" style="margin:6px 0"><s>${esc(k.split("|").slice(1).join("|"))}</s></div>${hoi(h)}`).join("")}</details>` : "");
}
function ganCo(r) {
  const da = $("#da").value, co = r.co || [];
  $("#n-ct").querySelectorAll(".co-bang tr").forEach(tr => {
    const c = co[+tr.dataset.j], k = KCO(c), f = tr.querySelector(".co-f"), bao = (m, er) => { f.innerHTML = `<div class="res ${er ? "er" : "ok"}">${esc(m)}</div>` };
    const gui = async (url, body) => { try { await api(url, Object.assign({du_an: da, id: r.id, khoa: k}, body)); await soNen() } catch (e) { bao(e.message, true) } };
    const bOk = tr.querySelector(".co-ok"); if (bOk) bOk.onclick = () => {
      if (c.muc !== "CHAN") return gui("/nen/xac-nhan-co", {});
      f.innerHTML = `<div class="nen-form" style="margin-top:6px"><input class="co-ly" placeholder="Lý do chấp nhận cờ CHẶN (bắt buộc)" style="flex:1 1 320px"><button class="btn ok sm co-gui">Xác nhận</button></div>`;
      const ip = f.querySelector(".co-ly"); ip.focus(); f.querySelector(".co-gui").onclick = () => ip.value.trim() ? gui("/nen/xac-nhan-co", {ly_do: ip.value}) : bao("Cần ghi lý do", true) };
    const bBo = tr.querySelector(".co-bo"); if (bBo) bBo.onclick = () => gui("/nen/xac-nhan-co", {bo: true});
    tr.querySelector(".co-hoi-b").onclick = () => {
      f.innerHTML = `<div class="nen-form" style="margin-top:6px"><textarea class="co-gc" rows="2" style="flex:1 1 360px" placeholder="Anh ghi điều cần Agent kiểm (vd: dòng 62 trùng dòng 28, bỏ qua; kiểm lại trang 5…) — để trống: Agent tự kiểm cờ đúng/sai"></textarea>
        <button class="btn sm co-gui">🤖 Gửi Agent</button></div><div class="note">Agent chỉ đọc lại phần liên quan tới cờ này (~30–90s), có sai thì sửa số rồi app tự soát lại.</div>`;
      const ta = f.querySelector(".co-gc"); ta.focus(); f.querySelector(".co-gui").onclick = () => gui("/nen/hoi-co", {ghi_chu: ta.value}) };
  });
}
function theNen(r, cuon = true) {
  NEN_SEL = r.id; NEN_SIG = sigNen(r);
  const a = r.ai || {}, kc = (a.khong_chac || []).join(" ").toLowerCase(), ng = a.nguon || {};
  const o = (k, t, v) => `<div class="kv"><span>${t}${kc.includes(k) ? ' <span class="chip wa">không chắc</span>' : ""}</span><b>${v ?? "—"}${ng[k] ? `<div class="note">${esc(ng[k])}</div>` : ""}</b></div>`;
  const P = v => typeof v === "number" ? (v * 100).toLocaleString("vi-VN", {maximumFractionDigits: 2}) + "%" : null;
  const la = (a.bang || []).filter(x => x.dvt || x.kl != null), tong = la.reduce((s, x) => s + (x.thanh_tien ?? ((x.kl || 0) * (x.don_gia || 0))), 0);
  const nc = chanCon(r).length, hd = ["HD_CDT", "HD_DOI_TAC"].includes(r.loai);
  $("#n-ct").innerHTML = `<div class="card" style="margin-top:14px"><div class="hd"><div><h3>${esc(r.ten)}</h3><div class="note">${esc(r.duong_dan)}</div></div><span class="sp"></span>
    <span class="chip">${esc(r.loai_ten)}</span><span class="chip ${(TT_NEN[r.trang_thai] || ["", ""])[1]}">${(TT_NEN[r.trang_thai] || [r.trang_thai])[0]}</span>
    ${r.ai_meta ? `<span class="note">AI đọc ${r.ai_meta.giay}s · ${a.so_trang ?? "?"} trang</span>` : ""}</div>
    ${r.trang_thai === "LOI_AI" ? `<div class="pad"><div class="warn"><b>Chưa đọc được:</b> ${esc(r.loi || "")}</div></div>` : ""}
    ${r.ai ? `<div class="body"><div class="left">
      ${a.ly_do_loai ? `<div class="note" style="margin-bottom:6px">🤖 ${esc(a.ly_do_loai)}</div>` : ""}
      ${o("so_hd", "Số hợp đồng", esc(a.so_hd))}${o("ngay_ky", "Ngày ký", dd(a.ngay_ky))}${o("ben_giao", "Bên giao", esc(a.ben_giao))}${o("ben_nhan", "Bên nhận", esc(a.ben_nhan))}
      ${o("doi_tac_ten", "Đối tác", esc(a.doi_tac_ten) + (r.ma_doi_tac ? ` <span class="chip ${r.doi_tac_moi ? "wa" : "ok"}">${esc(r.ma_doi_tac)}${r.doi_tac_moi ? " · mới" : ""}</span>` : ""))}
      ${o("loai_doi_tac", "Loại đối tác", esc(LOAI_DT_TEN[a.loai_doi_tac] || a.loai_doi_tac))}${o("dang_hd", "Dạng HĐ", esc(a.dang_hd))}<hr>
      ${o("gia_tri_truoc_vat", "Giá trị trước VAT", fmt(a.gia_tri_truoc_vat))}${o("vat_pct", "VAT", P(a.vat_pct))}${o("gia_tri_sau_vat", "Giá trị sau VAT", fmt(a.gia_tri_sau_vat))}
      ${o("pct_tam_ung", "% tạm ứng", P(a.pct_tam_ung))}${o("pct_tt_dot", "% TT mỗi đợt", P(a.pct_tt_dot))}${o("pct_tt_quyet_toan", "% TT quyết toán", P(a.pct_tt_quyet_toan))}
      ${o("pct_giu_lai", "% giữ lại", P(a.pct_giu_lai))}${o("han_tt_ngay", "Hạn TT", a.han_tt_ngay != null ? a.han_tt_ngay + " ngày " + (a.don_vi_han === "LV" ? "làm việc" : "lịch") : null)}
      ${o("han_qt_ngay", "Hạn quyết toán", a.han_qt_ngay != null ? a.han_qt_ngay + " ngày" : null)}${o("han_tra_gl_ngay", "Hạn trả giữ lại", a.han_tra_gl_ngay != null ? a.han_tra_gl_ngay + " ngày" : null)}
      ${o("bao_hanh_thang", "Bảo hành", a.bao_hanh_thang != null ? a.bao_hanh_thang + " tháng" : null)}${o("co_chu_ky", "Chữ ký / đóng dấu", (a.co_chu_ky ? "✓ ký" : "✗ chưa thấy ký") + " · " + (a.co_dong_dau ? "✓ dấu" : "✗ chưa thấy dấu"))}</div>
     <div class="right"><h4>Cờ đối chiếu (${(r.co || []).length})${nc ? ` · <span style="color:var(--er)">còn ${nc} chặn</span>` : ""}</h4>
      <div class="note" style="margin-bottom:6px">Từng cờ: <b>✓ OK</b> khi anh đã kiểm (CHẶN cần lý do) · <b>💬 Hỏi</b> để Agent kiểm lại đúng cờ đó — xong hết thì bấm Duyệt ở dưới.</div>${bangCo(r)}
      <h4 style="margin-top:14px">Bảng khối lượng / đơn giá (${la.length} dòng · Σ ${fmt(tong)})</h4><table><tr><th>STT</th><th>Nội dung</th><th>ĐVT</th><th class="n">KL</th><th class="n">Đơn giá</th><th class="n">Thành tiền</th></tr>
      ${(a.bang || []).map(x => { const nhom = !(x.dvt || x.kl != null); return `<tr style="${nhom ? "color:var(--mu);font-weight:600" : ""}"><td>${esc(x.stt)}</td><td>${esc(x.noi_dung)}</td><td>${esc(x.dvt)}</td>
        <td class="n">${fmt(x.kl)}</td><td class="n">${fmt(x.don_gia)}</td><td class="n">${fmt(x.thanh_tien)}</td></tr>` }).join("")}</table>
      ${a.ghi_chu ? `<h4 style="margin-top:14px">Ghi chú AI</h4><div class="note">${esc(a.ghi_chu)}</div>` : ""}</div></div>` : r.trang_thai === "CHO_AI" ? '<div class="empty">🤖 App đang đọc file này…</div>' : ""}
    <div class="acts"><button class="btn2" id="n-doclai">🤖 Đọc lại</button>
      ${r.can_nhap ? `<button class="btn ok" id="n-duyet" ${nc || r.trang_thai !== "CHO_DUYET" ? "disabled" : ""}>✓ Duyệt — ghi vào khung</button>` : ""}
      ${nc ? `<span class="note">Còn ${nc} cờ CHẶN chưa xử lý — ✓ OK kèm lý do hoặc 💬 Hỏi Agent từng cờ ở trên.</span>` : ""}<span class="res" id="n-res"></span></div></div>`;
  $("#n-doclai").onclick = async () => { await api("/doc-lai", {du_an: $("#da").value, id: r.id}); $("#n-res").textContent = "Đã xếp lại vào hàng đợi đọc"; soNen() };
  const bt = $("#n-duyet"); if (bt) bt.onclick = async () => {
    bt.disabled = true; bt.textContent = "Đang ghi vào khung…";
    try { const k = await api("/nhap-khung", {du_an: $("#da").value, id: r.id}); $("#n-res").className = "res ok"; $("#n-res").textContent = k.thong_bao || "Đã ghi vào khung"; soNen() }
    catch (e) { $("#n-res").className = "res er"; $("#n-res").textContent = e.message; bt.disabled = false; bt.textContent = "✓ Duyệt — ghi vào khung" } };
  ganCo(r);
  if (cuon) $("#n-ct").scrollIntoView({behavior: "smooth", block: "start"});
}
// ── HSTT PDF (bản ký) thả vào ô Nạp & duyệt: đoán HSTT Excel cùng đơn vị + đợt, anh xác nhận rồi gắn ──
async function thePdfHstt(f) {
  const el = document.createElement("div"); el.className = "card";
  el.innerHTML = `<div class="hd"><h3>${esc(f.name)}</h3><span class="chip">PDF — bản ký HSTT</span></div><div class="pad note">Đang tìm HSTT Excel khớp…</div>`;
  try {
    const ds = await api("/doan-hstt", {du_an: $("#da").value, ten: f.name});
    if (!ds.length) { el.querySelector(".pad").innerHTML = `Chưa có HSTT Excel nào trong dự án để gắn. <b>Anh nạp file Excel của đợt này trước</b>, rồi thả lại PDF — số liệu soát từ Excel, PDF là bản ký đi kèm.`; khoiAiScan(el, f); return el }
    const tot = ds[0].diem > 0 ? ds[0].id : "";
    el.querySelector(".pad").outerHTML = `<div class="pad"><div class="note" style="margin-bottom:8px">PDF là <b>bản scan có ký</b> — không đọc được số. App gắn nó vào HSTT Excel cùng HĐ + đợt;
      khi HSTT đó được ghi sổ, PDF tự vào cùng folder đợt.</div><div class="nen-form">
      <label style="flex:1 1 520px">Gắn vào HSTT Excel<select class="pdf-hs"><option value="">— chọn HSTT —</option>${ds.map(r =>
        `<option value="${r.id}" ${r.id === tot ? "selected" : ""}>${esc(r.ten)} · ${esc(r.ma_hd || "?")} · đợt ${esc(r.dot)} · ${esc((TT[r.trang_thai] || [r.trang_thai])[0])}${r.so_pdf ? ` · đã có ${r.so_pdf} PDF` : ""}</option>`).join("")}</select></label>
      <button class="btn ok pdf-gan">📎 Gắn làm bản ký</button></div>${tot ? '<div class="note">App đã chọn sẵn HSTT khớp nhất theo tên đơn vị + số đợt — anh kiểm lại trước khi bấm.</div>'
      : '<div class="note">Không đoán được HSTT khớp — anh chọn tay.</div>'}<div class="res pdf-kq" style="margin-top:6px"></div></div>`;
    el.querySelector(".pdf-gan").onclick = async () => {
      const id = el.querySelector(".pdf-hs").value, kq = el.querySelector(".pdf-kq");
      if (!id) { kq.className = "res er pdf-kq"; kq.textContent = "Chọn HSTT trước"; return }
      try { const r = await api("/dinh-kem", {du_an: $("#da").value, id, ten: f.name, b64: await b64(f)});
            const x = r.dinh_kem.at(-1); kq.className = "res ok pdf-kq";
            kq.textContent = `Đã gắn vào ${r.ten}` + (x.da_xep ? " · đã xếp vào folder đợt" : " · sẽ tự xếp vào folder đợt khi HSTT được ghi sổ");
            el.querySelector(".pdf-gan").disabled = true }
      catch (e) { kq.className = "res er pdf-kq"; kq.textContent = "Lỗi: " + e.message }
    };
  } catch (e) { el.querySelector(".pad").textContent = "Lỗi: " + e.message }
  khoiAiScan(el, f);
  return el;
}
// ── HSTT chỉ có bản scan: AI đọc số (engine chạy nền 1–3 phút) rồi soát như HSTT Excel ──
function khoiAiScan(el, f) {
  const d = document.createElement("div"); d.className = "acts"; d.style.background = "#f6f2ff";
  d.innerHTML = `<b style="color:#5b3fa0">Chỉ có bản scan?</b><span class="note">Cho <b>AI đọc số trên bản scan</b> rồi soát như HSTT Excel (1–3 phút, dùng gói Claude của anh).
    Số AI không chắc / số học không khép ⇒ bị CHẶN để anh kiểm.</span><button class="btn fix ai-scan">🤖 AI đọc bản scan</button>`;
  el.appendChild(d);
  const b = d.querySelector(".ai-scan");
  b.onclick = async () => {
    b.disabled = true; b.textContent = "AI đang đọc… (1–3 phút)";
    try {
      let r = await api("/doc-scan", {du_an: $("#da").value, ten: f.name, b64: await b64(f)});
      for (let n = 0; r.trang_thai === "DANG_DOC_AI" && n < 120; n++) {
        await new Promise(z => setTimeout(z, 5000)); b.textContent = `AI đang đọc… ${(n + 1) * 5}s`;
        const ds = await api("/ho-so?du_an=" + encodeURIComponent($("#da").value)); r = ds.find(x => x.id === r.id) || r;
      }
      el.replaceWith(the(r));
    } catch (e) { b.disabled = false; b.textContent = "🤖 AI đọc bản scan"; alert("Lỗi: " + e.message) }
  };
}
