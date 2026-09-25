// HỒ SƠ NỀN (BoQ, ngân sách, gói thầu, báo giá, HĐ, quyết toán — Excel/PDF) + HSTT PDF bản ký gắn vào HSTT Excel.
// PDF của dự án là bản SCAN ⇒ app LƯU làm chứng từ đúng folder; số liệu HĐ nhập qua form (không tự đọc số từ ảnh).
let DM = null;
const CAN_DT = ["HD_DOI_TAC", "BAO_GIA", "QUYET_TOAN"], LOAI_DT_TEN = {DTC: "Đội thi công", NTP: "Thầu phụ", NCC: "Nhà cung cấp", DVK: "Dịch vụ khác"};
async function taiDM() { DM = await api("/danh-muc?du_an=" + encodeURIComponent($("#da").value)); return DM }
async function tabNen() {
  const el = $("#t-nen");
  el.innerHTML = `<div class="sk" style="height:300px"></div>`;
  try { await taiDM() } catch (e) { el.innerHTML = `<div class="card"><div class="empty">${esc(e.message)}</div></div>`; return }
  el.innerHTML = `
  <div class="card"><div class="hd"><div><h3>Nạp hồ sơ nền</h3><div class="note">Excel hoặc PDF — app xếp vào đúng folder của dự án trên Drive. Nhiều file một lần (cùng loại, cùng đối tác).</div></div></div>
    <div class="pad nen-form">
      <label>Loại hồ sơ<select id="n-loai">${DM.loai.map(l => `<option value="${l.ma}">${esc(l.ten)}</option>`).join("")}</select></label>
      <label class="n-dt">Đối tác<select id="n-dt"><option value="">— chọn đối tác —</option>${DM.doi_tac.filter(d => d.loai !== "CĐT").map(d =>
        `<option value="${esc(d.ma)}">${esc(d.ten)} (${esc(d.ma)})</option>`).join("")}<option value="__moi">＋ Đối tác mới (chưa có trong khung)…</option></select></label>
      <label class="n-moi">Mã đối tác mới<input id="n-ma" placeholder="VD NVAn (không dấu, viết liền)"></label>
      <label class="n-moi">Loại đối tác<select id="n-ldt">${Object.entries(LOAI_DT_TEN).map(([k, v]) => `<option value="${k}">${v}</option>`).join("")}</select></label>
      <label class="n-goi">Gói thầu<select id="n-goi"><option value="">— chưa gán gói —</option>${DM.goi.map(g => `<option value="${esc(g.ma)}">${esc(g.ma)} · ${esc(g.ten)}</option>`).join("")}</select></label>
      <label style="flex:1 1 260px">Ghi chú<input id="n-gc" placeholder="VD: PL02 điều chỉnh đơn giá trát ngoài"></label>
    </div>
    <div class="drop" id="n-drop" style="margin:0 18px 16px"><b>Kéo thả file hồ sơ nền vào đây</b> — .pdf .xlsx .xls .docx .jpg<br>
      <span class="note">File lưu chỉ đọc vào folder dự án; trùng nội dung thì không lưu lần 2.</span><input type="file" id="n-f" multiple hidden
      accept=".pdf,.xlsx,.xlsm,.xls,.doc,.docx,.jpg,.jpeg,.png"></div>
    <div id="n-kq" style="padding:0 18px"></div></div>
  <div class="card"><div class="hd"><div><h3>Sổ hồ sơ nền</h3><div class="note">Mọi hồ sơ nền đã nạp của dự án — bấm "Mở folder" trên Drive để xem file gốc</div></div></div>
    <div class="right"><table class="bang" id="n-tb"></table></div></div>
  <p class="note src">PDF của dự án là bản scan (không có lớp chữ) ⇒ app không tự đọc số tiền. Hợp đồng / BoQ / ngân sách sẽ được nhập vào khung qua form xác nhận
    (cột “Vào khung”) — anh kiểm từng số trước khi ghi.</p>`;
  const hien = () => { const l = $("#n-loai").value, moi = $("#n-dt").value === "__moi";
    document.querySelectorAll(".n-dt").forEach(x => x.classList.toggle("hide", !CAN_DT.includes(l)));
    document.querySelectorAll(".n-moi").forEach(x => x.classList.toggle("hide", !(CAN_DT.includes(l) && moi)));
    document.querySelectorAll(".n-goi").forEach(x => x.classList.toggle("hide", l !== "CHON_THAU")) };
  $("#n-loai").onchange = hien; $("#n-dt").onchange = hien; hien();
  const dz = $("#n-drop"); dz.onclick = () => $("#n-f").click();
  $("#n-f").onchange = e => { napNen([...e.target.files]); e.target.value = "" };
  dz.ondragover = e => { e.preventDefault(); dz.classList.add("hot") }; dz.ondragleave = () => dz.classList.remove("hot");
  dz.ondrop = e => { e.preventDefault(); dz.classList.remove("hot"); napNen([...e.dataTransfer.files]) };
  soNen();
}
async function napNen(files) {
  const l = $("#n-loai").value, moi = $("#n-dt").value === "__moi";
  const tham = {loai: l, ma_dt: CAN_DT.includes(l) ? (moi ? $("#n-ma").value.trim() : $("#n-dt").value) : null,
                loai_dt: moi ? $("#n-ldt").value : null, goi: l === "CHON_THAU" ? $("#n-goi").value : null, ghi_chu: $("#n-gc").value.trim()};
  for (const f of files) {
    const dong = document.createElement("div"); dong.className = "n-dong"; dong.innerHTML = `<span class="chip">đang lưu…</span> ${esc(f.name)}`; $("#n-kq").prepend(dong);
    try { const r = await api("/nap-nen", {du_an: $("#da").value, ten: f.name, b64: await b64(f), ...tham});
      dong.innerHTML = `<span class="chip ${r.trung ? "wa" : "ok"}">${r.trung ? "đã có — không lưu lần 2" : "đã lưu"}</span> ${esc(f.name)} <span class="note">→ ${esc(r.duong_dan)}</span>` }
    catch (e) { dong.innerHTML = `<span class="chip er">lỗi</span> ${esc(f.name)} <span class="note">${esc(e.message)}</span>` }
  }
  soNen();
}
async function soNen() {
  const ds = await api("/ho-so-nen?du_an=" + encodeURIComponent($("#da").value));
  $("#n-tb").innerHTML = `<thead><tr><th>Ngày nạp</th><th>Loại</th><th>File</th><th>Đối tác / gói</th><th>Nơi lưu</th><th>Vào khung</th></tr></thead><tbody>` + (ds.map(r =>
    `<tr><td>${esc(r.luc.replace("T", " ").slice(0, 16))}</td><td><span class="chip">${esc(r.loai_ten)}</span></td><td>${esc(r.ten)}${r.ghi_chu ? `<div class="note">${esc(r.ghi_chu)}</div>` : ""}</td>
     <td>${esc(r.ma_doi_tac || r.goi || "—")}</td><td class="note">${esc(r.duong_dan.split("\\").slice(1, -1).join(" › "))}</td>
     <td>${!r.can_nhap ? '<span class="note">chỉ lưu chứng từ</span>' : r.da_nhap_khung ? '<span class="chip ok">đã nhập</span>' : '<span class="chip wa">chờ nhập vào khung</span>'}</td></tr>`).join("")
    || `<tr><td colspan="6" class="empty">Chưa có hồ sơ nền nào</td></tr>`) + "</tbody>";
}
// ── HSTT PDF (bản ký) thả vào ô Nạp & duyệt: đoán HSTT Excel cùng đơn vị + đợt, anh xác nhận rồi gắn ──
async function thePdfHstt(f) {
  const el = document.createElement("div"); el.className = "card";
  el.innerHTML = `<div class="hd"><h3>${esc(f.name)}</h3><span class="chip">PDF — bản ký HSTT</span></div><div class="pad note">Đang tìm HSTT Excel khớp…</div>`;
  try {
    const ds = await api("/doan-hstt", {du_an: $("#da").value, ten: f.name});
    if (!ds.length) { el.querySelector(".pad").innerHTML = `Chưa có HSTT Excel nào trong dự án để gắn. <b>Anh nạp file Excel của đợt này trước</b>, rồi thả lại PDF — số liệu soát từ Excel, PDF là bản ký đi kèm.`; return el }
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
  return el;
}
