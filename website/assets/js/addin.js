"use strict";
/* Taskpane chay trong Excel/Word: doc o dang chon va cham bai that. */
const norm = s => (s || "").toUpperCase().replace(/\s+/g, "");
let HOST = "unknown", D = null;

async function fillBai(mon) {
  const lessons = await App.load(mon === "word" ? "word_lessons.json" : "excel_lessons.json");
  const num = l => { const m = String(l.stt != null ? l.stt : (l.bai || "")).match(/(\d+)/); return m ? +m[1] : 999; };
  lessons.sort((a, b) => num(a) - num(b));
  document.getElementById("bai").innerHTML =
    lessons.map(l => `<option value="${App.esc(l.video_id)}">${App.esc(l.bai)}</option>`).join("");
}

async function loadTasks() {
  const mon = document.getElementById("mon").value;
  const vid = document.getElementById("bai").value;
  D = await App.load(`lessons/${mon}_${vid}.json`);
  document.getElementById("tasks").innerHTML =
    `<p class="btn-row"><button id="grade-range" ${HOST === "Excel" ? "" : "disabled"}>Chấm cả vùng đang chọn</button>
    ${HOST === "Word" ? `<button id="grade-word">Chấm đoạn văn đang chọn</button>` : ""}</p>` +
    (D.practice || []).map((t, i) => `
    <div class="card" style="margin-bottom:0.75rem" data-i="${i}">
      <p><strong>${i + 1}.</strong> ${App.esc(t.yeu_cau)}</p>
      ${t.goi_y ? `<p class="muted">Gợi ý: ${App.esc(t.goi_y)}</p>` : ""}
      <p class="btn-row"><button data-check ${HOST === "Excel" ? "" : "disabled"}>Kiểm tra ô đang chọn</button></p>
      <p class="muted" data-msg></p>
    </div>`).join("") || `<div class="empty">Bài này chưa có bài tập.</div>`;
  document.querySelectorAll("[data-check]").forEach(b => b.addEventListener("click", () => checkTask(+b.closest(".card").dataset.i, b)));
  const gr = document.getElementById("grade-range");
  if (gr) gr.addEventListener("click", gradeRange);
  const gw = document.getElementById("grade-word");
  if (gw) gw.addEventListener("click", gradeWord);
}

async function gradeRange() {
  const box = document.getElementById("tasks");
  let msg = document.getElementById("range-msg");
  if (!msg) {
    msg = document.createElement("p");
    msg.id = "range-msg";
    msg.className = "muted";
    box.prepend(msg);
  }
  msg.textContent = "Đang đọc vùng trong Excel…";
  try {
    const cells = await Excel.run(async (ctx) => {
      const r = ctx.workbook.getSelectedRange();
      r.load(["address", "formulas", "values", "rowCount", "columnCount"]);
      await ctx.sync();
      return { address: r.address, formulas: r.formulas, values: r.values };
    });
    const hay = [];
    cells.formulas.forEach(row => row.forEach(v => { if (v) hay.push(norm(String(v))); }));
    cells.values.forEach(row => row.forEach(v => { if (v !== "" && v != null) hay.push(norm(String(v))); }));
    let done = 0, total = 0;
    (D.practice || []).forEach(t => {
      if (!t.dap_an && !t.tu_khoa) return;
      total++;
      let ok = false;
      if (t.dap_an) ok = t.dap_an.some(a => hay.includes(norm(a)));
      if (!ok && t.tu_khoa) {
        const all = hay.join(" ");
        const hit = t.tu_khoa.filter(k => all.includes(norm(k)));
        ok = hit.length >= Math.max(2, Math.ceil(t.tu_khoa.length * 0.6));
      }
      if (ok) done++;
    });
    msg.innerHTML = `Vùng ${App.esc(cells.address)}: khớp <strong>${done}/${total}</strong> bài tập. Bôi đen vùng khác rồi chấm lại để kiểm tra tiếp.`;
  } catch (e) {
    msg.textContent = "Không đọc được vùng. Bôi đen 1 vùng ô trong Excel rồi thử lại.";
  }
}

async function gradeWord() {
  let msg = document.getElementById("range-msg");
  if (!msg) {
    msg = document.createElement("p");
    msg.id = "range-msg";
    msg.className = "muted";
    document.getElementById("tasks").prepend(msg);
  }
  msg.textContent = "Đang đọc đoạn văn trong Word…";
  try {
    const text = await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      await ctx.sync();
      return sel.text || "";
    });
    if (!text.trim()) { msg.textContent = "Bạn chưa bôi đen văn bản nào. Bôi đen 1 đoạn trong Word rồi bấm lại."; return; }
    const hay = norm(text);
    let done = 0, total = 0;
    const detail = [];
    (D.practice || []).forEach((t, i) => {
      if (!t.tu_khoa && !t.dap_an) return;
      total++;
      let ok = false;
      if (t.tu_khoa) {
        const hit = t.tu_khoa.filter(k => hay.includes(norm(k)));
        ok = hit.length >= Math.max(2, Math.ceil(t.tu_khoa.length * 0.6));
      } else if (t.dap_an) {
        ok = t.dap_an.some(a => hay.includes(norm(a).slice(0, 12)));
      }
      if (ok) { done++; detail.push(i + 1); }
    });
    msg.innerHTML = `Đoạn đã chọn khớp <strong>${done}/${total}</strong> bài tập${detail.length ? ` (bài ${detail.join(", ")})` : ""}.`;
  } catch (e) {
    msg.textContent = "Không đọc được vùng chọn. Bôi đen văn bản trong Word rồi thử lại.";
  }
}

async function readActiveCell() {
  return Excel.run(async (ctx) => {
    const cell = ctx.workbook.getActiveCell();
    cell.load(["address", "formulas", "values"]);
    await ctx.sync();
    return { address: cell.address, formula: cell.formulas[0][0], value: cell.values[0][0] };
  });
}

async function checkTask(i, btn) {
  const t = D.practice[i];
  const box = btn.closest(".card");
  const msg = box.querySelector("[data-msg]");
  msg.textContent = "Đang đọc ô trong Excel…";
  try {
    const cell = await readActiveCell();
    let ok = false, detail = "";
    if (t.loai === "cong_thuc" && t.dap_an) {
      const got = norm(String(cell.formula === "" ? cell.value : cell.formula));
      ok = t.dap_an.some(a => norm(a) === got);
      detail = `Ô ${cell.address}: <code>${App.esc(String(cell.formula === "" ? cell.value : cell.formula))}</code>`;
    } else if (t.tu_khoa) {
      const hay = norm(String(cell.value)) + norm(String(cell.formula));
      const hit = t.tu_khoa.filter(k => hay.includes(norm(k)));
      ok = hit.length >= Math.max(2, Math.ceil(t.tu_khoa.length * 0.6));
      detail = `Ô ${cell.address} trúng ${hit.length}/${t.tu_khoa.length} từ khóa`;
    }
    msg.innerHTML = (ok ? "<strong>Đúng ✓</strong> " : "<strong>Chưa đúng.</strong> ") + detail +
      (ok || !t.dap_an ? "" : `<br>Đáp án mẫu: <code>${App.esc(t.dap_an.join(" | "))}</code>`);
  } catch (e) {
    msg.textContent = "Không đọc được ô. Hãy bôi đen 1 ô trong Excel rồi thử lại.";
  }
}

Office.onReady(async (info) => {
  HOST = (info && info.host) || "unknown";
  const st = document.getElementById("status");
  if (HOST === "Excel") {
    st.innerHTML = 'Đã kết nối <span class="badge">Excel ✓</span> Bôi đen ô cần chấm rồi bấm Kiểm tra.';
  } else if (HOST === "Word") {
    st.innerHTML = 'Đang chạy trong <span class="badge">Word</span>. Chế độ Word: xem yêu cầu và tự đối chiếu (chấm ô tự động chỉ hỗ trợ Excel).';
  } else {
    st.innerHTML = 'Trang này mở ngoài Office — hãy <a href="thuc-hanh.html">luyện trên web</a> hoặc cài add-in vào Excel.';
  }
  await fillBai(document.getElementById("mon").value);
  document.getElementById("mon").addEventListener("change", async () => { await fillBai(document.getElementById("mon").value); loadTasks(); });
  document.getElementById("bai").addEventListener("change", loadTasks);
  loadTasks();
});
