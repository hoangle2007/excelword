"use strict";
const App = {
  cache: {},
  async load(name) {
    if (this.cache[name]) return this.cache[name];
    const res = await fetch("data/" + name);
    if (!res.ok) throw new Error("Không tải được " + name);
    const data = await res.json();
    this.cache[name] = data;
    return data;
  },
  esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  },
  async apiHealth() {
    if (this._api === undefined) {
      try {
        const r = await fetch("api/health");
        const j = await r.json();
        this._api = r.ok && !!j.ok;
      } catch { this._api = false; }
    }
    return this._api;
  },
  async apiGet(path) {
    try {
      const r = await fetch("api/" + path);
      if (!r.ok) return null;
      return await r.json();
    } catch { return null; }
  },
  async apiPost(path, body) {
    try {
      const r = await fetch("api/" + path, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
      });
      return r.ok;
    } catch { return false; }
  },
  debounce(fn, ms) {
    let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
  },
  store: {
    get(k, fb) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : fb; } catch { return fb; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} }
  },
  doneLessons() { return this.store.get("doneLessons", {}); },
  toggleLesson(id) {
    const d = this.doneLessons();
    if (d[id]) delete d[id]; else d[id] = Date.now();
    this.store.set("doneLessons", d);
    return !!d[id];
  },
  saveResult(r) {
    const h = this.store.get("quizHistory", []);
    h.unshift(r);
    this.store.set("quizHistory", h.slice(0, 30));
  },
  lessonPicker(el, opts) {
    const o = Object.assign({ placeholder: "Chọn bài học", searchPh: "Gõ để tìm bài…", options: [], value: "", onChange: null }, opts);
    let value = o.value, active = -1;
    el.innerHTML = `
      <button type="button" class="lp-btn" aria-haspopup="listbox" aria-expanded="false"><span class="lp-cur"></span><span aria-hidden="true">▾</span></button>
      <div class="lp-panel" hidden>
        <input type="search" class="lp-search" placeholder="${this.esc(o.searchPh)}" aria-label="Tìm bài học">
        <ul class="lp-list" role="listbox" aria-label="Danh sách bài học"></ul>
      </div>`;
    el.classList.add("lp");
    const btn = el.querySelector(".lp-btn"), panel = el.querySelector(".lp-panel"),
      search = el.querySelector(".lp-search"), list = el.querySelector(".lp-list");
    const paintBtn = () => {
      const cur = o.options.find(x => x.value === value);
      btn.querySelector(".lp-cur").innerHTML = cur
        ? `<strong>${this.esc(cur.label)}</strong>${cur.desc ? `<br><small class="muted">${this.esc(cur.desc)}</small>` : ""}`
        : `<span class="muted">${this.esc(o.placeholder)}</span>`;
    };
    const paintList = () => {
      const q = search.value.toLowerCase();
      const items = o.options.filter(x => !q || ((x.label || "") + " " + (x.desc || "")).toLowerCase().includes(q));
      list.innerHTML = items.length ? items.map((x, i) => `
        <li class="lp-opt${x.value === value ? " sel" : ""}${i === active ? " act" : ""}" role="option" tabindex="-1"
            aria-selected="${x.value === value}" data-v="${this.esc(x.value)}" data-i="${i}">
          <strong>${this.esc(x.label)}</strong>
          ${x.desc ? `<br><small class="muted">${this.esc(x.desc)}</small>` : ""}
          ${x.meta ? `<br><small><span class="badge">${this.esc(x.meta)}</span></small>` : ""}
        </li>`).join("") : `<li class="muted" style="padding:.5rem">Không tìm thấy bài phù hợp.</li>`;
      list.querySelectorAll(".lp-opt[data-v]").forEach(li => li.addEventListener("click", () => select(li.dataset.v)));
      return items;
    };
    const open = () => { panel.hidden = false; btn.setAttribute("aria-expanded", "true"); search.value = ""; active = -1; paintList(); search.focus(); };
    const close = () => { panel.hidden = true; btn.setAttribute("aria-expanded", "false"); };
    const select = v => { value = v; paintBtn(); close(); btn.focus(); if (o.onChange) o.onChange(v); };
    btn.addEventListener("click", () => panel.hidden ? open() : close());
    search.addEventListener("input", () => { active = -1; paintList(); });
    search.addEventListener("keydown", e => {
      const items = [...list.querySelectorAll(".lp-opt[data-v]")];
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        active = e.key === "ArrowDown" ? Math.min(active + 1, items.length - 1) : Math.max(active - 1, 0);
        paintList();
        const cur = list.querySelectorAll(".lp-opt[data-v]")[active];
        if (cur) cur.scrollIntoView({ block: "nearest" });
      } else if (e.key === "Enter") {
        const cur = list.querySelectorAll(".lp-opt[data-v]")[active >= 0 ? active : 0];
        if (cur) select(cur.dataset.v);
      } else if (e.key === "Escape") close();
    });
    document.addEventListener("click", function out(e) {
      if (!el.contains(e.target)) close();
    });
    paintBtn();
    return { get: () => value, set: v => { value = v; paintBtn(); } };
  }
};
document.addEventListener("DOMContentLoaded", () => {
  const page = document.body.dataset.page;
  document.querySelectorAll(".nav-links a").forEach(a => {
    if (a.dataset.nav === page) a.setAttribute("aria-current", "page");
  });
});
