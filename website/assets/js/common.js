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
  }
};
document.addEventListener("DOMContentLoaded", () => {
  const page = document.body.dataset.page;
  document.querySelectorAll(".nav-links a").forEach(a => {
    if (a.dataset.nav === page) a.setAttribute("aria-current", "page");
  });
});
