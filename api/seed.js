// POST /api/seed  header x-seed-key: <SEED_KEY>  (goi 1 lan sau deploy)
const path = require("path");
const fs = require("fs");
const { db, send } = require("./_db");

const num = (l) => { const m = String(l.stt != null ? l.stt : (l.bai || "")).match(/(\d+)/); return m ? +m[1] : 999; };

module.exports = async (req, res) => {
  try {
    if (req.method !== "POST") return send(res, 405, { error: "POST only" });
    if (!process.env.SEED_KEY || req.headers["x-seed-key"] !== process.env.SEED_KEY) {
      return send(res, 401, { error: "Sai seed key" });
    }
    const base = path.join(process.cwd(), "website", "data");
    const load = (f) => JSON.parse(fs.readFileSync(path.join(base, f), "utf-8"));
    const excelLessons = load("excel_lessons.json").map((l) => ({ ...l, mon: "excel", stt: num(l) }));
    const wordLessons = load("word_lessons.json").map((l) => ({ ...l, mon: "word", stt: num(l) }));
    const formulas = load("excel_formulas.json");
    const normQ = (q, mon) => ({ ...q, mon, do_kho_norm: String(q.do_kho || "").toLowerCase().replace(/_/g, " ").trim() });
    const questions = [
      ...load("excel_quiz.json").map((q) => normQ(q, "excel")),
      ...load("word_quiz.json").map((q) => normQ(q, "word")),
    ];
    const skills = load("word_skills.json");
    const d = await db();
    const out = {};
    for (const [name, docs] of [
      ["lessons", [...excelLessons, ...wordLessons]],
      ["formulas", formulas],
      ["questions", questions],
      ["word_skills", skills],
    ]) {
      await d.collection(name).deleteMany({});
      if (docs.length) await d.collection(name).insertMany(docs);
      out[name] = docs.length;
    }
    await d.collection("questions").createIndex({ mon: 1, video_id: 1 });
    await d.collection("formulas").createIndex({ ten: 1 });
    send(res, 200, { ok: true, ...out });
  } catch (e) {
    send(res, e.status || 500, { error: e.message });
  }
};
