// Seed du lieu len MongoDB Atlas: MONGODB_URI="mongodb+srv://..." npm run seed
import { readFileSync } from "fs";
import { MongoClient } from "mongodb";

const uri = process.env.MONGODB_URI;
if (!uri) { console.error("Thieu MONGODB_URI"); process.exit(1); }
const DB = process.env.MONGODB_DB || "webexcel";
const load = (p) => JSON.parse(readFileSync(p, "utf-8"));
const num = (l) => { const m = String(l.stt != null ? l.stt : (l.bai || "")).match(/(\d+)/); return m ? +m[1] : 999; };

const excelLessons = load("website/data/excel_lessons.json").map((l) => ({ ...l, mon: "excel", stt: num(l) }));
const wordLessons = load("website/data/word_lessons.json").map((l) => ({ ...l, mon: "word", stt: num(l) }));
const formulas = load("website/data/excel_formulas.json");
const excelQ = load("website/data/excel_quiz.json").map((q) => ({ ...q, mon: "excel", do_kho_norm: String(q.do_kho || "").toLowerCase().replace(/_/g, " ").trim() }));
const wordQ = load("website/data/word_quiz.json").map((q) => ({ ...q, mon: "word", do_kho_norm: String(q.do_kho || "").toLowerCase().replace(/_/g, " ").trim() }));
const skills = load("website/data/word_skills.json");

const client = new MongoClient(uri);
await client.connect();
const db = client.db(DB);
async function refill(name, docs) {
  const col = db.collection(name);
  await col.deleteMany({});
  if (docs.length) await col.insertMany(docs);
  console.log(name, docs.length);
}
await refill("lessons", [...excelLessons, ...wordLessons]);
await refill("formulas", formulas);
await refill("questions", [...excelQ, ...wordQ]);
await refill("word_skills", skills);
await db.collection("questions").createIndex({ mon: 1, video_id: 1 });
await db.collection("formulas").createIndex({ ten: 1 });
await client.close();
console.log("SEED DONE ->", DB);
