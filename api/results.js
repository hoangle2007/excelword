const { db, send } = require("./_db");

module.exports = async (req, res) => {
  try {
    const d = await db();
    const col = d.collection("results");
    if (req.method === "POST") {
      let body = req.body;
      if (typeof body === "string") { try { body = JSON.parse(body); } catch { body = {}; } }
      const { mon = "", score = 0, total = 0 } = body || {};
      const doc = { when: new Date().toISOString().slice(0, 16).replace("T", " "), mon, score: +score || 0, total: +total || 0 };
      await col.insertOne(doc);
      return send(res, 200, { ok: true });
    }
    const items = await col.find({}).sort({ _id: -1 }).limit(30).toArray();
    send(res, 200, items);
  } catch (e) {
    send(res, e.status || 500, { error: e.message });
  }
};
