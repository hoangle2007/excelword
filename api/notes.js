const { db, send } = require("./_db");

module.exports = async (req, res) => {
  try {
    const d = await db();
    const col = d.collection("notes");
    if (req.method === "POST") {
      let body = req.body;
      if (typeof body === "string") { try { body = JSON.parse(body); } catch { body = {}; } }
      const { owner = "", ref = "", text = "", t = null } = body || {};
      if (!owner || !ref || !text) return send(res, 400, { error: "Thieu owner/ref/text" });
      const doc = { owner: String(owner).slice(0, 64), ref: String(ref).slice(0, 128), text: String(text).slice(0, 2000), t, when: new Date().toISOString() };
      await col.insertOne(doc);
      return send(res, 200, { ok: true });
    }
    const { owner = "", ref = "" } = req.query || {};
    if (!owner || !ref) return send(res, 400, { error: "Thieu owner/ref" });
    const items = await col.find({ owner, ref }).sort({ _id: -1 }).limit(100).toArray();
    send(res, 200, items);
  } catch (e) {
    send(res, e.status || 500, { error: e.message });
  }
};
