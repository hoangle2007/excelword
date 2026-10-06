const { db, send } = require("./_db");

module.exports = async (req, res) => {
  try {
    const { q = "", limit = "60" } = req.query || {};
    const lim = Math.min(Math.max(parseInt(limit, 10) || 60, 1), 200);
    const filter = {};
    if (q) {
      const rx = { $regex: q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), $options: "i" };
      filter.$or = [{ ten: rx }, { thao_tac: rx }, { giai_thich: rx }];
    }
    const d = await db();
    const items = await d.collection("word_skills").find(filter).limit(lim).toArray();
    send(res, 200, items);
  } catch (e) {
    send(res, e.status || 500, { error: e.message });
  }
};
