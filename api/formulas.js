const { db, send } = require("./_db");

module.exports = async (req, res) => {
  try {
    const { q = "", nhom = "", limit = "80" } = req.query || {};
    const lim = Math.min(Math.max(parseInt(limit, 10) || 80, 1), 200);
    const filter = {};
    if (nhom) filter.nhom = nhom;
    if (q) {
      const rx = { $regex: q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), $options: "i" };
      filter.$or = [{ ten: rx }, { cu_phap: rx }, { giai_thich: rx }];
    }
    const d = await db();
    const items = await d.collection("formulas").find(filter).limit(lim).toArray();
    send(res, 200, items);
  } catch (e) {
    send(res, e.status || 500, { error: e.message });
  }
};
