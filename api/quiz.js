const { db, send } = require("./_db");

module.exports = async (req, res) => {
  try {
    const { mon, bai = "", dokho = "", n = "20" } = req.query || {};
    if (mon !== "excel" && mon !== "word") return send(res, 400, { error: "mon=excel|word" });
    const count = Math.min(Math.max(parseInt(n, 10) || 20, 1), 40);
    const match = { mon };
    if (bai) match.video_id = bai;
    if (dokho) match.do_kho_norm = String(dokho).toLowerCase().replace(/_/g, " ").trim();
    const d = await db();
    const items = await d.collection("questions")
      .aggregate([{ $match: match }, { $sample: { size: count } }]).toArray();
    send(res, 200, items);
  } catch (e) {
    send(res, e.status || 500, { error: e.message });
  }
};
