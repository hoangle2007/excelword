const { db, send } = require("./_db");

module.exports = async (req, res) => {
  try {
    const { mon } = req.query || {};
    if (mon !== "excel" && mon !== "word") return send(res, 400, { error: "mon=excel|word" });
    const d = await db();
    const items = await d.collection("lessons").find({ mon }).sort({ stt: 1 }).toArray();
    send(res, 200, items);
  } catch (e) {
    send(res, e.status || 500, { error: e.message });
  }
};
