const { db, send } = require("./_db");

module.exports = async (req, res) => {
  try {
    const d = await db();
    await d.command({ ping: 1 });
    const counts = {};
    for (const c of ["lessons", "formulas", "questions", "word_skills"]) {
      counts[c] = await d.collection(c).estimatedDocumentCount();
    }
    send(res, 200, { ok: true, counts });
  } catch (e) {
    send(res, e.status || 500, { ok: false, error: e.message });
  }
};
