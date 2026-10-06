const { MongoClient } = require("mongodb");

let client = null;

async function db() {
  const uri = process.env.MONGODB_URI;
  if (!uri) {
    const err = new Error("Thieu MONGODB_URI trong Environment Variables");
    err.status = 500;
    throw err;
  }
  if (!client) {
    client = new MongoClient(uri, { maxPoolSize: 5, serverSelectionTimeoutMS: 8000 });
    await client.connect();
  }
  return client.db(process.env.MONGODB_DB || "webexcel");
}

function send(res, code, obj) {
  res.statusCode = code;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(obj));
}

module.exports = { db, send };
