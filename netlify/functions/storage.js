const { getStore } = require("@netlify/blobs");

function json(obj, statusCode = 200) {
  return {
    statusCode,
    headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
    body: JSON.stringify(obj),
  };
}

exports.handler = async (event) => {
  const store = getStore("eca-data");
  const params = event.queryStringParameters || {};
  const { op, key } = params;
  const prefix = params.prefix ?? "";

  try {
    if (event.httpMethod === "GET" && op === "get") {
      if (!key) return json({ error: "key required" }, 400);
      const value = await store.get(key);
      return json({ value: value ?? null });
    }

    if (event.httpMethod === "POST" && op === "set") {
      if (!key) return json({ error: "key required" }, 400);
      const body = JSON.parse(event.body || "{}");
      await store.set(key, String(body.value ?? ""));
      return json({ ok: true });
    }

    if (event.httpMethod === "POST" && op === "delete") {
      if (!key) return json({ error: "key required" }, 400);
      await store.delete(key);
      return json({ ok: true });
    }

    if (event.httpMethod === "GET" && op === "list") {
      const { blobs } = await store.list({ prefix });
      return json({ keys: blobs.map((b) => b.key) });
    }

    return json({ error: "Unknown operation" }, 400);
  } catch (err) {
    return json({ error: err.message || "Storage error" }, 500);
  }
};
