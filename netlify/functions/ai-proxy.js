const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function jsonResponse(statusCode, body) {
  return {
    statusCode,
    headers: { "Content-Type": "application/json", ...corsHeaders },
    body: JSON.stringify(body),
  };
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return {
      statusCode: 200,
      headers: corsHeaders,
      body: "",
    };
  }

  if (event.httpMethod !== "POST") {
    return jsonResponse(405, { error: "Method Not Allowed" });
  }

  let payload;
  try {
    const raw = event.isBase64Encoded
      ? Buffer.from(event.body || "", "base64").toString("utf-8")
      : event.body || "";
    payload = JSON.parse(raw || "{}");
  } catch (err) {
    return jsonResponse(400, { error: "Invalid JSON body" });
  }

  const apiUrl = (payload.apiUrl || process.env.AI_UPSTREAM_URL || "").trim();
  const apiKey = (payload.apiKey || process.env.AI_API_KEY || "").trim();
  const model = (payload.model || "").trim();
  const prompt = (payload.prompt || "").trim();
  const imageDataUrl = payload.imageDataUrl || "";

  if (!apiUrl) return jsonResponse(400, { error: "Missing apiUrl" });
  if (!apiKey) return jsonResponse(400, { error: "Missing apiKey" });
  if (!model) return jsonResponse(400, { error: "Missing model" });
  if (!prompt) return jsonResponse(400, { error: "Missing prompt" });
  if (!imageDataUrl) return jsonResponse(400, { error: "Missing imageDataUrl" });

  const body = {
    model,
    messages: [
      { role: "system", content: prompt },
      {
        role: "user",
        content: [
          { type: "text", text: "请从图片中提取并输出 JSON" },
          { type: "image_url", image_url: { url: imageDataUrl } },
        ],
      },
    ],
    temperature: 0,
  };

  try {
    const resp = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(body),
    });

    const text = await resp.text();
    return {
      statusCode: resp.status,
      headers: { "Content-Type": "application/json", ...corsHeaders },
      body: text,
    };
  } catch (err) {
    return jsonResponse(502, { error: "Upstream fetch failed", detail: String(err) });
  }
};
