/*
 * Mock vLLM server for developing ExcelLocal without a GPU.
 *   node tools/mock-vllm.js   (listens on http://localhost:8000)
 * OpenAI-compatible endpoints:
 *   GET  /v1/models
 *   POST /v1/chat/completions  (SSE streaming)
 * NOTE: SSE prefixes are built from char codes to dodge a known
 * file-corruption hazard with the literal token in this repo.
 */
const http = require("http");

const PORT = Number(process.env.MOCK_PORT || 8000);
// char codes: d,a,t,a,:,space and [,D,O,N,E,]
const PREFIX = String.fromCharCode(100, 97, 116, 97, 58, 32);
const DONE_TOKEN = PREFIX + String.fromCharCode(91, 68, 79, 78, 69, 93);

function sseLine(res, obj) {
  res.write(PREFIX + JSON.stringify(obj) + "\n\n");
}

function startSse(res) {
  res.writeHead(200, {
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
  });
}

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
};

function handleChat(res) {
  startSse(res);
  const tokens = ["Hello", " from", " mock", " vLLM", "!"];
  let i = 0;
  const timer = setInterval(() => {
    if (i < tokens.length) {
      sseLine(res, {
        id: "mock",
        object: "chat.completion.chunk",
        choices: [{ index: 0, delta: { content: tokens[i++] }, finish_reason: null }],
      });
    } else {
      sseLine(res, {
        id: "mock",
        object: "chat.completion.chunk",
        choices: [{ index: 0, delta: {}, finish_reason: "stop" }],
      });
      res.write(DONE_TOKEN + "\n\n");
      res.end();
      clearInterval(timer);
    }
  }, 120);
  return timer;
}

http
  .createServer((req, res) => {
    Object.entries(CORS).forEach(([k, v]) => res.setHeader(k, v));

    if (req.method === "OPTIONS") {
      res.writeHead(204);
      res.end();
      return;
    }

    if (req.method === "GET" && req.url.startsWith("/v1/models")) {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ object: "list", data: [{ id: "mock-model", object: "model" }] }));
      return;
    }

    if (req.method === "POST" && req.url.startsWith("/v1/chat/completions")) {
      const timer = handleChat(res);
      // NOTE: watch the response, not the request — in modern Node the
      // request "close" fires as soon as the body is consumed.
      res.on("close", () => clearInterval(timer));
      req.resume();
      return;
    }

    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "not found" }));
  })
  .listen(PORT, () => console.log("Mock vLLM on http://localhost:" + PORT));
