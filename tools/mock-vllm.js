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
// Optional path prefix, to test proxies that forward with a base path (e.g. MOCK_BASE=/api).
const BASE = process.env.MOCK_BASE || "";
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

    if (req.method === "GET" && req.url.startsWith(BASE + "/v1/models")) {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ object: "list", data: [{ id: "mock-model", object: "model" }] }));
      return;
    }

    if (req.method === "POST" && req.url.startsWith(BASE + "/v1/chat/completions")) {
      if (process.env.MOCK_AGENT === "1") {
        handleAgentChat(req, res);
        return;
      }
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

/*
 * Scripted agent mode (MOCK_AGENT=1): stateless two-turn tool conversation.
 * Turn 1 (no role:"tool" in messages)  -> stream a read_range tool_call
 * Turn 2 (tool result present)         -> stream final content
 */
function handleAgentChat(req, res) {
  let body = "";
  req.on("data", (c) => (body += c));
  req.on("end", () => {
    let hadToolResult = false;
    try {
      const parsed = JSON.parse(body || "{}");
      hadToolResult = (parsed.messages || []).some((m) => m.role === "tool");
    } catch {
      hadToolResult = false;
    }
    startSse(res);
    if (!hadToolResult) {
      const name = "read_range";
      const args = JSON.stringify({ sheet: "Sheet1", address: "A1:B2" });
      // stream: start (id+name), then args in three fragments
      sseLine(res, {
        id: "mock",
        object: "chat.completion.chunk",
        choices: [{
          index: 0,
          delta: { tool_calls: [{ index: 0, id: "call_live_1", type: "function", function: { name: name, arguments: "" } }] },
          finish_reason: null,
        }],
      });
      const pieces = [args.slice(0, 8), args.slice(8, 16), args.slice(16)];
      let i = 0;
      const timer = setInterval(() => {
        if (i < pieces.length) {
          sseLine(res, {
            id: "mock",
            object: "chat.completion.chunk",
            choices: [{ index: 0, delta: { tool_calls: [{ index: 0, function: { arguments: pieces[i++] } }] }, finish_reason: null }],
          });
        } else {
          sseLine(res, {
            id: "mock",
            object: "chat.completion.chunk",
            choices: [{ index: 0, delta: {}, finish_reason: "tool_calls" }],
          });
          res.write(DONE_TOKEN + "\n\n");
          res.end();
          clearInterval(timer);
        }
      }, 60);
      res.on("close", () => clearInterval(timer));
    } else {
      const tokens = ["Live", " agent", " reply", " after", " tool", " use."];
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
      }, 40);
      res.on("close", () => clearInterval(timer));
    }
  });
}
