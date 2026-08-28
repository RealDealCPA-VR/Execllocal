/*
 * Shared LLM forwarding logic. Used in two places:
 *   1. webpack dev-server middleware  -> `npm start` alone bridges the pane to
 *      the local LLM server over the SAME origin (no CORS, no second terminal):
 *        /vllm    -> VLLM_URL    (default http://localhost:8000)
 *        /ollama  -> OLLAMA_URL  (default http://localhost:11434)
 *        /lmstudio-> LMSTUDIO_URL(default http://localhost:1234)
 *   2. llm-proxy.js (optional standalone HTTPS proxy on :4001) for custom setups.
 */
const http = require("http");
const https = require("https");

function createForwardHandler(vllmUrl, opts) {
  const options = opts || {};
  const stripPrefix = options.stripPrefix !== false; // standalone proxy strips /vllm itself
  const upstream = new URL(vllmUrl);
  const lib = upstream.protocol === "https:" ? https : http;

  return function forward(req, res) {
    // Same-origin calls do not need CORS, but the standalone proxy does.
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

    if (req.method === "OPTIONS") {
      res.writeHead(204);
      res.end();
      return;
    }

    let targetPath = req.url || "/";
    if (stripPrefix && targetPath.startsWith("/vllm")) {
      targetPath = targetPath.slice(5) || "/";
    }

    // Preserve any path prefix configured in the upstream URL.
    let fullPath;
    if (upstream.pathname === "/") {
      fullPath = targetPath;
    } else if (upstream.pathname.endsWith("/")) {
      fullPath = upstream.pathname.slice(0, -1) + targetPath;
    } else {
      fullPath = upstream.pathname + targetPath;
    }

    const upstreamReq = lib.request(
      {
        protocol: upstream.protocol,
        hostname: upstream.hostname,
        port: upstream.port || (upstream.protocol === "https:" ? 443 : 80),
        path: fullPath,
        method: req.method,
        headers: { ...req.headers, host: upstream.host },
      },
      (upstreamRes) => {
        res.writeHead(upstreamRes.statusCode || 502, upstreamRes.headers);
        upstreamRes.pipe(res); // SSE streams pass through untouched
      }
    );

    upstreamReq.on("error", (err) => {
      if (!res.headersSent) {
        res.writeHead(502, { "Content-Type": "application/json" });
      }
      res.end(
        JSON.stringify({
          error: "Cannot reach LLM server at " + vllmUrl + ": " + (err && err.message ? err.message : String(err)),
        })
      );
    });

    // If the client aborts (Stop button), kill the upstream request.
    const abortUpstream = () => upstreamReq.destroy();
    res.on("close", abortUpstream);
    req.on("aborted", abortUpstream);

    req.pipe(upstreamReq);
  };
}

module.exports = { createForwardHandler };
