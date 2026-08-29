/*
 * Shared LLM forwarding logic. Used in two places:
 *   1. webpack dev-server middleware  -> `npm start` alone bridges the pane to
 *      the local LLM server over the SAME origin (no CORS, no second terminal):
 *        /vllm     -> VLLM_URL     (default http://localhost:8000)
 *        /ollama   -> OLLAMA_URL   (default http://localhost:11434)
 *        /lmstudio -> LMSTUDIO_URL (default http://localhost:1234)
 *        /bridge   -> target from the x-llm-target request header (dynamic;
 *                     private/tailnet hosts only, see isAllowedBridgeTarget)
 *   2. llm-proxy.js (optional standalone HTTPS proxy on :4001) for custom setups;
 *      it also honors x-llm-target when dynamicTarget is enabled.
 *
 * Everything stays local / on your tailnet: requests are only forwarded to the
 * configured upstream or to a private-network target from the pane.
 */
const http = require("http");
const https = require("https");

function octets(host) {
  return host.split(".").map((p) => parseInt(p, 10));
}

/**
 * Guardrail for dynamic bridging: only loopback, RFC1918 LAN ranges,
 * the Tailscale CGNAT range (100.64.0.0/10), *.ts.net and *.local are allowed.
 * Public hosts are rejected so the dev server never becomes an open proxy.
 */
function isAllowedBridgeTarget(target) {
  let u;
  try {
    u = new URL(target);
  } catch {
    return false;
  }
  if (u.protocol !== "http:" && u.protocol !== "https:") {
    return false;
  }
  const host = u.hostname.toLowerCase();
  if (host === "localhost" || host.endsWith(".localhost") || host === "::1" || host === "[::1]") {
    return true;
  }
  if (host.endsWith(".ts.net") || host.endsWith(".local")) {
    return true;
  }
  if (!/^[0-9.]+$/.test(host)) {
    return false; // non-literal hostnames must be an allowlisted suffix
  }
  const parts = host.split(".").map((p) => parseInt(p, 10));
  if (parts.length !== 4 || parts.some((n) => isNaN(n) || n < 0 || n > 255)) {
    return false;
  }
  const [a, b] = parts;
  if (a === 127 || a === 10 || a === 192 && b === 168) {
    return true;
  }
  if (a === 172 && b >= 16 && b <= 31) {
    return true;
  }
  if (a === 100 && b >= 64 && b <= 127) {
    return true; // Tailscale CGNAT range
  }
  return false;
}

function joinUpstreamPath(upstream, targetPath) {
  if (upstream.pathname === "/") {
    return targetPath;
  }
  if (upstream.pathname.endsWith("/")) {
    return upstream.pathname.slice(0, -1) + targetPath;
  }
  return upstream.pathname + targetPath;
}

function createForwardHandler(vllmUrl, opts) {
  const options = opts || {};
  const stripPrefix = options.stripPrefix !== false; // standalone proxy strips /vllm itself
  const dynamicTarget = options.dynamicTarget === true;
  const fixedUpstream = vllmUrl ? new URL(vllmUrl) : null;

  return function forward(req, res) {
    // Same-origin calls do not need CORS, but the standalone proxy does.
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader(
      "Access-Control-Allow-Headers",
      "Content-Type, Authorization, X-Llm-Target"
    );

    if (req.method === "OPTIONS") {
      res.writeHead(204);
      res.end();
      return;
    }

    // Resolve the upstream for this request.
    let upstream = fixedUpstream;
    let insecureTlsOk = false;
    if (dynamicTarget) {
      const target = req.headers["x-llm-target"];
      if (typeof target === "string" && target.trim()) {
        if (!isAllowedBridgeTarget(target.trim())) {
          res.writeHead(403, { "Content-Type": "application/json" });
          res.end(
            JSON.stringify({
              error:
                "Bridge refused: target is not a private/tailnet address. Allowed: localhost, 127.0.0.0/8, 10.0.0.0/8, 172.16.0.0/12, 192.168.0.0/16, 100.64.0.0/10 (Tailscale), *.ts.net, *.local.",
            })
          );
          return;
        }
        upstream = new URL(target.trim());
        // Private servers often use self-signed certificates; this is a
        // user-directed, machine-local bridge, so tolerate them.
        insecureTlsOk = upstream.protocol === "https:";
      } else if (fixedUpstream) {
        // No header: fall back to the fixed upstream (standalone proxy mode).
      } else {
        res.writeHead(400, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "Missing x-llm-target header." }));
        return;
      }
    }

    let targetPath = req.url || "/";
    if (stripPrefix && targetPath.startsWith("/vllm")) {
      targetPath = targetPath.slice(5) || "/";
    }
    const fullPath = joinUpstreamPath(upstream, targetPath);

    const lib = upstream.protocol === "https:" ? https : http;
    const upstreamReq = lib.request(
      {
        protocol: upstream.protocol,
        hostname: upstream.hostname,
        port: upstream.port || (upstream.protocol === "https:" ? 443 : 80),
        path: fullPath,
        method: req.method,
        headers: { ...req.headers, host: upstream.host },
        ...(insecureTlsOk ? { rejectUnauthorized: false } : {}),
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
          error:
            "Cannot reach LLM server at " +
            upstream.origin +
            (upstream.pathname === "/" ? "" : upstream.pathname) +
            ": " +
            (err && err.message ? err.message : String(err)),
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

module.exports = { createForwardHandler, isAllowedBridgeTarget };
