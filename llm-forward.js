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
const net = require("net");
const https = require("https");

/**
 * Guardrail for dynamic bridging (personal-use semantics):
 *  - any HOSTNAME is allowed (your machine resolves it: Tailscale MagicDNS short names,
 *    *.ts.net, *.local, LAN names, FQDNs)
 *  - IP literals must be loopback or private ranges (RFC1918, Tailscale CGNAT 100.64/10,
 *    IPv6 ULA fc00::/7) - direct public IP literals are refused
 * The dev server only listens on localhost, so this bridge is reachable from your machine
 * alone; the guard exists to keep it from being usable as a general internet proxy.
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
  // URL keeps IPv6 literals bracketed ("[::1]"); net.isIP needs them bare.
  const host = u.hostname.toLowerCase().replace(/^\[|\]$/g, "");
  if (host === "localhost" || host.endsWith(".localhost")) {
    return true;
  }
  const ipVersion = net.isIP(host);
  if (ipVersion === 0) {
    return true; // hostname: resolved by this machine's own DNS (MagicDNS, LAN, FQDN)
  }
  if (ipVersion === 6) {
    // loopback (::1), IPv4-mapped loopback, and ULA fc00::/7 only
    return host === "::1" || host === "::ffff:127.0.0.1" || /^f[cd][0-9a-f]{0,2}:/.test(host);
  }
  // IPv4 literal: loopback + private ranges only
  const parts = host.split(".").map((p) => parseInt(p, 10));
  if (parts.length !== 4 || parts.some((n) => isNaN(n) || n < 0 || n > 255)) {
    return false;
  }
  const [a, b] = parts;
  if (a === 127 || a === 10 || (a === 192 && b === 168)) {
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

/**
 * Node's happy-eyeballs connect failures arrive as an AggregateError whose own
 * `message` is empty, which surfaced in the pane as a bare "AggregateError".
 * Unwrap it so the user sees ECONNREFUSED and the address that was tried.
 */
function describeConnectError(err) {
  if (!err) {
    return "unknown error";
  }
  const parts = [];
  const add = (msg, code) => {
    if (!msg && !code) {
      return;
    }
    parts.push(msg || code);
  };
  add(err.message, err.code);
  for (const e of Array.isArray(err.errors) ? err.errors : []) {
    add(e && e.message, e && e.code);
  }
  const unique = Array.from(new Set(parts.filter(Boolean)));
  return unique.length ? unique.join("; ") : err.code || String(err);
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
                "Bridge refused: public IP literals are not allowed. Use a hostname (Tailscale MagicDNS name, *.ts.net, LAN name) or a private-range IP (localhost, 10.x, 172.16-31.x, 192.168.x, 100.64-127.x).",
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
    // Tolerate OpenAI-style base URLs that include a trailing /v1 (the client
    // appends /v1/... itself, so a /v1 suffix on the target would double up).
    let pathPrefix = upstream.pathname;
    if (dynamicTarget && pathPrefix.toLowerCase().endsWith("/v1")) {
      pathPrefix = pathPrefix.slice(0, -3) || "/";
    }
    const fullPath = joinUpstreamPath({ ...upstream, pathname: pathPrefix }, targetPath);

    // Do not leak the bridge control header (or hop-by-hop headers) upstream.
    const outHeaders = { ...req.headers, host: upstream.host };
    delete outHeaders["x-llm-target"];
    delete outHeaders.connection;
    delete outHeaders["keep-alive"];
    delete outHeaders["proxy-connection"];
    delete outHeaders["transfer-encoding"];
    delete outHeaders.origin;
    delete outHeaders.referer;

    const lib = upstream.protocol === "https:" ? https : http;
    const upstreamReq = lib.request(
      {
        protocol: upstream.protocol,
        // URL.hostname keeps IPv6 brackets; net.connect wants the bare address.
        hostname: upstream.hostname.replace(/^\[|\]$/g, ""),
        port: upstream.port || (upstream.protocol === "https:" ? 443 : 80),
        path: fullPath,
        method: req.method,
        headers: outHeaders,
        ...(insecureTlsOk ? { rejectUnauthorized: false } : {}),
      },
      (upstreamRes) => {
        res.writeHead(upstreamRes.statusCode || 502, upstreamRes.headers);
        upstreamRes.pipe(res); // SSE streams pass through untouched
      }
    );

    upstreamReq.on("error", (err) => {
      if (res.writableEnded) {
        return;
      }
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
            describeConnectError(err),
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

module.exports = { createForwardHandler, isAllowedBridgeTarget, describeConnectError };
