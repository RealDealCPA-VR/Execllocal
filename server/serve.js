/*
 * ExcelLocal production server (headless mode).
 *
 * Serves the built pane (dist/) AND the same-origin LLM bridges from one port,
 * with no webpack and no terminal on the machine you actually use Excel on.
 *
 *   node server/serve.js            (http://localhost:3000)
 *
 * Env:
 *   SERVE_PORT   listen port (default 3000)
 *   SERVE_HOST   bind host   (default 127.0.0.1; set 0.0.0.0 if fronting from off-box)
 *   VLLM_URL / OLLAMA_URL / LMSTUDIO_URL   bridge upstreams (same as npm start)
 *
 * Office requires the pane over HTTPS. Front this HTTP port with `tailscale serve`
 * (or Caddy) to get a trusted certificate, e.g.:
 *
 *   tailscale serve --bg --https=443 http://127.0.0.1:3000
 *
 * Then generate + sideload the remote manifest from your Windows machine:
 *   npm run manifest:remote -- https://<your-box>.<tailnet>.ts.net
 *   npm run sideload:remote
 */
const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { createForwardHandler } = require("../llm-forward");

const PORT = Number(process.env.SERVE_PORT || 3000);
const HOST = process.env.SERVE_HOST || "127.0.0.1";
const DIST = path.join(__dirname, "..", "dist");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".map": "application/json",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
};

function serveStatic(req, res) {
  if (req.method !== "GET" && req.method !== "HEAD") {
    res.writeHead(405).end();
    return;
  }
  let urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  if (urlPath === "/" || urlPath === "") {
    urlPath = "/taskpane.html";
  }
  const resolved = path.normalize(path.join(DIST, urlPath));
  // Path traversal guard. The separator matters: a sibling directory named
  // "dist-something" also starts with DIST.
  if (resolved !== DIST && !resolved.startsWith(DIST + path.sep)) {
    res.writeHead(403).end();
    return;
  }
  fs.readFile(resolved, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" }).end("Not found: " + urlPath);
      return;
    }
    const ext = path.extname(resolved).toLowerCase();
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "application/octet-stream",
      // Excel's WebView caches aggressively: without this a rebuilt bundle
      // keeps serving the old pane until the user clears the WebView cache.
      "Cache-Control": ext === ".png" || ext === ".ico" ? "public, max-age=3600" : "no-cache, must-revalidate",
    });
    res.end(data);
  });
}

const bridges = {
  "/vllm": createForwardHandler(process.env.VLLM_URL || "http://localhost:8000", { stripPrefix: false }),
  "/ollama": createForwardHandler(process.env.OLLAMA_URL || "http://localhost:11434", { stripPrefix: false }),
  "/lmstudio": createForwardHandler(process.env.LMSTUDIO_URL || "http://localhost:1234", { stripPrefix: false }),
  "/bridge": createForwardHandler(null, { stripPrefix: false, dynamicTarget: true }),
};

function requestHandler(req, res) {
  const url = req.url || "/";
  const pathOnly = url.split("?")[0];
  for (const prefix of Object.keys(bridges)) {
    if (pathOnly === prefix || pathOnly.startsWith(prefix + "/")) {
      req.url = url.slice(prefix.length) || "/";
      return bridges[prefix](req, res);
    }
  }
  serveStatic(req, res);
}

// Office requires the pane over HTTPS. Use the Office dev certificates when they
// exist (same ones webpack uses - already trusted by Excel). Set SERVE_TLS=0 to
// serve plain HTTP, e.g. when fronting with `tailscale serve` on the GPU box.
const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const crtPath = path.join(certDir, "localhost.crt");
const keyPath = path.join(certDir, "localhost.key");
const useTls = process.env.SERVE_TLS !== "0" && fs.existsSync(crtPath) && fs.existsSync(keyPath);
const server = useTls
  ? https.createServer({ cert: fs.readFileSync(crtPath), key: fs.readFileSync(keyPath) }, requestHandler)
  : http.createServer(requestHandler);
if (!fs.existsSync(path.join(DIST, "taskpane.html"))) {
  console.error("dist/taskpane.html is missing - run `npm run build` before `npm run serve`.");
  process.exit(1);
}

server.on("error", (err) => {
  console.error(
    err && err.code === "EADDRINUSE"
      ? "Port " + PORT + " is already in use. Set SERVE_PORT to another port."
      : "Server error: " + (err && err.message ? err.message : String(err))
  );
  process.exit(1);
});

server.listen(PORT, HOST, () => {
  console.log("ExcelLocal production server:");
  console.log("  pane:    " + (useTls ? "https" : "http") + "://" + HOST + ":" + PORT + "/taskpane.html" + (useTls ? "  (Office dev certs)" : "  (front with tailscale serve / Caddy for HTTPS)"));
  console.log("  bridges: /vllm /ollama /lmstudio /bridge");
  console.log("  static:  " + DIST);
});
