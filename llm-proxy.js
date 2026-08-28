/*
 * ExcelLocal LLM proxy.
 *
 * Office requires task panes to be served over HTTPS, and an HTTPS page is
 * not allowed to call an HTTP endpoint (mixed content). Local LLM servers
 * (vLLM, Ollama, LM Studio...) speak plain HTTP, so this tiny proxy exposes
 * them over HTTPS using the same Office dev certificates that Excel already
 * trusts for the task pane itself.
 *
 *   https://localhost:4001/vllm/v1/chat/completions  ->  http://localhost:8000/v1/chat/completions
 *
 * Everything stays local: the proxy forwards to your machine's LLM server
 * and nothing is sent anywhere else.
 *
 * Env overrides:
 *   VLLM_URL    upstream LLM server (default http://localhost:8000)
 *   PROXY_PORT  listen port        (default 4001)
 */
const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PROXY_PORT = Number(process.env.PROXY_PORT || 4001);
const VLLM_URL = process.env.VLLM_URL || "http://localhost:8000";

async function getTlsOptions() {
  // Prefer the Office dev certificates that already exist on disk — loading
  // them directly avoids re-triggering the CA-install prompt on every start.
  // (First `npm start` may still ask once to trust "Developer CA for
  //  Microsoft Office Add-ins"; accept it and everything works afterwards.)
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const crt = path.join(certDir, "localhost.crt");
  const key = path.join(certDir, "localhost.key");
  if (fs.existsSync(crt) && fs.existsSync(key)) {
    return { cert: fs.readFileSync(crt), key: fs.readFileSync(key) };
  }
  const devCerts = require("office-addin-dev-certs");
  return await devCerts.getHttpsServerOptions();
}

async function main() {
  const upstream = new URL(VLLM_URL);
  const tlsOptions = await getTlsOptions();

  const server = https.createServer(tlsOptions, (clientReq, clientRes) => {
    // The task pane origin differs from ours; answer every preflight and
    // tag every response so the browser accepts the call.
    clientRes.setHeader("Access-Control-Allow-Origin", "*");
    clientRes.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    clientRes.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

    if (clientReq.method === "OPTIONS") {
      clientRes.writeHead(204);
      clientRes.end();
      return;
    }

    // Strip the /vllm prefix: /vllm/v1/chat/completions -> /v1/chat/completions
    const targetPath = clientReq.url.replace(/^\/vllm/, "") || "/";

    const upstreamReq = http.request(
      {
        protocol: upstream.protocol,
        hostname: upstream.hostname,
        port: upstream.port || (upstream.protocol === "https:" ? 443 : 80),
        path: targetPath,
        method: clientReq.method,
        headers: { ...clientReq.headers, host: upstream.host },
      },
      (upstreamRes) => {
        clientRes.writeHead(upstreamRes.statusCode || 502, upstreamRes.headers);
        upstreamRes.pipe(clientRes); // SSE streams pass through untouched
      }
    );

    upstreamReq.on("error", (err) => {
      if (!clientRes.headersSent) {
        clientRes.writeHead(502, { "Content-Type": "application/json" });
      }
      clientRes.end(
        JSON.stringify({
          error: `Cannot reach LLM server at ${VLLM_URL}: ${err.message}`,
        })
      );
    });

    // If the user hits "Stop" in the chat, kill the upstream request so the
    // GPU stops generating.
    const abortUpstream = () => upstreamReq.destroy();
    clientRes.on("close", abortUpstream);
    clientReq.on("aborted", abortUpstream);

    clientReq.pipe(upstreamReq);
  });

  server.listen(PROXY_PORT, () => {
    console.log(`ExcelLocal LLM proxy listening on https://localhost:${PROXY_PORT}/vllm`);
    console.log(`Forwarding to ${VLLM_URL}`);
  });
}

main().catch((err) => {
  console.error("Failed to start proxy:", err);
  process.exitCode = 1;
});
