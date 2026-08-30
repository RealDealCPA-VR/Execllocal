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
const fs = require("fs");
const path = require("path");
const os = require("os");
const { createForwardHandler } = require("./llm-forward");

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
  // Only required lazily: on a box that already has the certs on disk this
  // module (and its CA-install prompt) is never loaded.
  const devCerts = require("office-addin-dev-certs");
  return await devCerts.getHttpsServerOptions();
}

async function main() {
  const tlsOptions = await getTlsOptions();

  const forward = createForwardHandler(VLLM_URL, { stripPrefix: true, dynamicTarget: true });

  const server = https.createServer(tlsOptions, (clientReq, clientRes) => {
    forward(clientReq, clientRes);
  });

  server.on("error", (err) => {
    console.error(
      err && err.code === "EADDRINUSE"
        ? "Port " + PROXY_PORT + " is already in use. Set PROXY_PORT to another port."
        : "Proxy error: " + (err && err.message ? err.message : String(err))
    );
    process.exit(1);
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
