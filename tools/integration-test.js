/*
 * Integration test: REAL network stack, no Excel.
 *   real HttpTransport (fetch + SSE parse)
 *     -> https://localhost:4001 (TLS proxy, self-signed dev certs)
 *       -> mock vLLM in MOCK_AGENT=1 mode (streams a tool_call, then a final answer)
 *   driven by the REAL runAgent loop with a fake executor.
 * Run: node tools/integration-test.js   (after: npx tsc -p tools/tsconfig.test.json)
 */
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0"; // self-signed Office dev certs
const { spawn } = require("child_process");
const assert = require("assert");
const net = require("net");

// Pick free ports dynamically so stray servers can never collide with the test.
function getFreePort() {
  return new Promise((resolve, reject) => {
    const srv = net.createServer();
    srv.unref();
    srv.on("error", reject);
    srv.listen(0, "127.0.0.1", () => {
      const port = srv.address().port;
      srv.close(() => resolve(port));
    });
  });
}
// NOTE: mock and proxy bind to dynamically chosen free ports.
const { runAgent } = require("./build-test/src/taskpane/llm/agent");
const { HttpTransport } = require("./build-test/src/taskpane/llm/transport");

function wait(ms) { return new Promise((r) => setTimeout(r, ms)); }

function start(cmd, args, env) {
  const child = spawn(cmd, args, { env: { ...process.env, ...env }, stdio: ["ignore", "pipe", "pipe"] });
  child.stdout.on("data", () => {});
  child.stderr.on("data", (d) => process.stderr.write("[child] " + d));
  return child;
}

(async () => {
  const mockPort = await getFreePort();
  const proxyPort = await getFreePort();
  const mock = start("node", ["tools/mock-vllm.js"], { MOCK_AGENT: "1", MOCK_PORT: String(mockPort) });
  const proxy = start("node", ["llm-proxy.js"], { VLLM_URL: "http://localhost:" + mockPort, PROXY_PORT: String(proxyPort) });
  await wait(2500);

  try {
    const executed = [];
    const transport = new HttpTransport();
    const result = await runAgent({
      transport,
      transportOptions: { baseUrl: "https://localhost:" + proxyPort + "/vllm", model: "mock-model", temperature: 0 },
      systemPrompt: "integration test",
      history: [],
      userMessage: "please inspect A1:B2",
      tools: [{ type: "function", function: { name: "read_range", parameters: { type: "object", properties: {} } } }],
      executor: {
        execute: async (name, args) => {
          executed.push({ name, args });
          return { result: { ok: true, rows: 2 }, summary: "read " + args.sheet + "!" + args.address };
        },
      },
      callbacks: { confirmTool: async () => true },
    });

    assert.strictEqual(result.content, "Live agent reply after tool use.");
    assert.strictEqual(result.limitReached, false);
    assert.strictEqual(executed.length, 1);
    assert.strictEqual(executed[0].name, "read_range");
    assert.deepStrictEqual(executed[0].args, { sheet: "Sheet1", address: "A1:B2" });
    console.log("PASS integration: real fetch -> TLS proxy -> SSE tool_call -> agent loop -> executor -> final answer");
  } finally {
    mock.kill();
    proxy.kill();
  }

  // Second pass: a server that ignores `stream: true` and returns one JSON body.
  const plainPort = await getFreePort();
  const proxy2Port = await getFreePort();
  const plain = start("node", ["tools/mock-vllm.js"], { MOCK_NONSTREAM: "1", MOCK_PORT: String(plainPort) });
  const proxy2 = start("node", ["llm-proxy.js"], { VLLM_URL: "http://localhost:" + plainPort, PROXY_PORT: String(proxy2Port) });
  await wait(2500);
  try {
    const result = await runAgent({
      transport: new HttpTransport(),
      transportOptions: { baseUrl: "https://localhost:" + proxy2Port + "/vllm", model: "mock-model", temperature: 0 },
      systemPrompt: "integration test",
      history: [],
      userMessage: "hello",
      tools: [],
      executor: { execute: async () => ({ result: {}, summary: "" }) },
    });
    assert.strictEqual(result.content, "Whole-body reply.");
    assert.strictEqual(result.aborted, false);
    console.log("PASS integration: non-streaming server falls back to whole-body parsing");
    console.log("ALL INTEGRATION CHECKS PASSED");
  } finally {
    plain.kill();
    proxy2.kill();
  }
})().catch((e) => {
  console.error("INTEGRATION FAILURE:", e);
  process.exit(1);
});
