/**
 * Agent-loop tests: exercises runAgent with a scripted in-process transport
 * and a recording fake executor (no Excel.js, no network).
 * Run: npx tsc -p tools/tsconfig.test.json && node tools/build-test/tools/agent.test.js
 */
import * as assert from "assert";
import { runAgent } from "../src/taskpane/llm/agent";
import type { StreamEvent, Transport, TransportOptions, WireMessage } from "../src/taskpane/llm/transport";

class ScriptedTransport implements Transport {
  public requests: WireMessage[][] = [];
  private turn = 0;
  constructor(private script: StreamEvent[][]) {}
  async *stream(messages: WireMessage[], _opts: TransportOptions): AsyncIterable<StreamEvent> {
    this.requests.push(JSON.parse(JSON.stringify(messages)));
    const events = this.script[this.turn++] ?? [{ type: "finish", reason: "stop" }];
    for (const ev of events) {
      yield ev;
    }
  }
}

function toolCallEvents(index: number, id: string, name: string, argsJson: string): StreamEvent[] {
  const events: StreamEvent[] = [{ type: "tool-call-start", index, id, name }];
  for (let i = 0; i < argsJson.length; i += 7) {
    events.push({ type: "tool-call-args", index, argsDelta: argsJson.slice(i, i + 7) });
  }
  return events;
}

function makeExecutor() {
  const executed: Array<{ name: string; args: Record<string, unknown> }> = [];
  return {
    executed,
    execute: async (name: string, args: Record<string, unknown>) => {
      executed.push({ name, args });
      return { result: { ok: true, tool: name }, summary: "executed " + name };
    },
  };
}

const READ_ARGS = '{"sheet":"Sheet1","address":"A1:B2"}';
const WRITE_ARGS = '{"sheet":"Sheet1","address":"C1","values":[[4]]}';

async function testToolRoundTrip(): Promise<void> {
  const transport = new ScriptedTransport([
    [
      { type: "content-delta", text: "Let me look. " },
      ...toolCallEvents(0, "call_1", "read_range", READ_ARGS),
      { type: "finish", reason: "tool_calls" },
    ],
    [...toolCallEvents(0, "call_2", "write_range", WRITE_ARGS), { type: "finish", reason: "tool_calls" }],
    [
      { type: "content-delta", text: "Done: wrote 4." },
      { type: "finish", reason: "stop" },
    ],
  ]);
  const { execute, executed } = makeExecutor();
  const toolEnds: string[] = [];
  let confirmCalls = 0;

  const result = await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "read then write",
    tools: [],
    executor: { execute },
    callbacks: {
      confirmTool: async () => {
        confirmCalls++;
        return true;
      },
      onToolEnd: (call, summary) => toolEnds.push(call.name + ":" + summary),
    },
  });

  assert.strictEqual(result.content, "Done: wrote 4.");
  assert.strictEqual(result.limitReached, false);
  assert.strictEqual(transport.requests.length, 3);
  assert.strictEqual(confirmCalls, 1); // only write_range confirmed
  assert.strictEqual(executed.length, 2);

  // Second request must carry the assistant tool_calls + tool result.
  const req2 = transport.requests[1];
  assert.strictEqual(req2.length, 4);
  assert.strictEqual(req2[2].role, "assistant");
  assert.strictEqual(req2[2].tool_calls?.[0].id, "call_1");
  assert.strictEqual(req2[2].tool_calls?.[0].function.name, "read_range");
  assert.strictEqual(req2[2].tool_calls?.[0].function.arguments, READ_ARGS);
  assert.strictEqual(req2[3].role, "tool");
  assert.strictEqual(req2[3].tool_call_id, "call_1");
  assert.deepStrictEqual(JSON.parse(req2[3].content), { ok: true, tool: "read_range" });

  // Third request carries both tool results in order.
  const req3 = transport.requests[2];
  assert.strictEqual(req3.length, 6);
  assert.strictEqual(req3[5].role, "tool");
  assert.strictEqual(req3[5].tool_call_id, "call_2");

  assert.strictEqual(toolEnds.length, 2);
  assert.ok(toolEnds[0].startsWith("read_range:"));
  console.log("PASS testToolRoundTrip");
}

async function testDeclinedWrite(): Promise<void> {
  const transport = new ScriptedTransport([
    [...toolCallEvents(0, "call_9", "delete_sheet", '{"name":"Data"}'), { type: "finish", reason: "tool_calls" }],
    [
      { type: "content-delta", text: "Cancelled." },
      { type: "finish", reason: "stop" },
    ],
  ]);
  const { execute, executed } = makeExecutor();

  const result = await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "delete sheet",
    tools: [],
    executor: { execute },
    callbacks: { confirmTool: async () => false },
  });

  assert.strictEqual(result.content, "Cancelled.");
  assert.strictEqual(executed.length, 0); // executor never ran
  const toolMsg = transport.requests[1][3];
  assert.strictEqual(toolMsg.role, "tool");
  assert.deepStrictEqual(JSON.parse(toolMsg.content), { declined: true });
  console.log("PASS testDeclinedWrite");
}

async function testMaxStepsGuard(): Promise<void> {
  const infinite: StreamEvent[][] = [];
  for (let i = 0; i < 20; i++) {
    infinite.push([
      ...toolCallEvents(0, "call_x" + i, "get_workbook_info", "{}"),
      { type: "finish", reason: "tool_calls" },
    ]);
  }
  const transport = new ScriptedTransport(infinite);
  const { execute } = makeExecutor();

  const result = await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "loop forever",
    tools: [],
    executor: { execute },
    maxSteps: 3,
  });

  assert.strictEqual(result.steps, 3);
  assert.strictEqual(result.limitReached, true);
  assert.strictEqual(transport.requests.length, 3);
  console.log("PASS testMaxStepsGuard");
}

async function testAbort(): Promise<void> {
  const transport = new ScriptedTransport([
    [
      { type: "content-delta", text: "partial" },
      { type: "finish", reason: "stop" },
    ],
  ]);
  const controller = new AbortController();
  controller.abort();
  const result = await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "hi",
    tools: [],
    executor: { execute: makeExecutor().execute },
    signal: controller.signal,
  });
  assert.strictEqual(result.aborted, true);
  // Aborting before the first request must not issue one.
  assert.strictEqual(transport.requests.length, 0);
  console.log("PASS testAbort");
}

// ---- an abort raised by the transport resolves, it does not reject ----
async function testAbortDuringStream(): Promise<void> {
  const controller = new AbortController();
  const transport: Transport = {
    // eslint-disable-next-line require-yield
    async *stream(): AsyncIterable<StreamEvent> {
      controller.abort();
      const err = new Error("The operation was aborted.");
      err.name = "AbortError";
      throw err;
    },
  };
  const result = await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "hi",
    tools: [],
    executor: { execute: makeExecutor().execute },
    signal: controller.signal,
  });
  assert.strictEqual(result.aborted, true);
  assert.strictEqual(result.limitReached, false);
  console.log("PASS testAbortDuringStream");
}

// ---- duplicate ids from the server must not desync tool_call <-> tool result ----
async function testDuplicateToolCallIds(): Promise<void> {
  const transport = new ScriptedTransport([
    [
      ...toolCallEvents(0, "dup", "read_range", READ_ARGS),
      ...toolCallEvents(1, "dup", "get_selection", "{}"),
      { type: "finish", reason: "tool_calls" },
    ],
    [
      { type: "content-delta", text: "ok" },
      { type: "finish", reason: "stop" },
    ],
  ]);
  const { execute } = makeExecutor();
  await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "two calls",
    tools: [],
    executor: { execute },
  });
  const req2 = transport.requests[1];
  const ids = (req2[2].tool_calls ?? []).map((t) => t.id);
  assert.strictEqual(ids.length, 2);
  assert.notStrictEqual(ids[0], ids[1]); // ids must be unique
  // Every tool_call needs exactly one matching tool message, in order.
  const toolIds = req2.filter((m) => m.role === "tool").map((m) => m.tool_call_id);
  assert.deepStrictEqual(toolIds, ids);
  console.log("PASS testDuplicateToolCallIds");
}

// ---- args arriving before the opening fragment must not be dropped ----
async function testOutOfOrderToolFragments(): Promise<void> {
  const transport = new ScriptedTransport([
    [
      { type: "tool-call-args", index: 1, argsDelta: '{"name":"Data"}' },
      { type: "tool-call-start", index: 0, id: "a", name: "get_selection" },
      { type: "tool-call-start", index: 1, id: "b", name: "create_sheet" },
      { type: "finish", reason: "tool_calls" },
    ],
    [
      { type: "content-delta", text: "done" },
      { type: "finish", reason: "stop" },
    ],
  ]);
  const { execute, executed } = makeExecutor();
  await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "x",
    tools: [],
    executor: { execute },
    callbacks: { confirmTool: async () => true },
  });
  assert.strictEqual(executed.length, 2);
  assert.strictEqual(executed[0].name, "get_selection"); // sorted by index
  assert.strictEqual(executed[1].name, "create_sheet");
  assert.deepStrictEqual(executed[1].args, { name: "Data" }); // args survived
  console.log("PASS testOutOfOrderToolFragments");
}

// ---- the [DONE] sentinel must not mask a real finish_reason ----
async function testTruncatedFinishReason(): Promise<void> {
  const transport = new ScriptedTransport([
    [
      { type: "content-delta", text: "half a sen" },
      { type: "finish", reason: "length" },
      { type: "finish", reason: "done" },
    ],
  ]);
  const result = await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "write an essay",
    tools: [],
    executor: { execute: makeExecutor().execute },
  });
  assert.strictEqual(result.truncated, true);
  console.log("PASS testTruncatedFinishReason");
}

// ---- a throwing executor is reported to the model, and the loop continues ----
async function testExecutorErrorIsReported(): Promise<void> {
  const transport = new ScriptedTransport([
    [...toolCallEvents(0, "e1", "read_range", READ_ARGS), { type: "finish", reason: "tool_calls" }],
    [
      { type: "content-delta", text: "recovered" },
      { type: "finish", reason: "stop" },
    ],
  ]);
  let reportedOk: boolean | undefined;
  const result = await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "read",
    tools: [],
    executor: {
      execute: async () => {
        throw new Error('The worksheet does not exist. [ItemNotFound] Existing sheets: "Sheet1".');
      },
    },
    callbacks: { onToolEnd: (_c, _s, ok) => (reportedOk = ok) },
  });
  assert.strictEqual(result.content, "recovered");
  assert.strictEqual(reportedOk, false);
  const toolMsg = transport.requests[1].find((m) => m.role === "tool");
  assert.ok(String(JSON.parse(toolMsg!.content).error).includes("ItemNotFound"));
  console.log("PASS testExecutorErrorIsReported");
}

// ---- writes run ungated when confirmation is off ----
async function testNoConfirmationMode(): Promise<void> {
  const transport = new ScriptedTransport([
    [...toolCallEvents(0, "call_w", "write_range", WRITE_ARGS), { type: "finish", reason: "tool_calls" }],
    [
      { type: "content-delta", text: "Wrote it." },
      { type: "finish", reason: "stop" },
    ],
  ]);
  const { execute, executed } = makeExecutor();
  const result = await runAgent({
    transport,
    transportOptions: { baseUrl: "mock", model: "m", temperature: 0 },
    systemPrompt: "sys",
    history: [],
    userMessage: "write",
    tools: [],
    executor: { execute },
  });
  assert.strictEqual(result.content, "Wrote it.");
  assert.strictEqual(executed.length, 1); // executed with NO confirmation gate
  console.log("PASS testNoConfirmationMode");
}

// ---- OpenAI base URL normalization unit tests ----
function testNormalizeBase(): void {
  const { normalizeOpenAiBase } = require("../src/taskpane/llm/transport");
  assert.strictEqual(normalizeOpenAiBase("http://100.66.161.52:4000/v1"), "http://100.66.161.52:4000");
  assert.strictEqual(normalizeOpenAiBase("http://100.66.161.52:4000/v1/"), "http://100.66.161.52:4000");
  assert.strictEqual(normalizeOpenAiBase("http://100.66.161.52:4000/"), "http://100.66.161.52:4000");
  assert.strictEqual(normalizeOpenAiBase("http://100.66.161.52:4000"), "http://100.66.161.52:4000");
  assert.strictEqual(normalizeOpenAiBase("http://box:4000/API/V1"), "http://box:4000/API");
  assert.strictEqual(normalizeOpenAiBase("http://box:4000/api"), "http://box:4000/api");
  assert.strictEqual(normalizeOpenAiBase("  https://gpu.tail.ts.net/v1  "), "https://gpu.tail.ts.net");
  console.log("PASS testNormalizeBase");
}

// ---- SSE extraction unit tests ----
function testSseExtraction(): void {
  const { extractSsePayloads, DONE_PAYLOAD, SSE_PREFIX } = require("../src/taskpane/llm/sse");
  const NL = String.fromCharCode(10);
  const CR = String.fromCharCode(13);

  // Splits complete lines, retains the partial tail.
  const line1 = SSE_PREFIX + JSON.stringify({ a: 1 }) + NL;
  const line2 = SSE_PREFIX + JSON.stringify({ b: 2 }) + NL;
  const partial = SSE_PREFIX + "{" + '"c';
  const r1 = extractSsePayloads(line1 + line2 + partial);
  assert.deepStrictEqual(r1.payloads, [JSON.stringify({ a: 1 }), JSON.stringify({ b: 2 })]);
  assert.strictEqual(r1.rest, partial);

  // CRLF line endings are tolerated.
  const r2 = extractSsePayloads(SSE_PREFIX + " x" + CR + NL + SSE_PREFIX + " y" + CR + NL);
  assert.deepStrictEqual(r2.payloads, ["x", "y"]);

  // Comment/keep-alive lines and blanks are ignored.
  const r3 = extractSsePayloads(": ping" + NL + NL + SSE_PREFIX + " z" + NL);
  assert.deepStrictEqual(r3.payloads, ["z"]);

  // Constants are well-formed.
  assert.strictEqual(SSE_PREFIX.length, 5); // data: (colon incl.; space handled by trim)
  assert.strictEqual(DONE_PAYLOAD, "[" + "DONE" + "]");
  console.log("PASS testSseExtraction");
}

// ---- pure helpers from the Excel tool layer ----
function testExcelHelpers(): void {
  const {
    normalizeColor,
    cleanAddress,
    cleanSheet,
    gridShape,
    reshape,
    chartType,
    sanitizeTableName,
    alignment,
  } = require("../src/taskpane/llm/excelTools");

  // A color NAME must not get a "#" glued on - Excel rejects "#red".
  assert.strictEqual(normalizeColor("red"), "red");
  assert.strictEqual(normalizeColor("#FF0000"), "#FF0000");
  assert.strictEqual(normalizeColor("FF0000"), "#FF0000");
  assert.strictEqual(normalizeColor("f00"), "#f00");
  assert.strictEqual(normalizeColor("  "), undefined);
  assert.strictEqual(normalizeColor(42), undefined);

  // Sheet-qualified addresses are a very common model output.
  assert.strictEqual(cleanAddress("Sheet1!A1:B2"), "A1:B2");
  assert.strictEqual(cleanAddress("'Q1 Sales'!A1"), "A1");
  assert.strictEqual(cleanAddress(" B2:D10 "), "B2:D10");
  assert.strictEqual(cleanAddress(undefined), "");
  assert.strictEqual(cleanSheet("'Q1 Sales'"), "Q1 Sales");

  assert.deepStrictEqual(
    gridShape([
      [1, 2],
      [3, 4],
      [5, 6],
    ]),
    { rows: 3, cols: 2 }
  );
  assert.deepStrictEqual(gridShape([1, 2, 3]), { rows: 1, cols: 3 });
  assert.strictEqual(gridShape([[1, 2], [3]]), null); // ragged
  assert.strictEqual(gridShape([]), null);

  assert.deepStrictEqual(reshape([1, 2, 3, 4], 2, 2), [
    [1, 2],
    [3, 4],
  ]);
  assert.strictEqual(reshape([1, 2, 3], 2, 2), null);

  assert.strictEqual(chartType("PIE"), "Pie");
  assert.strictEqual(chartType("column"), "ColumnClustered");
  assert.strictEqual(chartType("donut"), undefined);

  assert.strictEqual(sanitizeTableName("Sales Data"), "Sales_Data");
  assert.strictEqual(sanitizeTableName("A1"), undefined); // looks like a cell ref
  assert.strictEqual(sanitizeTableName("2024"), undefined);

  assert.strictEqual(alignment("CENTRE"), "Center");
  assert.strictEqual(alignment("justify"), undefined);
  console.log("PASS testExcelHelpers");
}

// ---- bridge target guard + connect-error unwrapping ----
function testBridgeGuard(): void {
  // Resolved from the repo root: this file is executed from tools/build-test/.
  const { isAllowedBridgeTarget, describeConnectError } = require(require("path").join(process.cwd(), "llm-forward"));
  assert.strictEqual(isAllowedBridgeTarget("http://localhost:8000"), true);
  assert.strictEqual(isAllowedBridgeTarget("http://[::1]:8000"), true); // bracketed IPv6
  assert.strictEqual(isAllowedBridgeTarget("http://100.66.1.2:4000"), true);
  assert.strictEqual(isAllowedBridgeTarget("http://192.168.1.5:1234"), true);
  assert.strictEqual(isAllowedBridgeTarget("http://8.8.8.8"), false);
  assert.strictEqual(isAllowedBridgeTarget("http://[2001:db8::1]:80"), false);
  assert.strictEqual(isAllowedBridgeTarget("ftp://localhost"), false);
  assert.strictEqual(isAllowedBridgeTarget("not a url"), false);

  // AggregateError.message is empty; the pane used to show only "AggregateError".
  const agg: Error & { errors?: unknown[]; code?: string } = new Error("");
  agg.name = "AggregateError";
  agg.code = "ECONNREFUSED";
  agg.errors = [Object.assign(new Error("connect ECONNREFUSED 127.0.0.1:8000"), { code: "ECONNREFUSED" })];
  const text = describeConnectError(agg);
  assert.ok(text.includes("ECONNREFUSED"), text);
  assert.ok(text.includes("127.0.0.1:8000"), text);
  console.log("PASS testBridgeGuard");
}

(async () => {
  await testToolRoundTrip();
  await testDeclinedWrite();
  await testMaxStepsGuard();
  await testNoConfirmationMode();
  await testAbort();
  await testAbortDuringStream();
  await testDuplicateToolCallIds();
  await testOutOfOrderToolFragments();
  await testTruncatedFinishReason();
  await testExecutorErrorIsReported();
  testNormalizeBase();
  testSseExtraction();
  testExcelHelpers();
  testBridgeGuard();
  console.log("ALL AGENT TESTS PASSED");
})().catch((e) => {
  console.error("TEST FAILURE:", e);
  process.exit(1);
});
