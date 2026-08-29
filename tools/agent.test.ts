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
    [
      ...toolCallEvents(0, "call_2", "write_range", WRITE_ARGS),
      { type: "finish", reason: "tool_calls" },
    ],
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
    [{ type: "content-delta", text: "Cancelled." }, { type: "finish", reason: "stop" }],
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
    infinite.push([...toolCallEvents(0, "call_x" + i, "get_workbook_info", "{}"), { type: "finish", reason: "tool_calls" }]);
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
  const transport = new ScriptedTransport([[{ type: "content-delta", text: "partial" }, { type: "finish", reason: "stop" }]]);
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
  console.log("PASS testAbort");
}

(async () => {
  await testToolRoundTrip();
  await testDeclinedWrite();
  await testMaxStepsGuard();

  testNormalizeBase();
  testSseExtraction();
  await testAbort();
  console.log("ALL AGENT TESTS PASSED");
})().catch((e) => {
  console.error("TEST FAILURE:", e);
  process.exit(1);
});

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

// ---- SSE extraction unit tests (Pass 2) ----
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
