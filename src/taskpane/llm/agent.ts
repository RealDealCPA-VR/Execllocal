/**
 * The agent loop: drives the model through tool-calling turns until it
 * produces a final answer. Deliberately free of Office.js and DOM code so it
 * can be unit-tested in Node with a fake transport and fake executor.
 */
import type { Transport, TransportOptions, WireMessage } from "./transport";
import { WRITE_TOOLS } from "./tools";

export interface ToolCall {
  id: string;
  name: string;
  args: string;
}

export interface ToolExecutor {
  execute(name: string, args: Record<string, unknown>): Promise<{ result: unknown; summary: string }>;
}

export interface AgentCallbacks {
  onContentDelta?(text: string): void;
  onReasoningDelta?(text: string): void;
  onToolStart?(call: ToolCall): void;
  onToolEnd?(call: ToolCall, summary: string, ok: boolean, declined?: boolean): void;
  /** Return a Promise<boolean>; true = approved, false = declined. */
  confirmTool?(call: ToolCall): Promise<boolean>;
}

export interface RunOptions {
  transport: Transport;
  transportOptions: Omit<TransportOptions, "tools">;
  systemPrompt: string;
  history: Array<{ role: "user" | "assistant"; content: string }>;
  userMessage: string;
  tools: unknown[];
  executor: ToolExecutor;
  callbacks?: AgentCallbacks;
  maxSteps?: number;
  signal?: AbortSignal;
}

export interface RunResult {
  /** Final assistant text (empty if the loop exhausted maxSteps). */
  content: string;
  steps: number;
  aborted: boolean;
  /** True when the loop stopped because maxSteps ran out mid-tool-chain. */
  limitReached: boolean;
}

export async function runAgent(o: RunOptions): Promise<RunResult> {
  const messages: WireMessage[] = [
    { role: "system", content: o.systemPrompt },
    ...o.history.map((h) => ({ role: h.role, content: h.content })),
    { role: "user", content: o.userMessage },
  ];

  const maxSteps = o.maxSteps ?? 12;
  let finalContent = "";
  let steps = 0;

  while (steps < maxSteps) {
    steps++;
    let content = "";
    let finishReason = "";
    const calls: Array<{ index: number; id?: string; name?: string; args: string }> = [];

    for await (const ev of o.transport.stream(messages, { ...o.transportOptions, tools: o.tools })) {
      switch (ev.type) {
        case "content-delta":
          content += ev.text;
          o.callbacks?.onContentDelta?.(ev.text);
          break;
        case "reasoning-delta":
          o.callbacks?.onReasoningDelta?.(ev.text);
          break;
        case "tool-call-start": {
          const existing = calls.find((c) => c.index === ev.index);
          if (existing) {
            if (ev.id) existing.id = ev.id;
            if (ev.name) existing.name = ev.name;
          } else {
            calls.push({ index: ev.index, id: ev.id, name: ev.name, args: "" });
          }
          break;
        }
        case "tool-call-args": {
          const c = calls.find((x) => x.index === ev.index);
          if (c) {
            c.args += ev.argsDelta;
          }
          break;
        }
        case "finish":
          finishReason = ev.reason;
          break;
      }
    }

    if (o.signal?.aborted) {
      return { content, steps, aborted: true, limitReached: false };
    }

    finalContent = content;

    if (calls.length === 0) {
      return { content, steps, aborted: false, limitReached: false };
    }

    messages.push({
      role: "assistant",
      content: content || "",
      tool_calls: calls.map((c, i) => ({
        id: c.id || "call_" + i,
        type: "function" as const,
        function: { name: c.name || "unknown", arguments: c.args || "{}" },
      })),
    });

    for (let i = 0; i < calls.length; i++) {
      const call: ToolCall = {
        id: calls[i].id || "call_" + i,
        name: calls[i].name || "unknown",
        args: calls[i].args || "{}",
      };
      o.callbacks?.onToolStart?.(call);

      let args: Record<string, unknown> = {};
      try {
        args = JSON.parse(call.args);
      } catch {
        args = {};
      }

      let result: unknown;
      let summary = "";
      let ok = true;
      let declined = false;

      if (WRITE_TOOLS.has(call.name) && o.callbacks?.confirmTool) {
        const approved = await o.callbacks.confirmTool(call);
        if (!approved) {
          declined = true;
          ok = false;
          result = { declined: true };
          summary = "User declined";
        }
      }

      if (!declined) {
        try {
          const out = await o.executor.execute(call.name, args);
          result = out.result;
          summary = out.summary;
        } catch (e) {
          ok = false;
          result = { error: String((e as Error)?.message ?? e) };
          summary = "Tool error";
        }
      }

      o.callbacks?.onToolEnd?.(call, summary, ok, declined);
      messages.push({ role: "tool", tool_call_id: call.id, content: JSON.stringify(result) });
    }
  }

  return { content: finalContent, steps, aborted: false, limitReached: true };
}
