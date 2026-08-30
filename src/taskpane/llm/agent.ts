/**
 * The agent loop: drives the model through tool-calling turns until it
 * produces a final answer. Deliberately free of Office.js and DOM code so it
 * can be unit-tested in Node with a fake transport and fake executor.
 */
/* global AbortSignal */
import type { Transport, TransportOptions, WireMessage } from "./transport";
import { WRITE_TOOLS } from "./tools";

export interface ToolCall {
  id: string;
  name: string;
  args: string;
}

export interface ToolExecutor {
  execute(
    name: string,
    args: Record<string, unknown>
  ): Promise<{ result: unknown; summary: string }>;
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
  /** True when the model stopped because it hit the server's token limit. */
  truncated?: boolean;
}

/** An abort surfaces as a DOMException/Error named AbortError, not as a flag. */
function isAbortError(e: unknown): boolean {
  const err = e as { name?: string; message?: string };
  return err?.name === "AbortError" || /aborted/i.test(String(err?.message ?? ""));
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
  let truncated = false;

  while (steps < maxSteps) {
    if (o.signal?.aborted) {
      return { content: finalContent, steps, aborted: true, limitReached: false, truncated };
    }
    steps++;
    let content = "";
    let finishReason = "";
    const calls: Array<{ index: number; id?: string; name?: string; args: string }> = [];

    try {
      for await (const ev of o.transport.stream(messages, {
        ...o.transportOptions,
        tools: o.tools,
      })) {
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
            } else {
              // Some servers omit the opening fragment; keep the arguments
              // rather than dropping the call entirely.
              calls.push({ index: ev.index, args: ev.argsDelta });
            }
            break;
          }
          case "finish":
            // "done" is the [DONE] sentinel and must not mask a real reason.
            if (ev.reason && ev.reason !== "done") {
              finishReason = ev.reason;
            }
            break;
        }
      }
    } catch (e) {
      if (isAbortError(e) || o.signal?.aborted) {
        return {
          content: content || finalContent,
          steps,
          aborted: true,
          limitReached: false,
          truncated,
        };
      }
      throw e;
    }

    if (o.signal?.aborted) {
      return {
        content: content || finalContent,
        steps,
        aborted: true,
        limitReached: false,
        truncated,
      };
    }

    if (finishReason === "length") {
      truncated = true;
    }
    finalContent = content;

    if (calls.length === 0) {
      return { content, steps, aborted: false, limitReached: false, truncated };
    }

    // Servers may stream fragments out of order; the wire protocol pairs the
    // assistant tool_calls with the tool results by id, so ids must be unique
    // and identical in both places.
    calls.sort((a, b) => a.index - b.index);
    const seenIds = new Set<string>();
    const resolved: ToolCall[] = calls.map((c, i) => {
      let id = c.id || "call_" + i;
      if (seenIds.has(id)) {
        id = id + "_" + i;
      }
      seenIds.add(id);
      return { id, name: c.name || "unknown", args: c.args || "{}" };
    });

    messages.push({
      role: "assistant",
      content: content || "",
      tool_calls: resolved.map((c) => ({
        id: c.id,
        type: "function" as const,
        function: { name: c.name, arguments: c.args },
      })),
    });

    for (const call of resolved) {
      o.callbacks?.onToolStart?.(call);

      let args: Record<string, unknown> = {};
      try {
        const parsed = JSON.parse(call.args);
        args =
          parsed && typeof parsed === "object" && !Array.isArray(parsed)
            ? (parsed as Record<string, unknown>)
            : {};
      } catch {
        args = {};
      }

      let result: unknown;
      let summary = "";
      let ok = true;
      let declined = false;

      if (o.signal?.aborted) {
        declined = true;
        ok = false;
        result = { error: "Cancelled by the user before this tool ran." };
        summary = "Cancelled";
      } else if (WRITE_TOOLS.has(call.name) && o.callbacks?.confirmTool) {
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
          const message = String((e as Error)?.message ?? e);
          result = { error: message };
          summary = message.slice(0, 140) || "Tool error";
        }
      }

      o.callbacks?.onToolEnd?.(call, summary, ok, declined);
      // Every tool_call in the assistant message needs a matching tool message
      // or the next request is rejected by OpenAI-compatible servers.
      messages.push({
        role: "tool",
        tool_call_id: call.id,
        content: JSON.stringify(result ?? null),
      });
    }

    if (o.signal?.aborted) {
      return { content: finalContent, steps, aborted: true, limitReached: false, truncated };
    }
  }

  return { content: finalContent, steps, aborted: false, limitReached: true, truncated };
}
