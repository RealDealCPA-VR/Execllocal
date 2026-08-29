/**
 * Streaming transport for OpenAI-compatible chat APIs (vLLM via the local
 * HTTPS proxy). Converts the wire format into structured StreamEvents,
 * including streamed tool_call fragments.
 */
import { DONE_PAYLOAD, extractSsePayloads } from "./sse";

export type StreamEvent =
  | { type: "content-delta"; text: string }
  | { type: "reasoning-delta"; text: string }
  | { type: "tool-call-start"; index: number; id: string; name: string }
  | { type: "tool-call-args"; index: number; argsDelta: string }
  | { type: "finish"; reason: string };

export interface WireToolCall {
  id: string;
  type: "function";
  function: { name: string; arguments: string };
}

export interface WireMessage {
  role: "system" | "user" | "assistant" | "tool";
  content: string;
  tool_calls?: WireToolCall[];
  tool_call_id?: string;
}

/**
 * Accepts OpenAI-style base URLs with or without a trailing /v1 and slashes,
 * since the client always appends /v1/... itself.
 */
export function normalizeOpenAiBase(url: string): string {
  let u = url.trim().replace(/\/+$/,"");
  if (u.toLowerCase().endsWith("/v1")) {
    u = u.slice(0, -3);
  }
  return u;
}

export interface TransportOptions {
  baseUrl: string;
  /** Extra headers (e.g. X-Llm-Target for the dynamic bridge). */
  extraHeaders?: Record<string, string>;
  apiKey?: string;
  model: string;
  temperature: number;
  tools?: unknown[];
  signal?: AbortSignal;
}

export interface Transport {
  stream(messages: WireMessage[], opts: TransportOptions): AsyncIterable<StreamEvent>;
}

export class HttpTransport implements Transport {
  async *stream(messages: WireMessage[], opts: TransportOptions): AsyncIterable<StreamEvent> {
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      ...(opts.extraHeaders ?? {}),
    };
    if (opts.apiKey) {
      headers.Authorization = "Bearer " + opts.apiKey;
    }

    const res = await fetch(opts.baseUrl + "/v1/chat/completions", {
      method: "POST",
      headers,
      body: JSON.stringify({
        model: opts.model,
        messages: messages.map((m) => {
          const wire: Record<string, unknown> = { role: m.role, content: m.content };
          if (m.tool_calls) wire.tool_calls = m.tool_calls;
          if (m.tool_call_id) wire.tool_call_id = m.tool_call_id;
          return wire;
        }),
        temperature: opts.temperature,
        stream: true,
        ...(opts.tools && opts.tools.length > 0 ? { tools: opts.tools } : {}),
      }),
      signal: opts.signal,
    });

    if (!res.ok || !res.body) {
      const detail = await res.text().catch(() => "");
      throw new Error("LLM request failed: HTTP " + res.status + (detail ? " - " + detail.slice(0, 300) : ""));
    }

    const handlePayload = (payload: string): StreamEvent[] => {
      if (payload === DONE_PAYLOAD) {
        return [{ type: "finish", reason: "done" }];
      }
      let chunk: {
        choices?: Array<{
          delta?: {
            content?: string;
            reasoning_content?: string;
            reasoning?: string;
            tool_calls?: Array<{
              index?: number;
              id?: string;
              function?: { name?: string; arguments?: string };
            }>;
          };
          finish_reason?: string | null;
        }>;
      };
      try {
        chunk = JSON.parse(payload);
      } catch {
        return []; // keep-alive or malformed line
      }
      const choice = chunk.choices?.[0];
      if (!choice) {
        return [];
      }
      const events: StreamEvent[] = [];
      const delta = choice.delta ?? {};
      const reasoning = delta.reasoning_content ?? delta.reasoning;
      if (reasoning) {
        events.push({ type: "reasoning-delta", text: reasoning });
      }
      if (delta.content) {
        events.push({ type: "content-delta", text: delta.content });
      }
      if (delta.tool_calls) {
        for (const tc of delta.tool_calls) {
          const index = tc.index ?? 0;
          if (tc.id || tc.function?.name) {
            events.push({ type: "tool-call-start", index, id: tc.id ?? "", name: tc.function?.name ?? "" });
          }
          if (tc.function?.arguments) {
            events.push({ type: "tool-call-args", index, argsDelta: tc.function.arguments });
          }
        }
      }
      if (choice.finish_reason) {
        events.push({ type: "finish", reason: choice.finish_reason });
      }
      return events;
    };

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      buffer += decoder.decode(value, { stream: true });
      const parsed = extractSsePayloads(buffer);
      buffer = parsed.rest;
      for (const payload of parsed.payloads) {
        if (payload === DONE_PAYLOAD) {
          yield { type: "finish", reason: "done" };
          continue;
        }
        for (const ev of handlePayload(payload)) {
          yield ev;
        }
      }
    }
    // Flush any final line the server sent without a trailing newline.
    for (const payload of extractSsePayloads(buffer + "\n").payloads) {
      for (const ev of handlePayload(payload)) {
        yield ev;
      }
    }
  }
}
