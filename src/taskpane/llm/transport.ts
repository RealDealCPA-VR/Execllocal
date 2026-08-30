/**
 * Streaming transport for OpenAI-compatible chat APIs (vLLM via the local
 * HTTPS proxy). Converts the wire format into structured StreamEvents,
 * including streamed tool_call fragments.
 */
/* global fetch, AbortSignal, TextDecoder */
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
  let u = url.trim().replace(/\/+$/, "");
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

interface DeltaChunk {
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
}

interface WholeResponse {
  choices?: Array<{
    message?: {
      content?: string | null;
      reasoning_content?: string | null;
      reasoning?: string | null;
      tool_calls?: Array<{ id?: string; function?: { name?: string; arguments?: string } }>;
    };
    finish_reason?: string | null;
  }>;
  error?: { message?: string } | string;
}

/** Turn one SSE payload (a streaming delta chunk) into events. */
function eventsFromDelta(payload: string): StreamEvent[] {
  let chunk: DeltaChunk;
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
        events.push({
          type: "tool-call-start",
          index,
          id: tc.id ?? "",
          name: tc.function?.name ?? "",
        });
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
}

/**
 * Fallback for servers that ignore `stream: true` and answer with one JSON
 * body. Without this the pane would silently show an empty reply.
 */
function eventsFromWholeResponse(text: string): StreamEvent[] {
  const trimmed = text.trim();
  if (!trimmed) {
    return [];
  }
  let body: WholeResponse;
  try {
    body = JSON.parse(trimmed);
  } catch {
    throw new Error(
      "LLM returned a response that is neither SSE nor JSON: " + trimmed.slice(0, 300)
    );
  }
  if (body.error) {
    const msg =
      typeof body.error === "string"
        ? body.error
        : (body.error.message ?? JSON.stringify(body.error));
    throw new Error("LLM error: " + msg);
  }
  const choice = body.choices?.[0];
  if (!choice?.message) {
    return [];
  }
  const events: StreamEvent[] = [];
  const reasoning = choice.message.reasoning_content ?? choice.message.reasoning;
  if (reasoning) {
    events.push({ type: "reasoning-delta", text: String(reasoning) });
  }
  if (choice.message.content) {
    events.push({ type: "content-delta", text: String(choice.message.content) });
  }
  (choice.message.tool_calls ?? []).forEach((tc, index) => {
    events.push({ type: "tool-call-start", index, id: tc.id ?? "", name: tc.function?.name ?? "" });
    if (tc.function?.arguments) {
      events.push({ type: "tool-call-args", index, argsDelta: tc.function.arguments });
    }
  });
  events.push({ type: "finish", reason: choice.finish_reason || "stop" });
  return events;
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

    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      throw new Error(
        "LLM request failed: HTTP " + res.status + (detail ? " - " + detail.slice(0, 300) : "")
      );
    }

    // No readable stream (non-streaming server, or a WebView without streaming
    // fetch): fall back to reading the whole body.
    if (!res.body || typeof res.body.getReader !== "function") {
      const text = await res.text();
      for (const ev of eventsFromWholeResponse(text)) {
        yield ev;
      }
      return;
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let sawSse = false;
    let raw = "";
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      const text = decoder.decode(value, { stream: true });
      raw += text;
      buffer += text;
      const parsed = extractSsePayloads(buffer);
      buffer = parsed.rest;
      for (const payload of parsed.payloads) {
        sawSse = true;
        if (payload === DONE_PAYLOAD) {
          yield { type: "finish", reason: "done" };
          continue;
        }
        for (const ev of eventsFromDelta(payload)) {
          yield ev;
        }
      }
    }
    // Flush any final line the server sent without a trailing newline.
    for (const payload of extractSsePayloads(buffer + "\n").payloads) {
      sawSse = true;
      if (payload === DONE_PAYLOAD) {
        yield { type: "finish", reason: "done" };
        continue;
      }
      for (const ev of eventsFromDelta(payload)) {
        yield ev;
      }
    }

    if (!sawSse) {
      // The server answered 200 but never spoke SSE: treat it as one JSON body.
      for (const ev of eventsFromWholeResponse(raw)) {
        yield ev;
      }
    }
  }
}
