import * as React from "react";
import {
  SettingsRegular,
  SendRegular,
  StopRegular,
  ArrowClockwiseRegular,
  ChatAddRegular,
} from "@fluentui/react-icons";
import { HttpTransport, normalizeOpenAiBase } from "../llm/transport";
import { runAgent, ToolCall } from "../llm/agent";
import { TOOL_SCHEMAS } from "../llm/tools";
import { excelExecutor } from "../llm/excelTools";
import { buildWorkbookSummary, buildSystemPrompt } from "../llm/context";

/* global localStorage, AbortController */

export type ServerType = "vllm" | "ollama" | "lmstudio" | "custom";

interface Settings {
  serverType: ServerType;
  /** Used only when serverType is custom. Must be reachable over HTTPS. */
  customUrl: string;
  apiKey: string;
  model: string;
  temperature: number;
  confirmWrites: boolean;
}

interface ToolRun {
  name: string;
  argsPreview: string;
  summary?: string;
  status: "running" | "ok" | "declined" | "error";
}

interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  reasoning?: string;
  tools?: ToolRun[];
  isError?: boolean;
}

const SETTINGS_KEY = "excellocal.settings.v1";
/** Turns of prior chat replayed to the model. Older turns are dropped so a long
 *  session never silently overruns a local model's context window. */
const MAX_HISTORY_MESSAGES = 20;

const DEFAULT_SETTINGS: Settings = {
  // Same-origin bridge served by `npm start` itself (see llm-forward.js).
  serverType: "vllm",
  customUrl: "https://localhost:4001/vllm",
  apiKey: "",
  model: "",
  temperature: 0.7,
  confirmWrites: true,
};

/** The pane-facing URL for the chosen server type (same-origin bridges by default). */
function resolveBaseUrl(s: Settings): string {
  switch (s.serverType) {
    case "vllm":
      return "/vllm";
    case "ollama":
      return "/ollama";
    case "lmstudio":
      return "/lmstudio";
    default:
      // HTTP custom endpoints (e.g. a vLLM box on your tailnet) cannot be
      // called directly from the HTTPS pane; route them through the bridge.
      const base = normalizeOpenAiBase(s.customUrl);
      return base.toLowerCase().startsWith("http://") ? "/bridge" : base;
  }
}

/** The upstream target for the dynamic bridge, or null when not bridging. */
function httpBridgeTarget(s: Settings): string | null {
  if (s.serverType !== "custom") {
    return null;
  }
  const base = normalizeOpenAiBase(s.customUrl);
  return base.toLowerCase().startsWith("http://") ? base : null;
}

function displayTarget(s: Settings): string {
  return httpBridgeTarget(s) ?? resolveBaseUrl(s);
}

function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) {
      const stored = JSON.parse(raw) as Partial<Settings> & { baseUrl?: string };
      if (stored.baseUrl && !stored.serverType) {
        // migrate pre-multi-server settings
        stored.serverType = "custom";
        stored.customUrl = stored.baseUrl;
        delete stored.baseUrl;
      }
      return { ...DEFAULT_SETTINGS, ...(stored as Partial<Settings>) };
    }
  } catch {
    /* corrupted settings: fall through to defaults */
  }
  return { ...DEFAULT_SETTINGS };
}

function previewArgs(args: string, max = 80): string {
  const flat = args.replace(/\s+/g, " ").trim();
  return flat.length > max ? flat.slice(0, max) + "…" : flat;
}

export default function App(): React.ReactElement {
  const [settings, setSettings] = React.useState<Settings>(loadSettings);
  const [showSettings, setShowSettings] = React.useState(false);
  const [models, setModels] = React.useState<string[]>([]);
  const [connectionError, setConnectionError] = React.useState<string | null>(null);
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);
  const [input, setInput] = React.useState("");
  const [streaming, setStreaming] = React.useState(false);
  const [pendingConfirm, setPendingConfirm] = React.useState<{ call: ToolCall; resolve: (ok: boolean) => void } | null>(null);
  const abortRef = React.useRef<AbortController | null>(null);
  // Set by "Always allow" so the rest of the CURRENT run stops prompting; the
  // settings update alone cannot do that, the callback is already bound.
  const skipConfirmRef = React.useRef(false);
  const listRef = React.useRef<HTMLDivElement | null>(null);
  const transportRef = React.useRef<HttpTransport | null>(null);
  if (!transportRef.current) {
    transportRef.current = new HttpTransport();
  }

  React.useEffect(() => {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  }, [settings]);

  React.useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight });
  }, [messages, pendingConfirm]);

  const authHeaders = (s: Settings): Record<string, string> =>
    s.apiKey ? { Authorization: "Bearer " + s.apiKey } : {};

  const refreshModels = React.useCallback(async (s: Settings) => {
    setConnectionError(null);

    try {
      const bridge = httpBridgeTarget(s);
      const res = await fetch(resolveBaseUrl(s) + "/v1/models", {
        headers: { ...authHeaders(s), ...(bridge ? { "X-Llm-Target": bridge } : {}) },
      });
      if (!res.ok) {
        // include the response body: bridge refusals and server errors explain themselves
        const detail = await res.text().catch(() => "");
        throw new Error("HTTP " + res.status + (detail ? " - " + detail.slice(0, 200) : ""));
      }
      const data = await res.json();
      const ids: string[] = (data.data ?? []).map((m: { id: string }) => m.id);
      setModels(ids);
      setSettings((prev) =>
        ids.length && !ids.includes(prev.model) ? { ...prev, model: ids[0] } : prev
      );
    } catch (e) {
      setModels([]);
      setConnectionError(
        "Cannot reach LLM server at " + displayTarget(s) + " (" + (e as Error).message + "). " +
          "Is the LLM server running? vLLM :8000, Ollama :11434, LM Studio :1234."
      );
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  React.useEffect(() => {
    void refreshModels(loadSettings());
  }, [refreshModels]);

  // Re-probe the server (debounced) whenever the URL changes.
  React.useEffect(() => {
    const timer = setTimeout(() => void refreshModels(settings), 600);
    return () => clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [settings.serverType, settings.customUrl, refreshModels]);

  const updateSettings = (patch: Partial<Settings>) => setSettings((prev) => ({ ...prev, ...patch }));

  const appendToLastAssistant = (patch: { content?: string; reasoning?: string }) => {
    setMessages((prev) => {
      const next = [...prev];
      const last = next[next.length - 1];
      if (last && last.role === "assistant") {
        next[next.length - 1] = {
          ...last,
          content: last.content + (patch.content ?? ""),
          reasoning: (last.reasoning ?? "") + (patch.reasoning ?? ""),
        };
      }
      return next;
    });
  };

  const addToolRun = (call: ToolCall) => {
    setMessages((prev) => {
      const next = [...prev];
      const last = next[next.length - 1];
      if (last && last.role === "assistant") {
        const run: ToolRun = { name: call.name, argsPreview: previewArgs(call.args), status: "running" };
        next[next.length - 1] = { ...last, tools: [...(last.tools ?? []), run] };
      }
      return next;
    });
  };

  const finishToolRun = (call: ToolCall, summary: string, status: ToolRun["status"]) => {
    setMessages((prev) => {
      const next = [...prev];
      for (let i = next.length - 1; i >= 0; i--) {
        const m = next[i];
        if (m.role !== "assistant" || !m.tools) {
          continue;
        }
        const idx = m.tools.map((t) => t.status === "running" && t.name === call.name).lastIndexOf(true);
        if (idx >= 0) {
          const tools = [...m.tools];
          tools[idx] = { ...tools[idx], status, summary };
          next[i] = { ...m, tools };
          break;
        }
      }
      return next;
    });
  };

  const sendMessage = async () => {
    const text = input.trim();
    if (!text || streaming || pendingConfirm) {
      return;
    }
    if (!settings.model) {
      setMessages((prev) => [
        ...prev,
        { role: "assistant", content: "No model selected. Open Settings and pick a model from your LLM server.", isError: true },
      ]);
      setShowSettings(true);
      return;
    }

    const userMsg: ChatMessage = { role: "user", content: text };
    const assistantMsg: ChatMessage = { role: "assistant", content: "" };
    const history: Array<{ role: "user" | "assistant"; content: string }> = messages
      .filter((m) => !m.isError && m.content.trim())
      .map((m) => ({ role: m.role, content: m.content }))
      .slice(-MAX_HISTORY_MESSAGES);

    setMessages((prev) => [...prev, userMsg, assistantMsg]);
    setInput("");
    setStreaming(true);

    const controller = new AbortController();
    abortRef.current = controller;
    skipConfirmRef.current = false;

    try {
      const workbookSummary = await buildWorkbookSummary();
      const result = await runAgent({
        transport: transportRef.current!,
        transportOptions: {
          baseUrl: resolveBaseUrl(settings),
          extraHeaders: httpBridgeTarget(settings) ? { "X-Llm-Target": httpBridgeTarget(settings)! } : undefined,
          apiKey: settings.apiKey || undefined,
          model: settings.model,
          temperature: settings.temperature,
          signal: controller.signal,
        },
        systemPrompt: buildSystemPrompt(workbookSummary, settings.confirmWrites),
        history,
        userMessage: text,
        tools: TOOL_SCHEMAS,
        executor: excelExecutor,
        callbacks: {
          onContentDelta: (t) => appendToLastAssistant({ content: t }),
          onReasoningDelta: (t) => appendToLastAssistant({ reasoning: t }),
          onToolStart: (call) => addToolRun(call),
          onToolEnd: (call, summary, ok, declined) =>
            finishToolRun(call, summary, declined ? "declined" : ok ? "ok" : "error"),
          ...(settings.confirmWrites
            ? {
                confirmTool: (call: ToolCall) =>
                  skipConfirmRef.current
                    ? Promise.resolve(true)
                    : new Promise<boolean>((resolve) => {
                        setPendingConfirm({ call, resolve });
                      }),
              }
            : {}),
        },
        signal: controller.signal,
      });
      if (result.limitReached) {
        appendToLastAssistant({
          content: "\n\n(Reached the tool-step limit for one request. Ask me to continue if there is more to do.)",
        });
      }
      if (result.truncated) {
        appendToLastAssistant({ content: "\n\n(The model hit its output-token limit. Ask it to continue.)" });
      }
      if (result.aborted) {
        // Always mark it: a half-finished answer otherwise reads as complete.
        appendToLastAssistant({ content: "\n\n(stopped)" });
      }
    } catch (e) {
      const err = e as Error;
      if (err.name !== "AbortError") {
        setMessages((prev) => {
          const next = [...prev];
          const failure: ChatMessage = { role: "assistant", content: "Request failed: " + err.message, isError: true };
          const last = next[next.length - 1];
          if (last && last.role === "assistant" && !last.content && !last.reasoning && !last.tools?.length) {
            next[next.length - 1] = failure;
          } else {
            next.push(failure);
          }
          return next;
        });
      }
    } finally {
      setStreaming(false);
      abortRef.current = null;
      skipConfirmRef.current = false;
      // A failed run must never leave the confirm overlay (and the composer)
      // stuck; resolve anything still waiting.
      setPendingConfirm((p) => {
        p?.resolve(false);
        return null;
      });
    }
  };

  const stopStreaming = () => {
    abortRef.current?.abort();
    if (pendingConfirm) {
      pendingConfirm.resolve(false);
      setPendingConfirm(null);
    }
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      void sendMessage();
    }
  };

  const newChat = () => {
    if (streaming || pendingConfirm) {
      return;
    }
    setMessages([]);
    setInput("");
  };

  const resolveConfirm = (ok: boolean) => {
    if (pendingConfirm) {
      pendingConfirm.resolve(ok);
      setPendingConfirm(null);
    }
  };

  return (
    <div className="app">
      <header className="app-header">
        <div className="app-title">
          <span className="app-logo" aria-hidden />
          <span>ExcelLocal</span>
        </div>
        <div className="app-header-actions">
          <button
            className="icon-btn"
            title="New chat"
            aria-label="New chat"
            onClick={newChat}
            disabled={streaming || !!pendingConfirm || messages.length === 0}
          >
            <ChatAddRegular />
          </button>
          <button
            className="icon-btn"
            title="Refresh models"
            aria-label="Refresh models"
            onClick={() => void refreshModels(settings)}
            disabled={streaming}
          >
            <ArrowClockwiseRegular />
          </button>
          <button
            className="icon-btn"
            title="Settings"
            aria-label="Settings"
            onClick={() => setShowSettings((v) => !v)}
          >
            <SettingsRegular />
          </button>
        </div>
      </header>

      {showSettings && (
        <div className="settings-panel">
          <label className="field">
            <span>LLM server</span>
            <select
              value={settings.serverType}
              onChange={(e) => updateSettings({ serverType: e.target.value as ServerType })}
            >
              <option value="vllm">vLLM (localhost:8000)</option>
              <option value="ollama">Ollama (localhost:11434)</option>
              <option value="lmstudio">LM Studio (localhost:1234)</option>
              <option value="custom">Custom URL…</option>
            </select>
          </label>
          {settings.serverType === "custom" && (
            <label className="field">
              <span>Endpoint URL — http:// addresses are bridged via the dev server (localhost / tailnet / LAN only)</span>
              <input
                type="text"
                value={settings.customUrl}
                onChange={(e) => updateSettings({ customUrl: e.target.value })}
                spellCheck={false}
              />
            </label>
          )}
          <label className="field">
            <span>Model</span>
            <select value={settings.model} onChange={(e) => updateSettings({ model: e.target.value })}>
              {models.length === 0 && !settings.model && <option value="">(none detected)</option>}
              {/* Keep the saved model selectable when the server probe failed,
                  otherwise the <select> silently falls back to the first entry. */}
              {settings.model && !models.includes(settings.model) && (
                <option value={settings.model}>{settings.model} (saved)</option>
              )}
              {models.map((m) => (
                <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
          </label>
          <div className="field-row">
            <label className="field">
              <span>API key (optional)</span>
              <input
                type="password"
                value={settings.apiKey}
                onChange={(e) => updateSettings({ apiKey: e.target.value })}
                spellCheck={false}
              />
            </label>
            <label className="field field-narrow">
              <span>Temp</span>
              <input
                type="number"
                step="0.1"
                min="0"
                max="2"
                value={settings.temperature}
                onChange={(e) => updateSettings({ temperature: parseFloat(e.target.value) || 0 })}
              />
            </label>
          </div>
          <label className="check-field">
            <input
              type="checkbox"
              checked={settings.confirmWrites}
              onChange={(e) => updateSettings({ confirmWrites: e.target.checked })}
            />
            <span>Ask before the model writes to the workbook</span>
          </label>
        </div>
      )}

      {connectionError && <div className="banner banner-error">{connectionError}</div>}

      <div className="messages" ref={listRef}>
        {messages.length === 0 && (
          <div className="empty-state">
            <div className="empty-title">Your local AI for Excel</div>
            <div className="empty-sub">
              Ask about your data, write formulas, format, chart. The model sees a snapshot of this
              workbook and can act on it with tools. Everything runs on your machine.
            </div>
          </div>
        )}
        {messages.map((m, i) => (
          <div key={i} className={"msg msg-" + m.role + (m.isError ? " msg-error" : "")}>
            {m.reasoning && (
              <details className="reasoning">
                <summary>Thinking</summary>
                <div className="reasoning-body">{m.reasoning}</div>
              </details>
            )}
            {m.tools && m.tools.length > 0 && (
              <div className="tool-runs">
                {m.tools.map((t, j) => (
                  <div key={j} className={"tool-run tool-" + t.status}>
                    <span className="tool-name">{t.name}</span>
                    <span className="tool-args">{t.argsPreview}</span>
                    {t.summary && <span className="tool-summary">{t.summary}</span>}
                  </div>
                ))}
              </div>
            )}
            <div className="msg-content">{m.content || (streaming && i === messages.length - 1 ? "…" : "")}</div>
          </div>
        ))}
      </div>

      {pendingConfirm && (
        <div className="confirm-overlay">
          <div className="confirm-card">
            <div className="confirm-title">The model wants to modify your workbook</div>
            <div className="confirm-tool">{pendingConfirm.call.name}</div>
            <pre className="confirm-args">{previewArgs(pendingConfirm.call.args, 400)}</pre>
            <div className="confirm-actions">
              <button className="btn-secondary" onClick={() => resolveConfirm(false)}>
                Deny
              </button>
              <button className="btn-secondary" onClick={() => resolveConfirm(true)}>
                Allow once
              </button>
              <button
                className="btn-primary"
                title="Turns off future confirmation prompts - re-enable in Settings"
                onClick={() => {
                  skipConfirmRef.current = true;
                  updateSettings({ confirmWrites: false });
                  resolveConfirm(true);
                }}
              >
                Always allow
              </button>
            </div>
            <div className="confirm-hint">Allow once approves just this call. Always allow stops future prompts (Settings re-enables them).</div>
          </div>
        </div>
      )}

      <footer className="composer">
        <textarea
          value={input}
          placeholder="Ask about your workbook… (Enter to send)"
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={onKeyDown}
          disabled={!!pendingConfirm}
          rows={2}
        />
        {streaming ? (
          <button className="send-btn stop" title="Stop" onClick={stopStreaming}>
            <StopRegular />
          </button>
        ) : (
          <button className="send-btn" title="Send" onClick={() => void sendMessage()} disabled={!input.trim()}>
            <SendRegular />
          </button>
        )}
      </footer>
    </div>
  );
}
