import * as React from "react";
import { SettingsRegular, SendRegular, StopRegular, ArrowClockwiseRegular } from "@fluentui/react-icons";
import { HttpTransport } from "../llm/transport";
import { runAgent, ToolCall } from "../llm/agent";
import { TOOL_SCHEMAS } from "../llm/tools";
import { excelExecutor } from "../llm/excelTools";
import { buildWorkbookSummary, buildSystemPrompt } from "../llm/context";

/* global localStorage, AbortController */

interface Settings {
  baseUrl: string;
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

const DEFAULT_SETTINGS: Settings = {
  // Local HTTPS proxy in front of vLLM (see llm-proxy.js).
  baseUrl: "https://localhost:4001/vllm",
  apiKey: "",
  model: "",
  temperature: 0.7,
  confirmWrites: true,
};

function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) {
      return { ...DEFAULT_SETTINGS, ...(JSON.parse(raw) as Partial<Settings>) };
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
      const res = await fetch(s.baseUrl + "/v1/models", { headers: authHeaders(s) });
      if (!res.ok) {
        throw new Error("HTTP " + res.status);
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
        "Cannot reach LLM server at " + s.baseUrl + " (" + (e as Error).message + "). " +
          "Is vLLM running, and did you start the proxy (npm run proxy)?"
      );
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  React.useEffect(() => {
    void refreshModels(loadSettings());
  }, [refreshModels]);

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
      .map((m) => ({ role: m.role, content: m.content }));

    setMessages((prev) => [...prev, userMsg, assistantMsg]);
    setInput("");
    setStreaming(true);

    const controller = new AbortController();
    abortRef.current = controller;

    try {
      const workbookSummary = await buildWorkbookSummary();
      const result = await runAgent({
        transport: transportRef.current!,
        transportOptions: {
          baseUrl: settings.baseUrl,
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
          confirmTool: (call) =>
            new Promise<boolean>((resolve) => {
              setPendingConfirm({ call, resolve });
            }),
        },
        signal: controller.signal,
      });
      if (result.limitReached) {
        appendToLastAssistant({ content: "\n\nn(Reached the tool-step limit for one request. Ask me to continue if there is more to do.)" });
      }
      if (result.aborted) {
        appendToLastAssistant({ content: (result.content ? "" : "\n\n(stopped)") });
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
            title="Refresh models"
            onClick={() => void refreshModels(settings)}
            disabled={streaming}
          >
            <ArrowClockwiseRegular />
          </button>
          <button className="icon-btn" title="Settings" onClick={() => setShowSettings((v) => !v)}>
            <SettingsRegular />
          </button>
        </div>
      </header>

      {showSettings && (
        <div className="settings-panel">
          <label className="field">
            <span>Server URL</span>
            <input
              type="text"
              value={settings.baseUrl}
              onChange={(e) => updateSettings({ baseUrl: e.target.value })}
              spellCheck={false}
            />
          </label>
          <label className="field">
            <span>Model</span>
            <select value={settings.model} onChange={(e) => updateSettings({ model: e.target.value })}>
              {models.length === 0 && <option value="">(none detected)</option>}
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
                Cancel
              </button>
              <button className="btn-primary" onClick={() => resolveConfirm(true)}>
                Allow
              </button>
            </div>
          </div>
        </div>
      )}

      <footer className="composer">
        <textarea
          value={input}
          placeholder="Ask about your workbook… (Enter to send)"
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={onKeyDown}
          disabled={streaming || !!pendingConfirm}
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
