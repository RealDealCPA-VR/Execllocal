# ExcelLocal — a fully local "Claude for Excel" style add-in

An Excel task-pane add-in that chats with **your own locally hosted LLMs** (vLLM) and works
directly on your workbook with tools: reading data, writing values and formulas, formatting,
tables, sheets, and charts — the same agentic experience as cloud assistants, except nothing
ever leaves your machine.

## Architecture

```
+--------------------------------------------------------------+
|  Excel desktop (Microsoft 365)                                |
|  +--------------------------------------------------------+  |
|  |  Task pane (React + TS, served from https://localhost:3000)
|  |  chat UI -> agent loop -> Office.js tools              |  |
|  +-------------------------------+------------------------+  |
+----------------------------------|---------------------------+
                                   | HTTPS (dev certs)
                                   v
                    https://localhost:4001/vllm   (llm-proxy.js)
                                   | plain HTTP
                                   v
                    vLLM OpenAI server  http://localhost:8000
                    (your GPU: GLM-Flash / Qwen3 ...)
```

The proxy exists because Office serves task panes over HTTPS and browsers block calls from an
HTTPS page to an HTTP endpoint (mixed content). It forwards to vLLM using the Office dev
certificates already trusted by Excel.

## Prerequisites

- Node.js 18+
- Microsoft 365 Excel (desktop)
- A local LLM server. Tested with vLLM's OpenAI-compatible API; anything that implements
  `/v1/chat/completions` with **streaming and tool calling** works (Ollama, LM Studio, ...).

### vLLM flags for tool calling

Tool calling is a server feature; enable it when you launch vLLM, e.g.:

```
vllm serve <model> --enable-auto-tool-choice --tool-call-parser hermes   # Qwen3 family
vllm serve <model> --enable-auto-tool-choice --tool-call-parser glm45   # GLM-4.5/4.6 family
```

(Use the parser that matches your model — see `vllm serve --tool-call-parser=help`.)
Models appear in the pane's Settings dropdown automatically via `/v1/models`.

## Run it

```bash
npm install
```

Terminal A — the local HTTPS proxy in front of vLLM:

```bash
npm run proxy          # https://localhost:4001/vllm  ->  http://localhost:8000
```

Terminal B — build, serve the pane, and sideload into desktop Excel:

```bash
npm start
```

- The first run asks to trust "Developer CA for Microsoft Office Add-ins" — accept it once.
- Excel opens; go to the **Home** ribbon and click **ExcelLocal**.
- In the pane: Settings -> pick your model -> chat.

Stop with `npm run stop` (unloads the sideloaded add-in).

### No GPU? Dev against the mock server

```bash
node tools/mock-vllm.js        # fake vLLM on :8000 that streams a canned reply
```

Point the pane's Server URL at `https://localhost:4001/vllm` as usual.

## What the model can do (tools)

| Tool | Does |
|---|---|
| `get_workbook_info` | Sheets, used-range dimensions, sample rows, tables, selection |
| `get_selection` | Values/formulas of the current selection |
| `read_range` | Values + formulas from an A1 range (capped at 1500 cells) |
| `write_range` | Write a 2D array of values |
| `write_formulas` | Write formulas; returns the computed values for verification |
| `format_range` | Bold/italic, size, colors, number format, alignment, autofit |
| `create_sheet` / `delete_sheet` | Sheet management |
| `create_table` | Convert a range to a styled table |
| `create_chart` | column / bar / line / pie from a source range |

The pane injects a fresh workbook snapshot (sheets, headers, sample rows, selection) into every
turn, so the model reasons about real data instead of guessing. Write tools ask for confirmation
before touching cells (toggle in Settings). Tool activity is displayed inline in the chat.

## Development

- `npm run build` — production webpack build
- `npm test` — agent-loop tests (scripted transport + fake executor, no Excel needed)
- `npx office-addin-manifest validate manifest.xml`
- Layout: `src/taskpane/components/App.tsx` (UI), `src/taskpane/llm/` (transport, agent loop,
  tool schemas, Office.js executor, workbook context), `llm-proxy.js`, `tools/`

## Limitations (honest list)

- Pivot tables: creating/modifying pivot structures via Office.js is limited — build the source
  data and formulas instead.
- No VBA/macros — Office web add-ins cannot run VBA.
- Tool-calling quality depends on your model; use a tool-call parser matched to the model family.
- The workbook snapshot caps sample sizes; very wide sheets show the first rows only (the model
  can read more via tools).
- The first `npm start` requires accepting the dev certificate prompt.

## Privacy

Chat content goes: pane -> local proxy -> local vLLM -> back. No telemetry, no cloud, no
external requests beyond Office's own runtime.
