<div align="center">

# 🔥 EXCELLOCAL 🔥

### Your spreadsheet just grew a local AI brain. **No cloud. No API keys. No mercy.**

*The Claude-for-Excel experience — powered by **YOUR** GPU, running **YOUR** models, touching **ONLY** your data.*

[![Fully Local](https://img.shields.io/badge/cloud-%F0%9F%9A%AB%20none-107c41?style=for-the-badge)]()
[![vLLM](https://img.shields.io/badge/powered%20by-vLLM-ff6f00?style=for-the-badge)]()
[![Excel](https://img.shields.io/badge/lives%20in-Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel)]()
[![Agent](https://img.shields.io/badge/tool%20calls-10%20workbook%20superpowers-blue?style=for-the-badge)]()

</div>

---

## 😤 The pitch

*New here? Read the step-by-step [TUTORIAL.md](TUTORIAL.md) — zero to AI-driven spreadsheets in ten short parts.*

Cloud copilots read your spreadsheets, ship your data to someone else's datacenter, and charge
you monthly for the privilege. **ExcelLocal says no.**

It's a task-pane add-in that drops a full agentic chat sidebar into Excel — the same kind of
experience Claude for Excel delivers — except the brain is **your local vLLM server**. GLM, Qwen3,
whatever you serve: if it runs on your GPU and speaks OpenAI, it drives your workbook.

```
You:   "Sum column B, add a TOTALS row, make it bold, then chart it"

ExcelLocal's model:  🧠 *reads the sheet* → 🔧 read_range("B1:B48")
                     → 🔧 write_formulas("=SUM(B2:B48)")
                     → ✅ verifies the result by reading it back
                     → 🔧 format_range(bold)
                     → 🔧 create_chart(column)
You:   "🤯 and it never left my machine"
```

## ⚡ Feature flex

| 💪 | Capability |
|---|---|
| 🪟 | **Real task pane** — sideloads into desktop Excel, lives on the Home ribbon |
| 🧠 | **True agent loop** — streaming chat, reasoning visible, multi-step tool chains |
| 📊 | **10 workbook tools** — read, write, formulas, format, tables, sheets, charts |
| 🛡️ | **Confirm-before-write** — the model asks before it touches your cells (toggleable) |
| 👁️ | **Selection awareness** — fresh workbook snapshot (sheets, headers, samples, selection) injected every turn |
| 🎥 | **Streaming everything** — SSE deltas, collapsible "Thinking" for reasoning models (GLM/Qwen3) |
| 🔒 | **Zero egress** — pane → built-in bridge → your GPU → back. That's the whole journey |
| 🧪 | **Tested** — agent-loop suite runs with a scripted transport, no Excel or GPU needed |
| ♻️ | **GPU-optional dev** — mock vLLM server included so you can hack the UI anywhere |

## 🗺️ Architecture (beautifully paranoid)

```text
+------------------------------------------------------------------+
|  Excel desktop (Microsoft 365)                                   |
|  +-----------------------------------------------------------+  |
|  | Task pane  https://localhost:3000/taskpane.html            |  |
|  |   chat UI -> agent loop -> Office.js tools                 |  |
|  +------------------------------+-----------------------------+  |
+---------------------------------|--------------------------------+
                                  |  HTTPS (Office dev certs)
                                  v
  npm start dev server  https://localhost:3000
    same-origin bridges (built in):
    /vllm -> :8000   /ollama -> :11434   /lmstudio -> :1234
    /bridge -> any private/tailnet target you type as Custom URL
                                  |  plain HTTP (localhost / LAN / tailnet)
                                  v
  your LLM server  (vLLM, Ollama, LM Studio... anything OpenAI-shaped)
  (your GPU: GLM-Flash, Qwen3-27B, ...)
```

Why a bridge at all? Office serves task panes over HTTPS, and browsers refuse to let an
HTTPS page call plain HTTP (mixed content). `npm start` therefore embeds the bridge itself:
the pane talks same-origin to `/vllm` (etc.), and the dev server forwards to your plain-HTTP
LLM server using the Office dev certificates Excel already trusts. That's the whole trick.
Running an LLM on another box (tailnet/LAN)? Type its http:// address as a Custom URL and the
`/bridge` route forwards to it — private ranges only, never the public internet.
`npm run proxy` remains as an optional standalone bridge (https://localhost:4001/vllm) for
anything exotic.

## 🚀 90-second quickstart

**You need:** Node.js 18+ · Microsoft 365 Excel (desktop) · a local LLM server (vLLM, Ollama, LM Studio, llama.cpp — anything OpenAI-shaped).

**0. Get the code** (and Node):

```bash
git clone https://github.com/RealDealCPA-VR/Execllocal.git
cd Execllocal
```

> No git? Use the green **Code** button on the repo page → Download ZIP → unzip it.
> No Node? Grab the LTS from [nodejs.org](https://nodejs.org).

**2. Serve a brain** — vLLM with tool calling switched ON:

```bash
vllm serve <model> --enable-auto-tool-choice --tool-call-parser hermes    # Qwen3 family
vllm serve <model> --enable-auto-tool-choice --tool-call-parser glm45    # GLM-4.5/4.6 family
```

> Use the parser that matches your model (`vllm serve --tool-call-parser=help`).

**3. Launch — one terminal, everything built in:**

```bash
npm install   # first run only
npm start     # builds, serves the pane, bridges /vllm /ollama /lmstudio, sideloads into Excel
```

- Windows shortcut: double-click `start-excellocal.cmd` (installs + starts).
- **Keep this terminal open** — it is serving the add-in pane for as long as you use ExcelLocal.
- First launch asks you to trust *"Developer CA for Microsoft Office Add-ins"* — accept once.
- Done for the day? `npm run stop` unloads the add-in, then close the terminal.

**4. In Excel:** **Home** ribbon → **ExcelLocal** → gear icon → pick your **LLM server** (vLLM / Ollama / LM Studio / Custom URL) and **model** from the auto-populated list → chat.

> Serving models on another box (Tailscale/LAN, plain HTTP)? Pick **Custom URL** and type its `http://` address — see the next section.

### 🌐 Models on another machine (Tailscale/LAN, plain HTTP)

Serving vLLM on your tailnet or LAN at, say, `http://my-gpu-box:8000`? Choose **Custom URL** in Settings and type exactly that. The dev server bridges it automatically (same-origin), restricted to private ranges: localhost, 10.x, 172.16-31.x, 192.168.x, 100.64.0.0/10 (Tailscale), `*.ts.net`, `*.local`. Public addresses are refused — the bridge is your private tunnel, not an open proxy.

## 🖥️ Headless mode — run the server on the GPU box instead

Tired of the terminal on your Excel machine? Serve everything from the remote box:

```bash
# on the GPU box (one time): clone, npm ci, npm run build
node server/serve.js                     # serves dist/ + bridges on :3000
tailscale serve --bg --https=443 http://127.0.0.1:3000   # trusted HTTPS via your tailnet
```

# on Windows (one time): generate + register the remote manifest
npm run manifest:remote -- https://your-box.your-tailnet.ts.net
npm run sideload:remote

Excel now loads the pane straight from your GPU box over your tailnet. No webpack,
no local terminal, ever. The pane's Server type stays **vLLM (localhost:8000)** — the
bridge on the GPU box forwards to its own local vLLM automatically.

**No GPU at hand?** `node tools/mock-vllm.js` fakes a streaming LLM on :8000 so you can
develop the UI on a toaster. `MOCK_AGENT=1` makes it simulate a two-turn tool conversation.

Stop everything with `npm run stop`.

## 🧰 The tool belt

| Tool | What the model does with it |
|---|---|
| `get_workbook_info` | X-rays the workbook: sheets, sizes, sample rows, tables, your selection |
| `get_selection` | Reads exactly what you've highlighted |
| `read_range` | Values + formulas from any A1 range (1500-cell guard rail) |
| `write_range` | Writes 2D arrays of values (auto-reshapes to fit) |
| `write_formulas` | Writes formulas **and reads back computed values** to self-verify |
| `format_range` | Bold, italic, colors, number formats, alignment, autofit |
| `create_sheet` / `delete_sheet` | Sheet lifecycle (deletes ask first, obviously) |
| `create_table` | Range → styled, filterable Excel table |
| `create_chart` | column / bar / line / pie, placed where you want it |

## 🧠 How the agent thinks

Every turn:

1. **Snapshot** — the pane builds a fresh workbook summary (sheets, dimensions, first rows,
   tables, selection) and injects it into the system prompt. The model sees real data, never guesses.
2. **Decide** — the model streams a reply or emits tool calls (native OpenAI function calling via vLLM).
3. **Act** — write tools pause for your **Allow / Cancel**; read tools just run.
4. **Verify** — after writing, the model reads the range back before claiming victory.
5. **Repeat** — results go back to the model until it has a final answer.

The loop (`src/taskpane/llm/agent.ts`) is pure TypeScript — no Office.js, no DOM — so it's
unit-tested in Node with a scripted transport and a fake executor:

```bash
npm test      # 4 scenarios: tool round-trips, declined writes, runaway guard, abort
npm run check # build + tests + manifest validation, one command
```

## 📁 Project map

```
src/taskpane/components/App.tsx   chat UI (streaming, tool chips, confirm dialog)
src/taskpane/llm/transport.ts     OpenAI-compatible SSE streaming client
src/taskpane/llm/agent.ts         the agent loop (pure, unit-tested)
src/taskpane/llm/tools.ts         tool schemas exposed to the model
src/taskpane/llm/excelTools.ts    Office.js tool implementations
src/taskpane/llm/context.ts       workbook snapshot builder + system prompt
llm-forward.js                    shared forwarding logic (dev-server bridges + standalone proxy)
llm-proxy.js                       optional standalone HTTPS bridge (:4001) for exotic setups
tools/mock-vllm.js                fake vLLM for GPU-less dev
manifest.xml                      Office add-in manifest (validated)
```

## ⚠️ Fine print (we don't do vaporware claims)

- **Pivot tables** — Office.js can't build/reshape pivots deeply. Ask for source data + formulas instead.
- **No VBA** — web add-ins can't run macros. That's an Office rule, not our laziness.
- **Tool-call quality = model quality** — small models will fumble tool syntax; use a
  tool-call parser matched to your model family and a model that can actually follow instructions.
- **Big sheets** — snapshots cap sample sizes; the model can read more through tools, range by range.
- **Port 3000 busy** — `npm start` needs it for the pane; close whatever owns it (a reboot works) or run `npm run stop` first.
- **Office.js CDN** — the one standard exception to "fully local": the pane boots Microsoft's Office.js from their CDN (cached after first load). Your prompts, workbook data, and model traffic never leave your machine.
- **First `npm start`** — Windows asks once to trust the Office dev certificate. Say yes.
- **No debug dialog** — `npm start` runs without WebView debugging, so Office never asks to attach VS Code. Want to debug the pane? `npm run start:debug`.

## 🕊️ Privacy statement

```
your data  ->  your pane  ->  your bridge  ->  your GPU  ->  back
```

No telemetry. No accounts. No "your files help improve our services" fine print.
If the network light blinks, it's Office itself, not us.

---

<div align="center">

**Built for people who read their spreadsheet's privacy policy.**

⭐ the repo if your GPU finally earns its electricity bill

</div>
