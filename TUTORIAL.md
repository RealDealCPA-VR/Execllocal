# 🎓 ExcelLocal Tutorial — from zero to AI-driven spreadsheets

Everything in here runs on your machine. By the end you'll be chatting with your own
local LLM inside Excel and letting it safely edit your workbooks.

---

## Part 1 — What you need

| Thing | Why | Check |
|---|---|---|
| **Node.js 18+** | runs the pane dev server, proxy, tools | `node --version` |
| **Microsoft 365 Excel (desktop)** | where the add-in lives | open Excel 🎉 |
| **A local LLM server** | the brain. vLLM is the reference; Ollama/LM Studio/llama.cpp work if they support streaming + tool calling | `curl http://localhost:8000/v1/models` |
| **A tool-capable model** | Qwen3, GLM-4.5/4.6, Devstral, Llama-3.1+ … small models fumble tool calls | — |

> 💡 No GPU right now? Skip the brain entirely: `node tools/mock-vllm.js` gives you a fake
> LLM to explore the UI (and `MOCK_AGENT=1 node tools/mock-vllm.js` even simulates a
> two-turn tool conversation).

## Part 2 — Start your model with tool calling ON

This is the #1 thing people get wrong. vLLM does NOT parse tool calls unless you ask:

```bash
# Qwen3 family
vllm serve Qwen/Qwen3-32B --enable-auto-tool-choice --tool-call-parser hermes

# GLM-4.5 / 4.6 family
vllm serve zai-org/glm-4.6 --enable-auto-tool-choice --tool-call-parser glm45
```

Not sure which parser? `vllm serve --tool-call-parser=help` lists them all.

**Symptom of getting it wrong:** the chat works, but the model answers in prose like
"I would call read_range…" instead of actually doing things.

## Part 3 — Launch the add-in

**Get the code first** (one time):

```bash
git clone https://github.com/RealDealCPA-VR/Execllocal.git
cd Execllocal
```

> No git? Use the green **Code** button on the repo page → Download ZIP → unzip.
> Windows may flag the `start-excellocal.cmd` launcher with SmartScreen — choose *More info → Run anyway*; it only runs npm.

Then, in that folder, open a terminal:
Open the project folder in a terminal:

```bash
# everything in one command: dev server + built-in LLM bridges + sideload
npm start
```

Windows shortcut: double-click `start-excellocal.cmd` (installs + starts).
- **Keep this terminal open** — it serves the add-in pane while you use ExcelLocal. Done? `npm run stop`, then close it.

The dev server has **built-in same-origin bridges**, so the pane needs zero config:
`/vllm` forwards to vLLM on :8000, `/ollama` to :11434, `/lmstudio` to :1234.


- First launch: Windows asks to trust **"Developer CA for Microsoft Office Add-ins"** — accept once.
- Excel opens automatically. Go to **Home → ExcelLocal** and click the button.
- The pane opens on the right. If Excel opened without the pane, use **Insert → My Add-ins →
  Developer Add-ins → ExcelLocal**.

Done experimenting? `npm run stop` unloads the add-in.

## Part 4 — The 30-second tour of the pane

```
+------------------------------------------+
| ExcelLocal            [refresh] [gear]   |  <- gear = settings
|------------------------------------------|
|  (messages appear here)                  |
|  ┌ tool chips show what the model did ┐  |
|  │ read_range  Sheet1!A1:D5   ✓       │  |
|  └────────────────────────────────────┘   |
|  ▸ Thinking            (collapsed)       |
|------------------------------------------|
| Ask about your workbook…        [Send]   |
+------------------------------------------+
```

- **Tool chips** — every tool the model ran, with a ✓ / declined / error badge.
- **Thinking** — reasoning models (GLM, Qwen3) stream their internal monologue here, collapsed by default.
- **Stop** — appears while streaming; aborts generation (the GPU stops too — the proxy kills the request).
- **Allow / Cancel dialog** — appears before any write; Cancel tells the model "user declined", and it will not retry silently.

## Part 5 — Configure once

Click the **gear**:

| Field | What to put |
|---|---|
| **LLM server** | Pick vLLM (localhost:8000), Ollama (11434) or LM Studio (1234) from the dropdown — zero typing. Models on another machine (tailnet/LAN)? Choose *Custom URL...* and type its **http://** address — it is bridged automatically. Public addresses are refused by design. |
| **Model** | auto-populated from your server's `/v1/models` — pick one |
| **API key** | leave empty for local servers |
| **Temp** | 0.2–0.7 for spreadsheet work (lower = more precise tool calls) |
| **Ask before the model writes** | keep ✔ on until you trust your model |

Settings persist between sessions. The model list auto-refreshes when you change the URL.

## Part 6 — Your first five prompts (copy-paste)

Start easy, then escalate:

1. **Reconnaissance** (read-only, no confirmations):
   > What's in this workbook? Summarize each sheet.

2. **Targeted read:**
   > What are the column headers on Sheet1 and how many rows of data?

3. **First write** (watch the Allow/Cancel dialog appear):
   > Add a TOTALS row directly under the data with SUM formulas for each numeric column, and bold it.

4. **Formatting:**
   > Format row 1 as a header: bold, white text on a dark green background.

5. **Analysis + visualization:**
   > Which category in column A has the highest total in column B? Then create a column chart of totals by category starting at cell E2.

Notice the pattern: the model **reads before it writes**, and **verifies after it writes** —
the tool chips let you audit every step.

## Part 7 — Living with the confirm dialog

- **Allow** → the write executes exactly as shown (the dialog shows the raw tool arguments).
- **Cancel** → the model receives `{"declined": true}` and is instructed not to retry without asking you.
- Brave? Uncheck *Ask before the model writes* in Settings. The tool chips still log everything.

## Part 8 — Troubleshooting

| Symptom | Fix |
|---|---|
| Red banner *"Cannot reach LLM server..."* | Is the LLM server up (`curl http://localhost:8000/v1/models`)? Right server type picked in Settings? |
| Chat works but model never *does* anything | vLLM missing `--enable-auto-tool-choice --tool-call-parser <family>` (Part 2) |
| Model makes tool calls that error instantly | Wrong parser for the model family; or the model is too small for tool use |
| `npm start` says port 3000 is busy | Close whatever owns the port (a reboot works) or run `npm run stop` first |
| A dialog asks to *debug the webview* / attach VS Code | You launched with debugging on. Use `npm start` (dialog-free) instead of `npm run start:debug`; if it keeps appearing: `reg add HKCUSOFTWAREMicrosoftOffice.0WEFDeveloper<add-in-id> /v UseDirectDebugger /t REG_DWORD /d 0 /f` |
| Pane is blank / doesn't load | `npm run stop`, then `npm start` again |
| *"Reached the tool-step limit"* | Big task — reply "continue" and it picks up where it stopped |
| Huge sheet, model seems blind | Snapshot shows the first rows only; ask it to `read_range` specific columns |
| HTTPS warning on first start | Accept the dev-certificate dialog; it's a one-time local cert |
| Buttons greyed out | A confirmation dialog is open behind the composer — answer it |

## Part 9 — For developers

```bash
npm test                # agent-loop unit tests (scripted transport, no Excel/GPU)
npm run test:integration# full stack over real sockets: transport → proxy → mock LLM
npm run check           # build + tests + manifest validation
npm run build           # production webpack build
```

- Hot reload: `npm start` keeps webpack watching; edit `src/taskpane/**` and the pane refreshes.
- `npm run start:debug` — same launch but with WebView debugging enabled (Office will offer to attach VS Code).
- `npm run proxy` is optional now — only for reaching a Custom HTTP endpoint (standalone HTTPS bridge on :4001).
- Simulate tool conversations without a GPU: `MOCK_AGENT=1 node tools/mock-vllm.js`.
- The agent loop (`src/taskpane/llm/agent.ts`) is pure TypeScript — extend `tools.ts` +
  `excelTools.ts` to add new superpowers.

## Part 10 — The mental model (remember just this)

> **The model can see a snapshot of your workbook and can call tools that read and write it.
> Reads happen freely; writes ask first; every action is displayed. Your data never leaves
> your machine.**

That's it. Go make your GPU earn its electricity bill. ⚡
