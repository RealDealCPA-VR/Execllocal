# Active Goal

Status: COMPLETE
Goal: Make ExcelLocal plug-and-play: user only picks server type (vLLM/Ollama/LM Studio/Custom) + model in the app; same-origin LLM bridges inside `npm start` (no second terminal); Windows launcher; docs updated; push.
Started: 2026-08-28 18:55
Last updated: 2026-08-28 19:20

## Scope and acceptance criteria
- [x] A1: llm-forward.js shared handler; llm-proxy.js refactored onto it (standalone behavior unchanged)
- [x] A2: webpack dev-server middleware: /vllm -> VLLM_URL(:8000), /ollama -> :11434, /lmstudio -> :1234 (same-origin, no CORS)
- [x] A3: App.tsx: server-type picker + custom URL; legacy baseUrl migration; all call sites use resolveBaseUrl
- [x] A4: regression green: npm run check + integration test (standalone proxy scenario still passes)
- [x] A5: dev-server middleware smoke: real curl https://localhost:3000/vllm/v1/models returns models JSON
- [x] A6: start-excellocal.cmd launcher (install-if-needed + start)
- [x] A7: README + TUTORIAL rewritten to one-terminal flow; push to origin

## Constraints and decisions
- Same-origin bridge removes CORS entirely for the default flows; standalone HTTPS proxy stays for custom/advanced setups.
- Express mount strips the prefix (use stripPrefix:false for middleware); standalone proxy keeps manual /vllm strip.
- Known hazards: no backslash literals in patches (char codes only); verify every write.

## Todo
- [x] P101: llm-forward.js + llm-proxy refactor | Verify: node --check both
- [x] P102: webpack middleware | Verify: build + dev-server curl smoke
- [x] P103: App.tsx settings rework | Verify: npm run build +
 tests + integration
- [x] P104: start-excellocal.cmd launcher | Verify: exists, content sane
- [x] P105: docs rewrite (README quickstart, TUTORIAL parts 3/5/8) | Verify: manual read
- [x] P106: final regression + push | Verify: npm run check + test:integration + git push

## Progress and evidence
- 19:20 ALL PASS. P101: llm-forward.js shared handler; llm-proxy.js refactored (node --check OK). P102: webpack setupMiddlewares - same-origin bridges /vllm,/ollama,/lmstudio; live smoke via real dev server: models JSON + 6 SSE data events + clean 502 JSON for down upstream + taskpane.html served. P103: App.tsx 9 substitutions - ServerType picker (vLLM/Ollama/LM Studio/Custom), resolveBaseUrl, legacy baseUrl migration, all call sites. P104: start-excellocal.cmd. P105: README quickstart one-terminal + TUTORIAL parts 3/5/8 (line-index patches to survive multibyte). P106: npm run check green; integration test collision-proofed (dedicated ports 8123/4010 after a stale plain mock on :8000 caused a false failure) - PASS; LICENSE (MIT) + GitHub Actions CI added for production readiness. Pushed.
- FINDINGS: integration false failure root-caused to environment (stale server), not code; test now immune.