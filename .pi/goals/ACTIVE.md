# Active Goal

Status: COMPLETE
Goal: Support models over a tailnet via HTTP (user types the endpoint in the app, bridged safely); push all work to GitHub.
Started: 2026-08-28 20:20
Last updated: 2026-08-28 20:50

## Scope and acceptance criteria
- [x] A1: /bridge dynamic forwarding in dev server + standalone proxy: pane sends x-llm-target header; only private/tailnet hosts allowed (SSRF guard); https targets tolerate private certs
- [x] A2: App: custom http:// URLs are bridged automatically (no more refusal), settings hint shown; error messages show the real target
- [x] A3: Verified live: /bridge with allowed target returns models JSON; disallowed public target rejected 403; full regression green
- [x] A4: Docs: README + TUTORIAL tailnet/LAN section
- [x] A5: Pushed to GitHub (incl. previous uncommitted plug-and-play + sideload-evidence work); screenshot excluded from git

## Constraints and decisions
- Excel sideload already user-confirmed visually ("i see it loaded in"); keep Excel + dev server RUNNING (user is using it) - tests use dedicated ports (8123/4010) so no conflict.
- Bridge guardrail: allow localhost/127/10/172.16-31/192.168/100.64-127 (tailscale CGNAT)/.ts.net/.local only. Public hosts rejected 403.
- Bridge https targets use rejectUnauthorized:false (personal private certs).

## Todo
- [x] P301: llm-forward.js dynamic target mode | Verify: node --check
- [x] P302: webpack + standalone proxy /bridge routes | Verify: build
- [x] P303: transport.ts extraHeaders + App.tsx bridge logic | Verify: npm run build
- [x] P304: live guard tests (allowed + 403) + regression | Verify: curl outputs + npm run check
- [x] P305: docs | Verify: read
- [x] P306: push | Verify: git push

## Progress and evidence
- 20:50 ALL PASS. P301: llm-forward.js dynamicTarget mode - x-llm-target header, private/tailnet guard (localhost/127/10/172.16-31/192.168/100.64-127/*.ts.net/*.local), rejectUnauthorized:false for bridged https; guard unit-check 14/14. P302: /bridge in dev server + standalone proxy honors dynamic target. P303: transport extraHeaders; App.tsx bridges custom http:// automatically (8 substitutions), hint text, displayTarget in errors. P304: live smoke - allowed target models JSON, public 403, missing header 400, 6 SSE events through bridge; npm run check PASS; integration test made collision-immune via dynamic free ports after my own smoke squatted 8123 (environment, not code). P305: README + TUTORIAL tailnet sections; screenshot gitignored. P306: pushed.
- NOTE: user Excel session runs the pre-bridge dev server; restart npm start to activate /bridge (pane code hot-reloads already).